from __future__ import annotations

import re
import shutil
import threading
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from assessment_runner_core import (
    assess_controls,
    assessment_mode_label,
    build_findings,
    load_questionnaire_chunks,
    normalize_assessment_mode,
    read_template_controls,
    summarize_sections,
    write_assessment_workbook,
    write_filled_template_workbook,
)
from gemini_secret_store import resolve_api_key_with_source


ROOT_DIR = Path(__file__).resolve().parent
JOB_ROOT = ROOT_DIR / "gui_jobs"
JOB_ROOT.mkdir(parents=True, exist_ok=True)


@dataclass
class JobState:
    id: str
    status: str = "queued"
    stage: str = "Queued"
    progress: float = 0.0
    message: str = "Waiting to start"
    company: str = ""
    assessment_mode: str = "questionnaire"
    created_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    error: str = ""
    result: Dict[str, Any] = field(default_factory=dict)


JOBS: Dict[str, JobState] = {}
JOBS_LOCK = threading.Lock()


app = FastAPI(title="ISO Assessment GUI API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:5174",
        "http://127.0.0.1:5174",
        "http://localhost:5175",
        "http://127.0.0.1:5175",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def update_job(job_id: str, **changes: Any) -> None:
    with JOBS_LOCK:
        job = JOBS[job_id]
        for key, value in changes.items():
            setattr(job, key, value)
        job.updated_at = utc_now_iso()


def sanitize_user_message(message: str) -> str:
    message = re.sub(r"[A-Za-z]:\\[^\r\n\"']+", "[local file]", str(message))
    message = re.sub(r"/(?:[^/\s]+/)+[^/\s]+", "[local file]", message)
    return message


def choose_template_path() -> Path:
    candidates = [
        ROOT_DIR / "master_iso_template.xlsx",
        ROOT_DIR / "security_assessment_template.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        "Could not find a template workbook. Expected master_iso_template.xlsx or security_assessment_template.xlsx in the project root."
    )


def run_assessment_job(
    job_id: str,
    company: str,
    model_name: str,
    questionnaire_paths: List[str],
    assessment_mode: str,
) -> None:
    try:
        assessment_mode = normalize_assessment_mode(assessment_mode)
        update_job(job_id, status="running", stage="Preparing controls", progress=5.0, message="Loading template and controls")

        template_path = choose_template_path()
        controls = read_template_controls(str(template_path))

        update_job(job_id, stage="Loading questionnaires", progress=15.0, message="Parsing submitted evidence files")
        chunks = load_questionnaire_chunks(questionnaire_paths)

        api_key, _key_source = resolve_api_key_with_source()
        if not api_key:
            raise RuntimeError(
                "No Gemini API key found. Set GEMINI_API_KEY/GOOGLE_API_KEY or run py -3.12 manage_gemini_key.py --set."
            )

        update_job(
            job_id,
            stage="Analyzing controls",
            progress=25.0,
            message=f"{assessment_mode_label(assessment_mode)} is reviewing the submitted evidence.",
        )

        def progress_callback(start_index: int, end_index: int, total_controls: int) -> None:
            completed_controls = end_index
            ratio = completed_controls / total_controls if total_controls else 0.0
            progress = 25.0 + (ratio * 55.0)
            update_job(
                job_id,
                stage="Analyzing controls",
                progress=min(progress, 80.0),
                message=f"Assessed {completed_controls} of {total_controls} controls",
            )

        evaluations = assess_controls(
            api_key=api_key,
            model_name=model_name,
            company_name=company,
            controls=controls,
            chunks=chunks,
            batch_size=16,
            max_chunks=80,
            top_chunks_per_control=4,
            max_evidence_chars=50000,
            assessment_mode=assessment_mode,
            max_parallel_batches=4,
            progress_callback=progress_callback,
        )

        job_dir = JOB_ROOT / job_id
        output_assessment = job_dir / "assessment_report.xlsx"
        output_filled_template = job_dir / "security_assessment_template_filled.xlsx"

        update_job(job_id, stage="Compiling workbook", progress=88.0, message="Writing filled template workbook")
        write_filled_template_workbook(
            controls=controls,
            evaluations=evaluations,
            company_name=company,
            output_path=str(output_filled_template),
            assessment_mode=assessment_mode,
        )

        update_job(job_id, stage="Compiling workbook", progress=94.0, message="Writing executive assessment report")
        metrics = write_assessment_workbook(
            controls=controls,
            evaluations=evaluations,
            chunks=chunks,
            company_name=company,
            model_name=model_name,
            output_path=str(output_assessment),
            assessment_mode=assessment_mode,
        )

        section_summary = summarize_sections(controls, evaluations)
        findings = build_findings(controls, evaluations)

        result = {
            "company": company,
            "model": model_name,
            "assessment_mode": assessment_mode,
            "assessment_mode_label": assessment_mode_label(assessment_mode),
            "metrics": metrics,
            "section_summary": section_summary[:8],
            "top_findings": findings[:12],
            "files": {
                "assessment_report": output_assessment.name,
                "filled_template": output_filled_template.name,
            },
        }

        update_job(
            job_id,
            status="completed",
            stage="Completed",
            progress=100.0,
            message="Assessment complete. Results are ready for export.",
            result=result,
        )
    except Exception as exc:
        update_job(
            job_id,
            status="failed",
            stage="Failed",
            progress=100.0,
            message="Assessment failed",
            error=sanitize_user_message(str(exc)),
        )


@app.get("/api/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/api/assess")
async def create_assessment(
    company: str = Form(...),
    questionnaires: List[UploadFile] = File(...),
    model: str = Form("models/gemini-2.5-pro"),
    mode: str = Form("questionnaire"),
) -> Dict[str, Any]:
    company = company.strip()
    mode = normalize_assessment_mode(mode)
    if not company:
        raise HTTPException(status_code=400, detail="Company name is required.")

    if not questionnaires:
        raise HTTPException(status_code=400, detail="At least one questionnaire file is required.")

    job_id = str(uuid.uuid4())
    job_dir = JOB_ROOT / job_id
    input_dir = job_dir / "inputs"
    input_dir.mkdir(parents=True, exist_ok=True)

    saved_paths: List[str] = []
    for upload in questionnaires:
        file_name = Path(upload.filename or "uploaded_file").name
        target = input_dir / file_name
        with target.open("wb") as handle:
            shutil.copyfileobj(upload.file, handle)
        saved_paths.append(str(target))

    with JOBS_LOCK:
        JOBS[job_id] = JobState(id=job_id, company=company, assessment_mode=mode)

    worker = threading.Thread(
        target=run_assessment_job,
        args=(job_id, company, model, saved_paths, mode),
        daemon=True,
    )
    worker.start()

    return {"job_id": job_id}


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str) -> Dict[str, Any]:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            raise HTTPException(status_code=404, detail="Job not found")
        payload = {
            "id": job.id,
            "status": job.status,
            "stage": job.stage,
            "progress": round(job.progress, 2),
            "message": job.message,
            "company": job.company,
            "assessment_mode": job.assessment_mode,
            "assessment_mode_label": assessment_mode_label(job.assessment_mode),
            "created_at": job.created_at,
            "updated_at": job.updated_at,
            "error": job.error,
            "result": job.result,
        }

    if payload["status"] == "completed":
        payload["downloads"] = {
            "assessment_report": f"/api/jobs/{job_id}/download/assessment",
            "filled_template": f"/api/jobs/{job_id}/download/template",
        }

    return payload


@app.get("/api/jobs/{job_id}/download/assessment")
def download_assessment(job_id: str) -> FileResponse:
    file_path = JOB_ROOT / job_id / "assessment_report.xlsx"
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Assessment report not found")
    return FileResponse(
        path=file_path,
        filename=f"{job_id}_assessment_report.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/jobs/{job_id}/download/template")
def download_template(job_id: str) -> FileResponse:
    file_path = JOB_ROOT / job_id / "security_assessment_template_filled.xlsx"
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Filled template not found")
    return FileResponse(
        path=file_path,
        filename=f"{job_id}_security_assessment_template_filled.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
