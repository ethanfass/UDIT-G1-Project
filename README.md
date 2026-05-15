# ISO-Style Security Assessment Toolkit (UDIT-G1-Project)

This project takes one or more self-submitted security questionnaires (plus optional evidence) and produces an ISO-style maturity assessment.
The outputs are Excel workbooks designed for both technical reviewers and stakeholders: a filled control template and an assessment report with scoring, findings, and remediation guidance.

## What this produces

- `security_assessment_template_filled.xlsx`: filled control template (per-control results)
- `assessment_report.xlsx`: assessment report workbook (includes `Executive Summary`, `Section Summary`, `Findings`, `Assessment`, `Rubric`, and `Sources` sheets)

## Prerequisites

- Python 3.12+ (examples use `py -3.12`)
- Node.js 18+ (only needed for the React GUI)
- A master template workbook (commonly named `master_iso_template.xlsx` in the repo root)

## Setup (Python)

Create and activate a virtual environment:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies (covers both CLI + GUI API):

```powershell
py -3.12 -m pip install -r requirements_gui.txt
```

## Configure your Gemini API key (no secrets committed)

The runner resolves the key in this order:

1) `GEMINI_API_KEY` (environment variable)
2) `GOOGLE_API_KEY` (environment variable)
3) Windows-only encrypted local store (`.local_secrets/gemini_api_key.json`)

### Option A: environment variable

PowerShell:

```powershell
$env:GEMINI_API_KEY = "<your-key-here>"
```

Optional: `.env` file

- If `python-dotenv` is installed (it is included via `requirements_gui.txt`), the CLI and GUI API will load a `.env` file automatically.
- `.env` is gitignored.

Example `.env` (do not commit real keys):

```text
GEMINI_API_KEY=<your-key-here>
```

### Option B: encrypted local store (Windows DPAPI)

```powershell
py -3.12 manage_gemini_key.py --set
```

Useful commands:

```powershell
py -3.12 manage_gemini_key.py --status
py -3.12 manage_gemini_key.py --test
py -3.12 manage_gemini_key.py --clear
```

## Run an assessment (CLI)

Run the modern entrypoint directly:

```powershell
py -3.12 assessment_runner_core.py `
  --template master_iso_template.xlsx `
  --company Grammarly `
  --questionnaires questiondocs\gramvsa.xlsx questiondocs\gramsigcore.xlsx questiondocs\gramhecvat.xlsx `
  --assessment_mode questionnaire `
  --output_assessment assessment_report.xlsx `
  --template_output security_assessment_template.xlsx `
  --rubric_output security_assessment_rubric.xlsx `
  --report_template_output assessment_report_template.xlsx `
  --filled_template_output security_assessment_template_filled.xlsx
```

Compatibility wrapper (delegates to the same implementation):

```powershell
py -3.12 iso_assessment_runner.py --help
```

### Generate templates only

```powershell
py -3.12 assessment_runner_core.py --template master_iso_template.xlsx --build_only
```

Or use the dedicated builder:

```powershell
py -3.12 build_assessment_assets.py --template master_iso_template.xlsx
```

### Scoring model

- Each control receives a maturity level: `Very Low`, `Low`, `Medium`, `High`, `Very High`.
- Levels map to a numeric score `0..4`.
- Overall percent is the unweighted maturity score percent across all controls.
- The report includes both an overall security level and a letter grade.

### Utility: list available Gemini models

```powershell
py -3.12 list_models.py
```

## Run the Web GUI (React + FastAPI)

The GUI supports:

- Entering a company name
- Uploading questionnaires (drag/drop)
- Viewing live progress updates
- Downloading the generated workbooks

### 1) Start the Python API

From the repo root:

```powershell
py -3.12 -m uvicorn gui_api:app --reload --host 127.0.0.1 --port 8000
```

### 2) Start the React frontend

In a second terminal:

```powershell
cd assessment_gui
npm install
npm run dev
```

Development behavior:

- The Vite dev server proxies `/api` to `http://127.0.0.1:8000`.
- The GUI writes job artifacts to `gui_jobs/<job_id>/` (gitignored).

## Environment variables

- `GEMINI_API_KEY`: Gemini API key (preferred)
- `GOOGLE_API_KEY`: alternate env var name supported by the code
- `VITE_API_BASE`: frontend-only; base URL prefix for API calls (e.g. `https://your-api.example.com`). If unset, the GUI uses relative URLs and relies on a dev proxy.

## High-level architecture

- CLI path
  - `assessment_runner_core.py` loads controls from the master template
  - evidence is chunked from submitted questionnaires
  - Gemini is called in batches (with top-k evidence per control)
  - workbooks are written with `openpyxl`

- GUI path
  - `assessment_gui/` submits uploads to `gui_api.py`
  - `gui_api.py` runs a background thread per job, updates in-memory progress, and exposes download endpoints
  - outputs are written under `gui_jobs/<job_id>/`

## Major files and folders

- `assessment_runner_core.py`: core assessment engine + CLI
- `gui_api.py`: FastAPI backend for the GUI
- `assessment_gui/`: React + Vite frontend
- `build_assessment_assets.py`: generates blank template/rubric/report template workbooks
- `gemini_secret_store.py`: resolves keys from env or Windows DPAPI-encrypted local file
- `manage_gemini_key.py`: CLI to set/status/test/clear the DPAPI-encrypted local key
- `gui_jobs/`: runtime artifacts (uploads + generated workbooks; gitignored)
- `questiondocs/`: sample inputs

## Deployment (if applicable)

This is a course project and is not production-hardened. If you deploy it, treat it as an internal tool.

Backend (API):

- Run `uvicorn gui_api:app --host 0.0.0.0 --port 8000` behind a reverse proxy.
- Update the CORS allowlist in `gui_api.py` for your deployed frontend origin(s).

Frontend (UI):

- Build with `npm run build`.
- If the API is hosted on a different origin, set `VITE_API_BASE` at build time.
- Serve `assessment_gui/dist` via a static host (IIS, Nginx, etc.).

## Known limitations / incomplete features

- No authentication/authorization; anyone who can reach the API can submit jobs and download results.
- Job status is in-memory; restarting the API server loses active job state (generated files remain on disk).
- `gui_jobs/` can grow unbounded; no automatic cleanup policy yet.
- Local encrypted key storage is Windows-only (DPAPI). On non-Windows, set `GEMINI_API_KEY`/`GOOGLE_API_KEY`.

## Future work recommendations

- Add authentication and per-user job isolation.
- Persist jobs/results in a database and support resume/retry.
- Add a job cleanup scheduler and configurable retention.
- Make secret storage cross-platform (or use a managed secret store in deployment).
- Add deterministic tests for report generation and scoring.
