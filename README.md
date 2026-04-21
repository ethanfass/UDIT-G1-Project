# ISO-Style Security Assessment Toolkit

This project combines one or more self-submitted questionnaires, maps them to a shared control template, evaluates each control against a maturity rubric, and produces both a filled template and an assessment report with scoring, findings, and remediation guidance.

## What Changed

- Replaced the brittle numeric-only scoring flow with an unweighted maturity model:
  `Very Low`, `Low`, `Medium`, `High`, `Very High`
- Added richer workbook outputs:
  `Executive Summary`, `Section Summary`, `Findings`, `Assessment`, `Legacy Assessment`, `Rubric`, and `Sources`
- Added remediation-focused output for every weak control
- Reduced dependency on the checked-in virtualenv by using `openpyxl` plus direct Gemini REST calls
- Added local evidence retrieval so each Gemini prompt is based on the most relevant questionnaire excerpts instead of blindly sending every document every time

## Store Your API Key Once

If you do not want to set the API key in PowerShell every time, store it locally in encrypted form:

```powershell
py -3.12 manage_gemini_key.py --set
```

This stores the key in `.local_secrets/gemini_api_key.json`, encrypted with your current Windows user account via DPAPI.

Useful commands:

```powershell
py -3.12 manage_gemini_key.py --status
py -3.12 manage_gemini_key.py --clear
```

After that, the runner and model lister will automatically use the stored key.

## Build The Detailed Assets

```powershell
py -3.12 build_assessment_assets.py --template master_iso_template.xlsx
```

This generates the built-in workbook set:

- `security_assessment_template.xlsx`
- `security_assessment_rubric.xlsx`
- `assessment_report_template.xlsx`

## Run An Assessment

Either set your API key in `GEMINI_API_KEY` / `GOOGLE_API_KEY` or store it once with `manage_gemini_key.py`, then run:

```powershell
py -3.12 iso_assessment_runner.py `
  --template master_iso_template.xlsx `
  --company Grammarly `
  --questionnaires questiondocs\gramvsa.xlsx questiondocs\gramsigcore.xlsx questiondocs\gramhecvat.xlsx `
  --output_assessment assessment_report.xlsx `
  --template_output security_assessment_template.xlsx `
  --rubric_output security_assessment_rubric.xlsx `
  --report_template_output assessment_report_template.xlsx `
  --filled_template_output security_assessment_template_filled.xlsx
```

## Scoring Model

- Each control receives a maturity level from `Very Low` to `Very High`
- Each level maps to a numeric score from `0` to `4`
- Overall percent = earned level score / maximum possible level score
- The final report includes both:
  - an overall security level
  - a letter grade

## Output Highlights

- `Template`: the detailed control template, either blank or filled with company-specific results
- `Assessment`: full per-control assessment with evidence, gaps, risk impact, remediation, confidence, and rationale
- `Findings`: prioritized list of below-target controls
- `Section Summary`: category-level rollup for the assessment
- `Executive Summary`: stakeholder-friendly summary with strongest areas and biggest risks
- `Rubric`: the maturity scoring criteria from `Very Low` to `Very High`

## Web GUI (React + FastAPI)

A browser interface is included for business-friendly use:

- Enter company name
- Drag and drop completed questionnaires
- Watch live analysis progress
- Review summary results
- Export generated Excel workbooks

### 1) Install GUI API dependencies

```powershell
py -3.12 -m pip install -r requirements_gui.txt
```

### 2) Start the Python API

From the project root:

```powershell
py -3.12 -m uvicorn gui_api:app --reload --host 127.0.0.1 --port 8000
```

### 3) Start the React frontend

In a second terminal:

```powershell
cd assessment_gui
npm install
npm run dev
```

By default, Vite proxies `/api` to `http://127.0.0.1:8000`.

### Notes

- The GUI writes job files to `gui_jobs/<job_id>/`
- Assessment export and filled template export are both available after completion
- If your Gemini key is not set, store it once with:

```powershell
py -3.12 manage_gemini_key.py --set
```
