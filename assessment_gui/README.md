# Assessment GUI (React + Vite)

This folder contains the browser UI for submitting questionnaires, tracking assessment progress, and downloading generated Excel workbooks.

## Prerequisites

- Node.js 18+
- The FastAPI backend running from the repo root (`gui_api.py`)

## Run locally

From this folder:

```bash
npm install
npm run dev
```

Development behavior:

- The dev server proxies `/api` to `http://127.0.0.1:8000` (see `vite.config.js`).
- The UI polls job status via `/api/jobs/{jobId}`.

## Configure API base URL

If you deploy the frontend separately from the backend, set the API base prefix:

- `VITE_API_BASE` (example: `https://your-api.example.com`)

In dev, you typically leave this unset and rely on the Vite proxy.

## Build

```bash
npm run build
```

Static assets are emitted into `dist/`.

## Lint

```bash
npm run lint
```
