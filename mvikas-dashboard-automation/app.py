from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Optional

from apscheduler.schedulers.background import BackgroundScheduler
from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from mvikas_backend.config import settings
from mvikas_backend.service import commit_outputs, update_dashboard

app = FastAPI(title="MVIKAS Dashboard Automation Backend", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

settings.static_dir.mkdir(parents=True, exist_ok=True)
app.mount("/dashboard", StaticFiles(directory=settings.static_dir, html=True), name="dashboard")

scheduler: BackgroundScheduler | None = None


def run_scheduled_update() -> None:
    if not settings.source_url:
        print("[MVIKAS] Scheduled update skipped: MVIKAS_SOURCE_URL is empty")
        return
    try:
        result = update_dashboard(settings.source_url, settings.static_dir)
        if settings.commit_to_git:
            commit_outputs(settings)
        print("[MVIKAS] Scheduled update completed", result)
    except Exception as exc:
        print("[MVIKAS] Scheduled update failed:", repr(exc))


@app.on_event("startup")
def start_scheduler() -> None:
    global scheduler
    if not settings.auto_start:
        return
    hour, minute = (int(x) for x in settings.update_time.split(":", 1))
    scheduler = BackgroundScheduler(timezone=settings.timezone)
    scheduler.add_job(run_scheduled_update, "cron", hour=hour, minute=minute, id="daily-dashboard-update", replace_existing=True)
    scheduler.start()


@app.on_event("shutdown")
def stop_scheduler() -> None:
    if scheduler:
        scheduler.shutdown(wait=False)


@app.get("/", response_class=HTMLResponse)
def home() -> str:
    return """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>MVIKAS Dashboard Backend</title>
  <style>
    body{font-family:Inter,system-ui,-apple-system,sans-serif;background:#f8fafc;color:#0f172a;margin:0;padding:32px}
    .wrap{max-width:860px;margin:auto;background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:26px;box-shadow:0 1px 3px rgba(0,0,0,.05)}
    h1{margin:0 0 8px;font-size:24px}.muted{color:#64748b}.row{display:grid;gap:18px;grid-template-columns:1fr 1fr}@media(max-width:720px){.row{grid-template-columns:1fr}}
    form{border:1px solid #e2e8f0;border-radius:10px;padding:18px;margin-top:18px;background:#fbfdff}label{display:block;font-weight:700;margin:12px 0 6px}input{width:100%;padding:10px;border:1px solid #cbd5e1;border-radius:8px}button{margin-top:14px;background:#004df5;color:#fff;border:0;border-radius:8px;padding:11px 16px;font-weight:700;cursor:pointer}a{color:#004df5}.pill{display:inline-block;padding:4px 9px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:999px;font-size:12px;font-weight:700}
  </style>
</head>
<body>
<div class="wrap">
  <span class="pill">MVIKAS Automation</span>
  <h1>Dashboard Update Backend</h1>
  <p class="muted">Upload an Excel file for manual update, or trigger a Google Sheet update. Generated files: <b>script.js</b> and <b>latest_data.json</b>.</p>
  <p><a href="/dashboard/" target="_blank">Open local dashboard preview</a> · <a href="/health" target="_blank">Health</a></p>
  <div class="row">
    <form method="post" action="/api/upload" enctype="multipart/form-data">
      <h3>Manual Excel update</h3>
      <label>Excel file (.xlsx)</label>
      <input type="file" name="file" accept=".xlsx,.xlsm,.xls" required />
      <label>Report date override (optional)</label>
      <input type="date" name="report_date" />
      <button type="submit">Generate dashboard JS</button>
    </form>
    <form method="post" action="/api/update">
      <h3>Google Sheet / URL update</h3>
      <label>Source URL (optional if MVIKAS_SOURCE_URL is set)</label>
      <input type="url" name="source_url" placeholder="https://docs.google.com/spreadsheets/d/..." />
      <label>Report date override (optional)</label>
      <input type="date" name="report_date" />
      <button type="submit">Update from source</button>
    </form>
  </div>
</div>
</body>
</html>
"""


@app.get("/health")
def health() -> dict[str, object]:
    return {
        "ok": True,
        "staticDir": str(settings.static_dir),
        "sourceConfigured": bool(settings.source_url),
        "schedulerEnabled": settings.auto_start,
        "dailyUpdateTime": settings.update_time,
        "timezone": settings.timezone,
    }


@app.post("/api/update")
def api_update(
    source_url: Optional[str] = Form(default=None),
    report_date: Optional[str] = Form(default=None),
    commit: bool = Query(default=False),
) -> JSONResponse:
    source = (source_url or settings.source_url or "").strip()
    if not source:
        raise HTTPException(status_code=400, detail="source_url is required or set MVIKAS_SOURCE_URL")
    try:
        result = update_dashboard(source, settings.static_dir, report_date=report_date or None)
        if commit or settings.commit_to_git:
            commit_outputs(settings)
        return JSONResponse(result)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc


@app.post("/api/upload")
async def api_upload(
    file: UploadFile = File(...),
    report_date: Optional[str] = Form(default=None),
    commit: bool = Query(default=False),
) -> JSONResponse:
    suffix = Path(file.filename or "upload.xlsx").suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(await file.read())
        tmp_path = Path(tmp.name)
    try:
        result = update_dashboard(tmp_path, settings.static_dir, report_date=report_date or None)
        if commit or settings.commit_to_git:
            commit_outputs(settings)
        return JSONResponse(result)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    finally:
        tmp_path.unlink(missing_ok=True)
