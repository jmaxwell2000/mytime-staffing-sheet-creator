"""FastAPI entrypoints for MyTime Staffing Sheet Creator."""

from __future__ import annotations

from io import BytesIO
import os
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from app.parser import parse_workbook
from app.workbook import build_staffing_workbook, filename_for_records

ROOT = Path(__file__).resolve().parent
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "10"))
XLSX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

app = FastAPI(title="MyTime Staffing Sheet Creator", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (ROOT / "templates" / "index.html").read_text(encoding="utf-8")


@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/convert")
async def convert(
    workbook: UploadFile = File(...),
    include_back_pages: bool = Form(False),
    back_page_offset: float = Form(0.0),
) -> StreamingResponse:
    if not _is_xlsx(workbook):
        raise HTTPException(status_code=400, detail="Upload a valid .xlsx workbook.")

    if include_back_pages and not -0.10 <= back_page_offset <= 0.30:
        raise HTTPException(status_code=400, detail="Back page offset must be between -0.10 and 0.30 inches.")

    raw = await workbook.read()
    max_bytes = MAX_UPLOAD_MB * 1024 * 1024
    if len(raw) > max_bytes:
        raise HTTPException(status_code=413, detail=f"Workbook is too large. Maximum size is {MAX_UPLOAD_MB} MB.")

    try:
        parsed = parse_workbook(raw)
    except Exception as exc:
        raise HTTPException(status_code=422, detail="Could not interpret the workbook structure.") from exc

    if not parsed.records:
        raise HTTPException(status_code=422, detail="No readable staffing records were found.")

    try:
        output = build_staffing_workbook(
            parsed.records,
            include_back_pages=include_back_pages,
            back_page_offset=back_page_offset if include_back_pages else 0.0,
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail="Could not generate the staffing workbook.") from exc

    filename = filename_for_records(parsed.records)
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return StreamingResponse(BytesIO(output), media_type=XLSX_MEDIA_TYPE, headers=headers)


def _is_xlsx(upload: UploadFile) -> bool:
    name = upload.filename or ""
    return name.lower().endswith(".xlsx")
