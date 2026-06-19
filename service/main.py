"""Excel Parser FastAPI 앱 (:18055).

엔드포인트:
  GET    /healthz              — 라이브니스 프로브.
  GET    /info                 — 서비스 메타 + 가용 청크타입 + 한도.
  POST   /parse                — 동기 파싱 (sync_max_bytes 초과 시 413).
  POST   /parse/jobs/file      — 비동기 잡 생성 → {job_id, status}.
  GET    /parse/jobs/{id}      — 잡 상태/진행/결과/에러.
  DELETE /parse/jobs/{id}      — 잡 취소 (204).

실행:
    .venv/bin/uvicorn service.main:app --port 18055
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Callable

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

# 프로젝트 루트를 sys.path 에 추가 (uvicorn service.main:app 실행 시)
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from excel_parser_rag import PARSER_VERSION
from excel_parser_rag.backends import get_backend
from excel_parser_rag.chunking.chunk_schema import CHUNK_TYPES
from excel_parser_rag.config import DEFAULT_CHUNK_PROFILES, ParserConfig
from excel_parser_rag.loaders.xls_converter import find_soffice
from service.jobs import JobStore

_SYNC_MAX_BYTES: int = 5 * 1024 * 1024  # 5 MB
_JOB_TIMEOUT_S: float = 300.0
_ALLOWED_SUFFIXES = {".xlsx", ".xlsm", ".xls"}

# ─── 서버 기본 백엔드 설정 (env 로 1회 지정 → 요청마다 플래그 불필요) ───
#   EXCEL_PARSER_BACKEND : "kordoc"(기본) | "openpyxl"
#   KORDOC_BIN           : kordoc CLI (예: "node /tmp/kordoc/dist/cli.js") — md 자동생성
#   KORDOC_MD_OUT        : 자동생성 md 저장 디렉토리 (기본 임시)
_BACKEND = os.environ.get("EXCEL_PARSER_BACKEND", "kordoc")
_KORDOC_BIN = os.environ.get("KORDOC_BIN")
_KORDOC_MD_OUT = os.environ.get("KORDOC_MD_OUT")

app = FastAPI(
    title="excel_parser",
    description="Excel → RAG JSONL 청크 변환 서비스",
    version="1.0.0",
)

_job_store: JobStore | None = None


@app.on_event("startup")
def _startup() -> None:
    global _job_store
    _job_store = JobStore(timeout_s=_JOB_TIMEOUT_S)


@app.on_event("shutdown")
def _shutdown() -> None:
    if _job_store is not None:
        _job_store.shutdown()


def _get_store() -> JobStore:
    if _job_store is None:
        raise RuntimeError("JobStore 미초기화")
    return _job_store


# ─────────────────────────── 헬퍼 ───────────────────────────

def _validate_upload(file: UploadFile) -> None:
    if not file.filename:
        raise HTTPException(status_code=422, detail="파일명이 없습니다.")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in _ALLOWED_SUFFIXES:
        raise HTTPException(
            status_code=422,
            detail=f"지원하지 않는 확장자: {suffix!r}. 허용: {sorted(_ALLOWED_SUFFIXES)}",
        )


def _parse_options(options_json: str | None) -> ParserConfig:
    """서버 기본(env: backend/kordoc_bin) 위에 요청별 options_json 을 덮어쓴다."""
    data: dict[str, Any] = {}
    if options_json:
        try:
            data = dict(json.loads(options_json))
        except json.JSONDecodeError as exc:
            raise HTTPException(status_code=422, detail=f"options_json JSON 파싱 실패: {exc}") from exc
    data.setdefault("backend", _BACKEND)
    if _KORDOC_BIN and "kordoc_bin" not in data:
        data["kordoc_bin"] = _KORDOC_BIN
    if _KORDOC_MD_OUT and "kordoc_md_out" not in data:
        data["kordoc_md_out"] = _KORDOC_MD_OUT
    return ParserConfig.from_dict(data)


def _run_parse(file_bytes: bytes, filename: str, config: ParserConfig, *, file_suffix: str = ".xlsx") -> dict[str, Any]:
    """실제 파싱 실행 — asyncio.to_thread 에서 호출.

    file_suffix: 업로드 원본 파일의 확장자 (openpyxl 인식용). filename 은 결과의 source_file 표시용.
    """
    with tempfile.NamedTemporaryFile(suffix=file_suffix, delete=False, prefix="excel_parser_") as tmp:
        tmp.write(file_bytes)
        tmp_path = Path(tmp.name)
    try:
        t0 = time.monotonic()
        chunks, stats = get_backend(config.backend).parse(tmp_path, config)
        timing_ms = round((time.monotonic() - t0) * 1000, 1)
    finally:
        tmp_path.unlink(missing_ok=True)
    return {
        "source_file": filename,
        "parser_version": PARSER_VERSION,
        "chunks": chunks,
        "stats": stats,
        "timing_ms": timing_ms,
    }


# ─────────────────────────── 엔드포인트 ───────────────────────────

@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/info")
def info() -> dict[str, Any]:
    return {
        "service": "excel_parser",
        "version": app.version,
        "parser_version": PARSER_VERSION,
        "chunk_types": list(CHUNK_TYPES),
        "default_profiles": list(DEFAULT_CHUNK_PROFILES),
        "limits": {
            "sync_max_bytes": _SYNC_MAX_BYTES,
            "job_timeout_s": _JOB_TIMEOUT_S,
        },
        "backend": _BACKEND,
        "kordoc_bin": _KORDOC_BIN,
        "kordoc_auto_md": bool(_KORDOC_BIN),
        "soffice_available": find_soffice() is not None,
    }


@app.post("/parse")
async def parse_sync(
    file: UploadFile = File(...),
    doc_name: str | None = Form(default=None),
    options_json: str | None = Form(default=None),
) -> dict[str, Any]:
    """동기 파싱 (소형/디버그 전용). sync_max_bytes 초과 시 413."""
    _validate_upload(file)
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=422, detail="empty file")
    if len(raw) > _SYNC_MAX_BYTES:
        raise HTTPException(
            status_code=413,
            detail=(
                f"파일 크기({len(raw):,} bytes)가 동기 한도({_SYNC_MAX_BYTES:,} bytes)를 초과합니다. "
                "POST /parse/jobs/file (비동기) 을 사용하세요."
            ),
        )
    config = _parse_options(options_json)
    upload_name = file.filename or "upload.xlsx"
    file_suffix = Path(upload_name).suffix.lower()
    filename = doc_name or upload_name
    try:
        return await asyncio.to_thread(_run_parse, raw, filename, config, file_suffix=file_suffix)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"{type(exc).__name__}: {exc}") from exc


@app.post("/parse/jobs/file", status_code=202)
async def create_parse_job(
    file: UploadFile = File(...),
    doc_name: str | None = Form(default=None),
    options_json: str | None = Form(default=None),
) -> dict[str, Any]:
    """비동기 파싱 잡 생성. 즉시 {job_id, status} 반환."""
    _validate_upload(file)
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=422, detail="empty file")
    config = _parse_options(options_json)
    upload_name = file.filename or "upload.xlsx"
    file_suffix = Path(upload_name).suffix.lower()
    filename = doc_name or upload_name

    async def work(report: Callable[[float], None]) -> dict[str, Any]:
        report(0.05)
        result = await asyncio.to_thread(_run_parse, raw, filename, config, file_suffix=file_suffix)
        report(1.0)
        return result

    store = _get_store()
    job = await asyncio.to_thread(store.create, work)
    return {"job_id": job.id, "status": job.status.value}


@app.get("/parse/jobs/{job_id}")
async def get_parse_job(job_id: str) -> dict[str, Any]:
    """잡 상태/진행/결과/에러 조회."""
    store = _get_store()
    job = await asyncio.to_thread(store.get, job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"job not found: {job_id}")
    return job.to_public()


@app.delete("/parse/jobs/{job_id}", status_code=204)
async def cancel_parse_job(job_id: str) -> Response:
    """잡 취소 (멱등 204)."""
    store = _get_store()
    job = await asyncio.to_thread(store.cancel, job_id)
    if job is None:
        raise HTTPException(status_code=404, detail=f"job not found: {job_id}")
    return Response(status_code=204)
