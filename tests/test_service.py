"""Excel Parser 서비스 테스트 (TestClient — 라이브 호출 0).

검증 항목:
  - /healthz, /info 정상 응답.
  - /parse 동기 413 (sync_max_bytes 초과).
  - /parse/jobs/file 잡 submit → poll 성공 (위임전결기준표.xlsx 실파일).
  - 빈 파일 422, 미지원 확장자 422.
  - options_json → ParserConfig 전달 (chunk_profiles 필터 확인).
  - 파싱 실패 잡 → status=failed 표면화.
"""

from __future__ import annotations

import io
import json
import sys
import time
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

TESTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = TESTS_DIR.parent
for _p in (str(PROJECT_ROOT), str(TESTS_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from service.main import app, _SYNC_MAX_BYTES, _ALLOWED_SUFFIXES

REAL_EXCEL = PROJECT_ROOT / "2-1. 위임전결기준표(2026.04.17. 개정).xlsx"


@pytest.fixture(scope="module")
def client():
    with TestClient(app) as c:
        yield c


# ─────────────────────────── /healthz / /info ───────────────────────────

def test_healthz(client):
    resp = client.get("/healthz")
    assert resp.status_code == 200
    assert resp.json() == {"status": "ok"}


def test_info(client):
    resp = client.get("/info")
    assert resp.status_code == 200
    body = resp.json()
    assert body["service"] == "excel_parser"
    assert "parser_version" in body
    assert "chunk_types" in body
    assert "default_profiles" in body
    assert isinstance(body["limits"]["sync_max_bytes"], int)
    assert isinstance(body["limits"]["job_timeout_s"], (int, float))
    assert "soffice_available" in body


# ─────────────────────────── /parse (동기) ───────────────────────────

def test_parse_sync_413_when_too_large(client):
    """sync_max_bytes 초과 → 413."""
    oversized = b"x" * (_SYNC_MAX_BYTES + 1)
    resp = client.post(
        "/parse",
        files={"file": ("big.xlsx", io.BytesIO(oversized), "application/octet-stream")},
    )
    assert resp.status_code == 413
    assert "비동기" in resp.json()["detail"] or "/parse/jobs/file" in resp.json()["detail"]


def test_parse_sync_422_empty(client):
    resp = client.post(
        "/parse",
        files={"file": ("empty.xlsx", io.BytesIO(b""), "application/octet-stream")},
    )
    assert resp.status_code == 422


def test_parse_sync_422_bad_ext(client):
    resp = client.post(
        "/parse",
        files={"file": ("doc.pdf", io.BytesIO(b"fake"), "application/octet-stream")},
    )
    assert resp.status_code == 422
    assert "확장자" in resp.json()["detail"] or "지원하지 않는" in resp.json()["detail"]


# ─────────────────────────── /parse/jobs/file ───────────────────────────

def test_parse_job_empty_file(client):
    resp = client.post(
        "/parse/jobs/file",
        files={"file": ("empty.xlsx", io.BytesIO(b""), "application/octet-stream")},
    )
    assert resp.status_code == 422


def test_parse_job_bad_ext(client):
    resp = client.post(
        "/parse/jobs/file",
        files={"file": ("doc.pdf", io.BytesIO(b"fake"), "application/octet-stream")},
    )
    assert resp.status_code == 422


def test_parse_job_invalid_options_json(client, tmp_path):
    fake_xlsx = tmp_path / "test.xlsx"
    fake_xlsx.write_bytes(b"PK")  # 최소 non-empty bytes
    resp = client.post(
        "/parse/jobs/file",
        files={"file": ("test.xlsx", fake_xlsx.read_bytes(), "application/octet-stream")},
        data={"options_json": "not-valid-json"},
    )
    assert resp.status_code == 422
    assert "options_json" in resp.json()["detail"]


def test_parse_job_parsing_failure_surfaces_as_failed(client, tmp_path):
    """파싱 불가 파일 → 잡 status=failed 로 표면화 (성공 위장 금지)."""
    bad = tmp_path / "corrupt.xlsx"
    bad.write_bytes(b"NOTANEXCEL" * 100)
    resp = client.post(
        "/parse/jobs/file",
        files={"file": ("corrupt.xlsx", bad.read_bytes(), "application/octet-stream")},
    )
    assert resp.status_code == 202
    job_id = resp.json()["job_id"]

    # 잡 완료까지 폴링 (최대 15초)
    deadline = time.monotonic() + 15.0
    while True:
        poll = client.get(f"/parse/jobs/{job_id}")
        assert poll.status_code == 200
        body = poll.json()
        if body["status"] in ("succeeded", "failed", "cancelled"):
            break
        assert time.monotonic() < deadline, "잡 완료 타임아웃"
        time.sleep(0.2)

    assert body["status"] == "failed"
    assert body.get("error")


@pytest.mark.skipif(not REAL_EXCEL.exists(), reason="실 Excel 파일 없음")
def test_parse_job_real_excel_e2e(client):
    """위임전결기준표.xlsx → 잡 submit → poll 성공 → 청크 수/응답 구조 확인."""
    raw = REAL_EXCEL.read_bytes()
    resp = client.post(
        "/parse/jobs/file",
        files={"file": (REAL_EXCEL.name, raw, "application/octet-stream")},
        # 테스트 환경엔 kordoc 바이너리가 없으므로 openpyxl 백엔드로 e2e 검증
        data={"doc_name": "위임전결기준표_test", "options_json": '{"backend": "openpyxl"}'},
    )
    assert resp.status_code == 202
    job_id = resp.json()["job_id"]
    assert resp.json()["status"] in ("queued", "running")

    # 폴링 (최대 60초)
    deadline = time.monotonic() + 60.0
    while True:
        poll = client.get(f"/parse/jobs/{job_id}")
        assert poll.status_code == 200
        body = poll.json()
        if body["status"] in ("succeeded", "failed", "cancelled"):
            break
        assert time.monotonic() < deadline, "E2E 잡 완료 타임아웃"
        time.sleep(1.0)

    assert body["status"] == "succeeded", f"잡 실패: {body.get('error')}"
    result = body["result"]

    # 응답 구조 확인
    assert result["source_file"] == "위임전결기준표_test"
    assert "parser_version" in result
    assert isinstance(result["chunks"], list)
    assert len(result["chunks"]) > 100, f"청크 수 너무 적음: {len(result['chunks'])}"
    assert isinstance(result["stats"], dict)
    assert isinstance(result["timing_ms"], (int, float))

    # 청크 필수키 샘플 확인
    sample = result["chunks"][0]
    for key in ("id", "source_file", "chunk_type", "content_text"):
        assert key in sample, f"청크에 {key!r} 없음"


@pytest.mark.skipif(not REAL_EXCEL.exists(), reason="실 Excel 파일 없음")
def test_parse_job_options_chunk_profiles_filter(client):
    """options_json chunk_profiles → 필터 적용 확인."""
    raw = REAL_EXCEL.read_bytes()
    opts = json.dumps({"chunk_profiles": ["delegation_rule"]})
    resp = client.post(
        "/parse/jobs/file",
        files={"file": (REAL_EXCEL.name, raw, "application/octet-stream")},
        data={"options_json": opts},
    )
    assert resp.status_code == 202
    job_id = resp.json()["job_id"]

    deadline = time.monotonic() + 60.0
    while True:
        poll = client.get(f"/parse/jobs/{job_id}")
        body = poll.json()
        if body["status"] in ("succeeded", "failed", "cancelled"):
            break
        assert time.monotonic() < deadline
        time.sleep(1.0)

    assert body["status"] == "succeeded"
    chunks = body["result"]["chunks"]
    types = {c["chunk_type"] for c in chunks}
    # delegation_rule + 플러그인 타입(unsupported_artifact 있다면) 만 존재해야 함
    assert types <= {"delegation_rule", "unsupported_artifact"}, f"필터 미적용: {types}"


# ─────────────────────────── 잡 폴링/취소 ───────────────────────────

def test_get_job_not_found(client):
    resp = client.get("/parse/jobs/nonexistent123")
    assert resp.status_code == 404


def test_cancel_job_not_found(client):
    resp = client.delete("/parse/jobs/nonexistent456")
    assert resp.status_code == 404


def test_info_default_backend_is_auto():
    """env(EXCEL_PARSER_BACKEND) 미설정 시 기본 백엔드가 정확히 'auto' 여야 한다."""
    import os
    if os.environ.get("EXCEL_PARSER_BACKEND"):
        import pytest
        pytest.skip("EXCEL_PARSER_BACKEND env 가 설정되어 기본값 검증 불가")
    from service import main as svc
    assert svc._BACKEND == "auto"
    from fastapi.testclient import TestClient
    with TestClient(svc.app) as client:
        body = client.get("/info").json()
        assert body["backend"] == "auto"
