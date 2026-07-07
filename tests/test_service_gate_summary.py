"""Tests: gate_summary 노출 in /parse response (Task 2)."""
import json
import pathlib

from fastapi.testclient import TestClient

from service.main import app

EXCEL = pathlib.Path("/Users/xxx/workspace/7.excel-parser/test_doc_excel")
_OPTS = json.dumps({"backend": "openpyxl"})


def test_parse_includes_gate_summary():
    f = EXCEL / "신한자산신탁_외부테이터_필요사이트 정리.xlsx"
    with TestClient(app) as client, f.open("rb") as fh:
        r = client.post(
            "/parse",
            files={
                "file": (
                    f.name,
                    fh,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            data={"options_json": _OPTS},
        )
    assert r.status_code == 200
    gs = r.json()["stats"]["gate_summary"]
    assert gs["ok"] is False
    assert any("side_by_side" in {x["code"] for x in s["findings"]} for s in gs["sheets"])


def test_gate_computation_failure_blocks(monkeypatch):
    # gate '계산' 예외는 보수적 차단(ok=False) 이어야 한다 (spec §8).
    import service.main as m

    def boom(*a, **k):
        raise RuntimeError("gate boom")

    monkeypatch.setattr(m, "compute_gate_summary", boom)
    f = EXCEL / "신한자산신탁_외부테이터_필요사이트 정리.xlsx"
    with TestClient(app) as client, f.open("rb") as fh:
        r = client.post(
            "/parse",
            files={
                "file": (
                    f.name,
                    fh,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
            data={"options_json": _OPTS},
        )
    assert r.status_code == 200
    assert r.json()["stats"]["gate_summary"]["ok"] is False
