"""게이트 수용 기준(실파일) — 차단/통과 파일이 의도대로 동작하는지 end-to-end 확인.

설계: 8.kb-pipeline/docs/superpowers/specs/2026-06-29-excel-gate-postparse-design.md §9
게이트 단위 = 파일 단위(gate_summary.ok = 모든 시트 findings 0건).
"""
import pathlib

import pytest

from excel_parser_rag.gate.excel_gate import compute_gate_summary
from excel_parser_rag.pipeline import parse_excel_for_rag

ROOT = pathlib.Path("/Users/xxx/workspace")

# (path, 차단 사유가 나타날 시트 substring | None)
CASES_BLOCK = [
    (ROOT / "7.excel-parser/test_doc_excel/신한자산신탁_외부테이터_필요사이트 정리.xlsx", "법령리스트"),
    (ROOT / "excel-parser-markitdown/test_doc_excel/251210_중소형그룹사_AX추진지원_WBS_v0.1_sys.xlsx", None),
    # 파일 단위 차단(사용자 결정): NAC연계 시트가 진짜 side_by_side 라 파일 ok=False.
    (ROOT / "excel-parser-markitdown/test_doc_excel/신한자산신탁_자산목록_v20251013.xlsx", "NAC연계"),
]
CASES_PASS = [
    pathlib.Path("/Users/xxx/Downloads/aws_cost_estimate.xlsx"),
    ROOT / "7.excel-parser/test_doc_excel/2-1. 위임전결기준표(2026.04.17. 개정).xlsx",
]


def _summ(p):
    chunks, _ = parse_excel_for_rag(str(p))
    return compute_gate_summary(p, chunks)


@pytest.mark.parametrize("path,sheet", CASES_BLOCK, ids=lambda v: getattr(v, "name", v))
def test_block_cases(path, sheet):
    if not path.exists():
        pytest.skip(f"missing test file: {path}")
    s = _summ(path)
    assert s["ok"] is False, f"{path.name} 은 차단되어야 한다"
    if sheet is not None:
        blocked = {x["sheet"] for x in s["sheets"] if not x["ok"]}
        assert any(sheet in b for b in blocked), f"{path.name}: '{sheet}' 시트가 차단 사유여야 한다 (실제: {blocked})"


@pytest.mark.parametrize("path", CASES_PASS, ids=lambda v: v.name)
def test_pass_cases(path):
    if not path.exists():
        pytest.skip(f"missing test file: {path}")
    s = _summ(path)
    assert s["ok"] is True, f"{path.name} 은 통과되어야 한다 (실제 차단 시트: {[x['sheet'] for x in s['sheets'] if not x['ok']]})"
