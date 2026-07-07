"""delegation_rule: 비-○ 마커(보고) 캡처 + 병합 시 비고/관계 보존."""
from __future__ import annotations

import pathlib

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.config import ParserConfig

JIKMU = pathlib.Path("/Users/xxx/workspace/7.excel-parser/test_doc_excel/직무전결기준표(2026.05.04).xlsx")


def _deleg(path):
    cfg = ParserConfig(); cfg.backend = "openpyxl"
    ch, _ = get_backend("openpyxl").parse(path, cfg)
    return [c for c in ch if c["chunk_type"] == "delegation_rule"]


@pytest.mark.skipif(not JIKMU.exists(), reason="직무전결 원본 없음")
def test_report_marker_captured():
    # '경영실적 보고 > 나. 월간' = 부문장 ○ + CEO '보고'(비-○). 병합그룹 content 에 관계 '보고 CEO' 가 남아야.
    dr = _deleg(JIKMU)
    hit = [c for c in dr if "경영실적 보고" in (c.get("content_text") or "")
           and "나. 월간" in (c.get("content_text") or "")]
    assert hit, "경영실적>나.월간 그룹 없음"
    txt = hit[0].get("content_text") or ""
    assert "부문장" in txt, f"나.월간 전결권자 부문장 누락: {txt}"
    # '보고 CEO' 는 relations(fields['관계']) 에서만 나오는 연속 문자열(path/다른행 오탐 차단).
    assert "보고 CEO" in txt, f"CEO 보고 관계 미캡처: {txt}"


@pytest.mark.skipif(not JIKMU.exists(), reason="직무전결 원본 없음")
def test_header_bigo_preserved_through_merge():
    # r10 '3. 내규 관리' 헤더의 고유 비고 '내규의 제정 및 개폐 시…' 가 병합 후에도 delegation_rule 에 남아야.
    dr = _deleg(JIKMU)
    assert any("내규의 제정 및 개폐 시" in (c.get("content_text") or "") for c in dr), \
        "r10 헤더행 비고가 병합 시 소실(_line 이 비고 미포함)"
