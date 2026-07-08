"""delegation_rule: 마커 해석 없이 header:값 매핑."""
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
def test_header_value_no_marker_interpretation():
    dr = _deleg(JIKMU)
    allc = " || ".join(c.get("content_text") or "" for c in dr)
    # 나.월간: 원문 ○ + 보고 를 header:값 으로 (전결권자/관계 라벨 없이)
    w = [c for c in dr if "나. 월간" in (c.get("content_text") or "") and "경영실적 보고" in (c.get("content_text") or "")]
    assert w, "나.월간 그룹 없음"
    txt = w[0].get("content_text") or ""
    assert "부문장:○" in txt, f"부문장:○ 원문 매핑 누락: {txt}"
    assert "CEO:보고" in txt, f"CEO:보고 매핑 누락: {txt}"
    # 파서가 붙이던 해석 라벨 제거 — '전결권자는 X이다' 문장·'관계:' 래퍼. (셀 데이터에 '전결권자'
    # 문자열이 들어간 건 원문 header:값 이므로 허용 → 파서-전용 문장형만 검사.)
    assert "전결권자는" not in allc, "전결권자 해석 문장 잔존"
    assert "관계:" not in allc, "관계: 라벨 잔존"


@pytest.mark.skipif(not JIKMU.exists(), reason="직무전결 원본 없음")
def test_header_bigo_as_value():
    dr = _deleg(JIKMU)
    assert any("비고:내규의 제정 및 개폐 시" in (c.get("content_text") or "") for c in dr), \
        "비고 header:값 매핑 누락"
