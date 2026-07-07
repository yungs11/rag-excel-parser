"""다중 계층열(대분류 병합 카테고리 + 소분류 번호) 파싱 검증 + 위임전결 무회귀."""
from __future__ import annotations

import pathlib

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.config import ParserConfig

REPO = pathlib.Path("/Users/xxx/workspace/7.excel-parser/test_doc_excel")
JIKMU_REPO = REPO / "직무전결기준표(2026.05.04).xlsx"   # in-repo(hermetic) — 이걸로 fix 검증
JIKMU_FIX = pathlib.Path("/Users/xxx/Downloads/00.테스트문서/직무전결기준표(2026.05.04)-수정본.xlsx")  # optional
WIJUM = REPO / "2-1. 위임전결기준표(2026.04.17. 개정).xlsx"


def _paths(path):
    cfg = ParserConfig(); cfg.backend = "openpyxl"
    ch, _ = get_backend("openpyxl").parse(path, cfg)
    out = {}
    for c in ch:
        if c["sheet"] != "신한DS" and "위임전결" not in c["sheet"]:
            continue
        src = c.get("source") or {}
        r = src.get("start_row")
        p = " > ".join(c.get("path") or [])
        if r and c["chunk_type"] in ("matrix_fact", "delegation_rule", "hierarchy_node"):
            out.setdefault(r, p)
    return out


@pytest.mark.skipif(not JIKMU_REPO.exists(), reason="직무전결 원본(in-repo) 없음")
def test_jikmu_category_parent_preserved():
    # in-repo 직무전결 원본(hermetic). 병합 카테고리 '경영 일반' 이 하위의 부모로 유지 + '1. 사업계획' 복원.
    paths = _paths(JIKMU_REPO)
    all_txt = " || ".join(paths.values())
    assert "경영 일반 > 1. 사업계획" in all_txt, "1.사업계획 소실/부모 누락"
    # '2. 조직개편' 이 '경영 일반' 밑(고아 top 아님)
    assert any("경영 일반 > 2. 조직개편" in p for p in paths.values()), "조직개편 부모 붕괴"


@pytest.mark.skipif(not JIKMU_FIX.exists(), reason="수정본(Downloads) 없음 — optional")
def test_jikmu_revised_also_fixed():
    all_txt = " || ".join(_paths(JIKMU_FIX).values())
    assert "경영 일반 > 1. 사업계획" in all_txt


@pytest.mark.skipif(not WIJUM.exists(), reason="위임전결 없음")
def test_wijum_deep_paths_unchanged():
    paths = _paths(WIJUM)
    # 위임전결 대표 깊은 path 가 그대로 나온다(회귀 없음)
    assert any("1. 업무전반 > 라. 직원의 출장" in p and "(1)" in p and "(가) 팀장" in p
               for p in paths.values()), "위임전결 깊은 계층 회귀"
    # (위임전결 전수 회귀는 test_column_hierarchy_regression 골든 스냅샷이 담당)
