"""실문서 스모크 — 전결은 병합/계층, 평면은 kordoc 라우팅. 파일 없으면 skip."""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.backends.auto_backend import detect_delegation_keyword
from excel_parser_rag.config import ParserConfig

DOC_DIR = Path("/Users/xxx/workspace/excel-parser-markitdown/test_doc_excel")
DELEG = DOC_DIR / "2-1. 위임전결기준표(2026.04.17. 개정).xlsx"
ASSET = DOC_DIR / "신한자산신탁_자산목록_v20251013.xlsx"


@pytest.mark.skipif(not DELEG.exists(), reason="실 위임전결 파일 없음")
def test_real_delegation_routes_openpyxl_and_merges():
    cfg = ParserConfig(); cfg.backend = "auto"
    chunks, stats = get_backend("auto").parse(DELEG, cfg)
    assert stats["routed_backend"] == "openpyxl"
    deleg = [c for c in chunks if c["chunk_type"] == "delegation_rule"]
    merged = [c for c in deleg if c.get("metadata", {}).get("merged")]
    assert merged, "실 전결 문서에서 병합 청크 0"
    assert all(len(c["content_text"]) <= 1100 for c in merged)
    # note/code_mapping 보존
    assert any(c["chunk_type"] == "note" for c in chunks)
    assert any(c["chunk_type"] == "code_mapping" for c in chunks)
    # 계층 fallback 0: leaf 가 '행 N' 인 delegation_rule 없음
    bad = [c for c in deleg if c["path"] and str(c["path"][-1]).startswith("행 ")]
    assert not bad, f"계층 fallback 발생: {len(bad)}"


@pytest.mark.skipif(not ASSET.exists() or not shutil.which("kordoc"),
                    reason="자산목록 파일 또는 kordoc CLI 없음")
def test_real_asset_list_routes_kordoc_no_explosion():
    cfg = ParserConfig(); cfg.backend = "auto"; cfg.kordoc_bin = "kordoc"
    chunks, stats = get_backend("auto").parse(ASSET, cfg)
    assert stats["routed_backend"] == "kordoc"
    # 쓰레기 matrix_fact 폭발 없음 (openpyxl 미접촉)
    mf = [c for c in chunks if c["chunk_type"] == "matrix_fact"]
    assert len(mf) < 100, f"matrix_fact 과다({len(mf)}) — kordoc 라우팅 실패 의심"


# 다중 좌측 라벨열(구분+업무내용) — 업무내용 태스크명이 항목으로 보존되어야 함 (과적합 수정)
JIKMU = Path("/Users/xxx/workspace/7.excel-parser/test_doc_excel/직무전결기준표(2026.05.04).xlsx")


@pytest.mark.skipif(not JIKMU.exists(), reason="직무전결 파일 없음")
def test_jikmu_task_names_survive_as_items():
    cfg = ParserConfig(); cfg.backend = "openpyxl"
    cfg.chunk_profiles = ["delegation_rule", "note", "code_mapping"]
    chunks, _ = get_backend("openpyxl").parse(JIKMU, cfg)
    deleg = [c for c in chunks if c["chunk_type"] == "delegation_rule" and c.get("sheet") == "신한DS"]
    assert deleg, "직무전결 delegation_rule 0"
    blob = "\n".join(c["content_text"] for c in deleg)
    # 업무내용 태스크명이 항목으로 등장(구분 그룹으로 붕괴되지 않음)
    assert "연간 사업계획" in blob, "업무내용 태스크명(연간 사업계획)이 항목에서 소실"
    # 해당 태스크의 CEO 열이 header:값(CEO:○)으로 매핑됨
    assert "연간 사업계획: CEO:○" in blob.replace("  ", " "), (
        "연간 사업계획 → CEO:○ 매핑 실패"
    )


@pytest.mark.skipif(not JIKMU.exists(), reason="직무전결 파일 없음")
def test_jikmu_task_column_not_in_matrix():
    """업무내용(col3)이 matrix(역할)축이 아니라 계층축이어야 함."""
    from excel_parser_rag.pipeline import build_canvases, detect_and_classify

    cfg = ParserConfig()
    for region, canvas in detect_and_classify(build_canvases(JIKMU, cfg), cfg):
        if canvas.sheet_name == "신한DS":
            names = {str(v) for v in region.matrix_cols.values()}
            assert "업무내용" not in names, f"업무내용이 matrix 로 오분류: {sorted(names)}"
            assert 3 in region.hierarchy_cols, f"업무내용(col3) 계층축 미포함: {region.hierarchy_cols}"
            return
    pytest.fail("신한DS region 미검출")


@pytest.mark.skipif(not ASSET.exists(), reason="자산목록 파일 없음")
def test_wide_asset_matrix_not_over_absorbed():
    """자산목록 전체자산: ○ boolean 열이 데이터열 사이에 흩어진 광폭표.
    연속 역할밴드가 없으므로 라벨열 다중 흡수(과흡수)가 일어나면 안 된다."""
    from excel_parser_rag.pipeline import build_canvases, detect_and_classify

    cfg = ParserConfig()
    for region, canvas in detect_and_classify(build_canvases(ASSET, cfg), cfg):
        if canvas.sheet_name == "전체자산":
            hc = list(region.hierarchy_cols or [])
            assert len(hc) <= 2, f"전체자산 계층열 과흡수: {hc}"
            return
    pytest.fail("전체자산 region 미검출")


@pytest.mark.skipif(not JIKMU.exists(), reason="직무전결 파일 없음")
def test_full_width_banner_not_emitted_as_rule():
    """전폭 병합 섹션배너('1. 경영 관리' A5:H5)가 '비고=배너텍스트' delegation 청크로
    오출력되면 안 된다 — 병합 텍스트가 비고 열로 퍼진 echo."""
    from excel_parser_rag.textutil import compact as _c

    cfg = ParserConfig(); cfg.backend = "openpyxl"
    cfg.chunk_profiles = ["delegation_rule", "note", "code_mapping"]
    chunks, _ = get_backend("openpyxl").parse(JIKMU, cfg)
    deleg = [c for c in chunks if c["chunk_type"] == "delegation_rule"]
    # glitch signature(header:값): 그 행의 항목 텍스트가 자기 '값'(header:값 매핑) 안에 echo 됨
    # = 전폭 배너가 값 열로 퍼진 흔적. (정상 값은 ○/보고/비고문구 — 항목명을 포함하지 않음.)
    glitch = [
        c for c in deleg
        if (item := _c(str(c.get("fields", {}).get("항목", ""))))
        and item in _c(str(c.get("fields", {}).get("값", "")))
    ]
    assert not glitch, f"배너 glitch 청크 {len(glitch)}건: {[c['content_text'][:60] for c in glitch[:3]]}"
