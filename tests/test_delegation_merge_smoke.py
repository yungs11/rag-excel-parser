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
