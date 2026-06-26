"""WBS 병합 회귀 — 평면목록 불변 + WBS 병합 수치. 파일/CLI 없으면 skip."""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.config import ParserConfig

DOC = Path("/Users/xxx/workspace/excel-parser-markitdown/test_doc_excel")
ASSET = DOC / "신한자산신탁_자산목록_v20251013.xlsx"
WBS = DOC / "251210_중소형그룹사_AX추진지원_WBS_v0.1_sys.xlsx"


@pytest.mark.skipif(not ASSET.exists() or not shutil.which("kordoc"), reason="자산목록/CLI 없음")
def test_flat_asset_list_unchanged_by_wbs_merge():
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    chunks, _ = get_backend("kordoc").parse(ASSET, cfg)
    # 평면목록 → 점-십진번호 필드 없음 → 병합 0
    assert not any(c.get("metadata", {}).get("merged") for c in chunks)


@pytest.mark.skipif(not WBS.exists() or not shutil.which("kordoc"), reason="WBS/CLI 없음")
def test_wbs_merge_reduces_chunk_count():
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    cfg0 = ParserConfig(); cfg0.backend = "kordoc"; cfg0.kordoc_bin = "kordoc"; cfg0.kordoc_md_out = "/tmp/kordoc_md"
    cfg0.numbering_merge_max_chars = 0
    base, _ = get_backend("kordoc").parse(WBS, cfg0)
    merged, _ = get_backend("kordoc").parse(WBS, cfg)
    wbs_base = [c for c in base if c["sheet"].startswith("중소형")]
    wbs_merged = [c for c in merged if c["sheet"].startswith("중소형")]
    assert len(wbs_merged) < len(wbs_base)  # 병합으로 감소
    assert any(c.get("metadata", {}).get("merged") for c in wbs_merged)
