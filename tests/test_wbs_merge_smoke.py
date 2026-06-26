"""실 WBS kordoc 스모크 — 십진번호 Method B 병합. 파일/CLI 없으면 skip."""
from __future__ import annotations

import shutil
from pathlib import Path

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.config import ParserConfig

WBS = Path("/Users/xxx/workspace/excel-parser-markitdown/test_doc_excel/"
           "251210_중소형그룹사_AX추진지원_WBS_v0.1_sys.xlsx")


@pytest.mark.skipif(not WBS.exists() or not shutil.which("kordoc"), reason="WBS 파일 또는 kordoc CLI 없음")
def test_kordoc_wbs_merges_decimal_hierarchy():
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    chunks, _ = get_backend("kordoc").parse(WBS, cfg)
    merged = [c for c in chunks if c.get("metadata", {}).get("merged")]
    assert merged, "WBS 병합 청크 0"
    # 부모 1.1 묶음 존재 + 자식 포함 + 임베딩 텍스트 갱신 + 캡
    for c in merged:
        md = c["metadata"]
        assert md["core_text"] == md["embedding_text"] == c["content_text"]
        assert len(md["embedding_text"]) <= 1100
        assert md["merged_count"] >= 2


@pytest.mark.skipif(not WBS.exists() or not shutil.which("kordoc"), reason="WBS 파일 또는 kordoc CLI 없음")
def test_kordoc_wbs_merge_disabled_via_config():
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    cfg.numbering_merge_max_chars = 0
    chunks, _ = get_backend("kordoc").parse(WBS, cfg)
    assert not any(c.get("metadata", {}).get("merged") for c in chunks)
