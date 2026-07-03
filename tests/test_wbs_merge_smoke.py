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
    # CGH: 내부노드는 hierarchy_node 요약청크(metadata.merged=True)로 발행된다.
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    chunks, _ = get_backend("kordoc").parse(WBS, cfg)
    hnodes = [c for c in chunks if c.get("chunk_type") == "hierarchy_node"]
    assert hnodes, "WBS hierarchy_node 0 (계층 미발화)"
    # 모든 내부노드는 직속 자식 아웃라인 보유(단절 0) + 임베딩 텍스트 동기화
    for c in hnodes:
        md = c["metadata"]
        assert md.get("merged") is True
        assert md.get("child_nos"), "내부노드에 자식 아웃라인 없음(단절)"
        assert md["core_text"] == md["embedding_text"] == c["content_text"]


@pytest.mark.skipif(not WBS.exists() or not shutil.which("kordoc"), reason="WBS 파일 또는 kordoc CLI 없음")
def test_kordoc_wbs_merge_disabled_via_config():
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    cfg.numbering_merge_max_chars = 0
    chunks, _ = get_backend("kordoc").parse(WBS, cfg)
    assert not any(c.get("metadata", {}).get("merged") for c in chunks)
