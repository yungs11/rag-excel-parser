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
def test_wbs_cgh_emits_connected_internal_nodes():
    # CGH: 병합으로 청크 수가 줄지 않고(요약노드 추가), 모든 내부노드가 자식 아웃라인
    # 을 품는다(부모-자식 단절 0). 이것이 Method B 대체의 핵심 계약.
    cfg = ParserConfig(); cfg.backend = "kordoc"; cfg.kordoc_bin = "kordoc"; cfg.kordoc_md_out = "/tmp/kordoc_md"
    merged, _ = get_backend("kordoc").parse(WBS, cfg)
    wbs_merged = [c for c in merged if c["sheet"].startswith("중소형")]
    hnodes = [c for c in wbs_merged if c.get("chunk_type") == "hierarchy_node"]
    assert hnodes, "hierarchy_node 미발화"
    # 단절 0: 모든 내부노드에 child_nos 존재
    disconnected = [c["metadata"].get("node_no") for c in hnodes
                    if not c["metadata"].get("child_nos")]
    assert disconnected == [], f"단절된 내부노드: {disconnected}"
    # 루트 '1' 이 요약노드로 존재(기존 단독 방치 → 요약 발행으로 전환)
    assert any(c["metadata"].get("node_no") == "1" for c in hnodes)
