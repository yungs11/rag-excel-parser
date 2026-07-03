"""wbs_merger — CGH 위임 shim 동작 검증(재작성).

기존 Method B(내부노드 단독 방치)는 CGH 로 대체됐다. merge_wbs_rows 는 이제
merge_hierarchy_rows 위임 shim 이며, 내부노드는 hierarchy_node 요약청크로 발행된다.
"""
from __future__ import annotations

from excel_parser_rag.chunking.wbs_merger import merge_wbs_rows


def _row(r, num_desc, *, sheet="WBS", stage="사전 행정 처리", owner="정을용", extra_field="태스크"):
    val = num_desc  # 예: "1.1.1. 그룹사별 시나리오 정의"
    return {
        "id": f"doc::{sheet}::A{r}:G{r}::row", "source_file": "doc.xlsx", "sheet": sheet,
        "range": f"A{r}:G{r}", "chunk_type": "table_row", "region_type": "flat_table",
        "title": sheet, "path": [sheet],
        "fields": {"단계": stage, "담당자": owner, extra_field: val},
        "facts": [], "content_text": f"...{val}", "keywords": [],
        "source": {"file": "doc.xlsx", "sheet": sheet, "range": f"A{r}:G{r}",
                   "start_row": r, "end_row": r, "start_col": 1, "end_col": 7},
        "metadata": {"workbook_title": "doc", "section": None, "core_text": f"...{val}"[:900],
                     "embedding_text": f"...{val}"[:900]},
        "quality": {"confidence": 0.86, "review_required": False,
                    "parser_version": "excel-parser-rag-v1"},
    }


def _wbs_rows():
    return [
        _row(10, "1. GPU 리소스 산정"),
        _row(11, "1.1. AI Agent GPU 사용량 추정"),
        _row(12, "1.1.1. 시나리오 정의"),
        _row(13, "1.1.2. 트래픽 가정"),
        _row(14, "1.1.3. 모델 후보 선정"),
        _row(15, "1.2. 최적화 구성"),
        _row(16, "1.2.1. 테넌트 정의"),
    ]


def _hnodes(out):
    return [c for c in out if c.get("chunk_type") == "hierarchy_node"]


def _node(out, no):
    return [c for c in _hnodes(out) if c["metadata"].get("node_no") == no]


def test_internal_node_1_1_summarizes_leaf_children():
    out = merge_wbs_rows(_wbs_rows(), max_chars=1100)
    g = _node(out, "1.1")
    assert len(g) == 1
    assert g[0]["metadata"]["child_nos"] == ["1.1.1", "1.1.2", "1.1.3"]
    assert "1.1.1" in g[0]["content_text"] and "1.1.3" in g[0]["content_text"]


def test_internal_node_stays_singleton():
    # CGH 재작성: '1' 은 내부노드 → hierarchy_node 요약청크 1개(자식 1.1,1.2 아웃라인 보유).
    # (기존 Method B 는 '1' 을 단독 방치했으나 그것이 부모-자식 단절의 원인이었음)
    out = merge_wbs_rows(_wbs_rows(), max_chars=1100)
    n1 = _node(out, "1")
    assert len(n1) == 1, "내부노드 '1' 은 요약 hierarchy_node 로 1회 발행되어야"
    md = n1[0]["metadata"]
    assert md["child_nos"] == ["1.1", "1.2"], "직속 자식 아웃라인 보유"
    assert "1.1" in n1[0]["content_text"] and "1.2" in n1[0]["content_text"]
    # '1' 은 부모 없는 루트
    assert md.get("parent_no") in (None, "")
    assert "1" in (md.get("roots") or [])


def test_hierarchy_node_embedding_text_synced():
    out = merge_wbs_rows(_wbs_rows(), max_chars=1100)
    for c in _hnodes(out):
        md = c["metadata"]
        assert md["core_text"] == md["embedding_text"] == c["content_text"]


def test_self_gating_no_number_field_passthrough():
    flat = [{
        "id": f"d::A::A{r}::row", "source_file": "d.xlsx", "sheet": "자산", "range": f"A{r}:C{r}",
        "chunk_type": "table_row", "region_type": "flat_table", "title": "자산", "path": ["자산"],
        "fields": {"ID": f"SHT-SV-0{r}", "상태": "운영", "그룹사명": "신한DS"},
        "facts": [], "content_text": "x", "keywords": [],
        "source": {"file": "d.xlsx", "sheet": "자산", "range": f"A{r}:C{r}",
                   "start_row": r, "end_row": r, "start_col": 1, "end_col": 3},
        "metadata": {}, "quality": {"confidence": 0.86, "review_required": False,
                                    "parser_version": "excel-parser-rag-v1"},
    } for r in range(2, 8)]
    out = merge_wbs_rows(flat, max_chars=1100)
    assert out == flat  # 불변
    assert _hnodes(out) == []


def test_disabled_when_max_chars_zero():
    rows = _wbs_rows()
    out = merge_wbs_rows(rows, max_chars=0)
    assert out == rows


def test_cap_splits_into_parts():
    rows = _wbs_rows()
    cap = 70  # 작은 캡 → 1.1 의 자식 아웃라인이 여러 part 로 분할
    out = merge_wbs_rows(rows, max_chars=cap)
    parts = [c for c in _hnodes(out) if c["metadata"]["node_no"] == "1.1"]
    assert len(parts) >= 2  # 여러 part
    assert [c["metadata"]["part_index"] for c in parts] == sorted(
        c["metadata"]["part_index"] for c in parts)
    assert all("1.1" in c["content_text"] for c in parts)  # 각 part 부모 헤더 재명시
