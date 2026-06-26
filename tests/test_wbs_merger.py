"""wbs_merger — kordoc 십진번호 Method B 병합 단위 테스트(hand-built dict)."""
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
        "quality": {"confidence": 0.86, "review_required": False, "parser_version": "x"},
    }


def _wbs_rows():
    # 1 (내부노드: 자식 1.1,1.2 가 부모) / 1.1 (부모+leaf 자식) / 1.1.1~1.1.3 (leaf)
    return [
        _row(10, "1. GPU 리소스 산정"),
        _row(11, "1.1. AI Agent GPU 사용량 추정"),
        _row(12, "1.1.1. 시나리오 정의"),
        _row(13, "1.1.2. 트래픽 가정"),
        _row(14, "1.1.3. 모델 후보 선정"),
        _row(15, "1.2. 최적화 구성"),
        _row(16, "1.2.1. 테넌트 정의"),
    ]


def test_detects_and_merges_parent_with_leaf_children():
    out = merge_wbs_rows(_wbs_rows(), max_chars=1100)
    merged = [c for c in out if c.get("metadata", {}).get("merged")]
    assert merged, "병합 청크 없음"
    # 1.1 부모가 1.1.1~1.1.3 자식을 헤더+줄로 묶어야
    g = [c for c in merged if c["metadata"]["parent_no"] == "1.1"]
    assert len(g) == 1
    assert g[0]["metadata"]["child_nos"] == ["1.1.1", "1.1.2", "1.1.3"]
    assert "1.1.1" in g[0]["content_text"] and "1.1.3" in g[0]["content_text"]


def test_internal_node_stays_singleton():
    out = merge_wbs_rows(_wbs_rows(), max_chars=1100)
    # '1' 은 자식(1.1,1.2)이 전부 부모 → 단독(merged 아님)
    one = [c for c in out if c.get("fields", {}).get("태스크", "").startswith("1. GPU")]
    assert one and not one[0].get("metadata", {}).get("merged")
    # 병합 청크 어디에도 parent_no == "1" 은 없어야 (내부노드는 부모로 묶이지 않음)
    assert not any(c.get("metadata", {}).get("parent_no") == "1" for c in out)


def test_embedding_text_updated_and_capped():
    out = merge_wbs_rows(_wbs_rows(), max_chars=1100)
    for c in out:
        if c.get("metadata", {}).get("merged"):
            md = c["metadata"]
            assert md["core_text"] == md["embedding_text"] == c["content_text"]
            assert len(md["embedding_text"]) <= 1100


def test_self_gating_no_number_field_passthrough():
    # ID=SHT-SV-08 같은 평면목록 → 점-십진번호 필드 없음 → 병합 0
    flat = [{
        "id": f"d::A::A{r}::row", "source_file": "d.xlsx", "sheet": "자산", "range": f"A{r}:C{r}",
        "chunk_type": "table_row", "region_type": "flat_table", "title": "자산", "path": ["자산"],
        "fields": {"ID": f"SHT-SV-0{r}", "상태": "운영", "그룹사명": "신한DS"},
        "facts": [], "content_text": "x", "keywords": [],
        "source": {"file": "d.xlsx", "sheet": "자산", "range": f"A{r}:C{r}",
                   "start_row": r, "end_row": r, "start_col": 1, "end_col": 3},
        "metadata": {}, "quality": {"confidence": 0.86, "review_required": False, "parser_version": "x"},
    } for r in range(2, 8)]
    out = merge_wbs_rows(flat, max_chars=1100)
    assert out == flat  # 불변
    assert not any(c.get("metadata", {}).get("merged") for c in out)


def test_disabled_when_max_chars_zero():
    rows = _wbs_rows()
    out = merge_wbs_rows(rows, max_chars=0)
    assert out == rows


def test_cap_splits_into_parts():
    rows = _wbs_rows()
    cap = 70  # 작은 캡 → 1.1 의 자식들이 여러 part 로 분할
    out = merge_wbs_rows(rows, max_chars=cap)
    parts = [c for c in out if c.get("metadata", {}).get("merged") and c["metadata"]["parent_no"] == "1.1"]
    assert len(parts) >= 2  # 한 묶음이 아니라 여러 part 로 쪼개짐
    # pack() 계약: 단일 자식(분할 불가)은 cap 을 넘어도 그대로 둠 → '자식 1개' 이거나 'cap 이내'
    for c in parts:
        md = c["metadata"]
        assert len(md["child_nos"]) == 1 or len(md["embedding_text"]) <= cap
    assert [c["metadata"]["part_index"] for c in parts] == sorted(c["metadata"]["part_index"] for c in parts)
    assert all("1.1" in c["content_text"] for c in parts)  # 각 part 가 부모 헤더 재명시
