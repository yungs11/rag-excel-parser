"""sibling_merger — 연속 형제(delegation_rule) 병합 단위 테스트."""
from __future__ import annotations

from excel_parser_rag.chunking.chunk_schema import RagChunk
from excel_parser_rag.chunking.sibling_merger import merge_sibling_rules


def _rule(row, path, approvers, *, region_id="r1", extra=None):
    fields = {"항목": path[-1], "경로": " > ".join(path), "전결권자": list(approvers)}
    if extra:
        fields.update(extra)
    facts = [{"predicate": "전결권자", "value": a} for a in approvers]
    content = f"S1 시트에서 '{' > '.join(path)}' 항목의 전결권자는 {', '.join(approvers)}이다."
    return RagChunk(
        source_file="doc.xlsx", sheet="S1", range=f"B{row}:I{row}",
        chunk_type="delegation_rule", region_type="hierarchical_matrix",
        title="위임전결표", path=list(path), fields=fields, facts=facts,
        content_text=content, keywords=[],
        source={"file": "doc.xlsx", "sheet": "S1", "range": f"B{row}:I{row}",
                "start_row": row, "end_row": row, "start_col": 2, "end_col": 9},
        metadata={"region_id": region_id, "excel_row": row, "is_decision_row": True},
        quality={"confidence": 0.9, "review_required": False, "parser_version": "x"},
    )


def test_consecutive_siblings_merge_into_one():
    parent = ["1.업무전반", "라.출장", "(1)시외출장"]
    chunks = [
        _rule(14, parent + ["(가)팀장"], ["본부장"]),
        _rule(15, parent + ["(나)팀원"], ["팀장"]),
    ]
    out = merge_sibling_rules(chunks, max_chars=1100)
    assert len(out) == 1
    m = out[0]
    assert m.chunk_type == "delegation_rule"
    assert m.metadata["merged"] is True
    assert m.metadata["merged_count"] == 2
    assert m.path == parent  # 공유 직속 부모
    # 두 형제의 leaf 가 모두 content 에
    assert "(가)팀장" in m.content_text and "(나)팀원" in m.content_text
    # 부모 헤더가 한 번 등장
    assert m.content_text.startswith("S1 시트") or "라.출장" in m.content_text
    assert m.metadata["child_rows"] == [14, 15]


def test_single_row_group_passthrough_identity():
    c = _rule(9, ["1.업무전반", "가.처리원칙"], ["팀장"])
    out = merge_sibling_rules([c], max_chars=1100)
    assert len(out) == 1
    assert out[0] is c  # 원본 그대로 (불변)


def test_same_top_parent_siblings_merge():
    # 둘 다 직속 부모 [1.업무전반] (path[:-1] 동일) → 연속 형제로 병합
    a = _rule(9, ["1.업무전반", "가.처리원칙"], ["팀장"])
    b = _rule(16, ["1.업무전반", "마.보고"], ["팀장"])
    out = merge_sibling_rules([a, b], max_chars=1100)
    assert len(out) == 1 and out[0].metadata["merged_count"] == 2


def test_parent_boundary_breaks_run():
    # 가(부모 1.업무전반) → (가)팀장(부모 1>라>(1)) → 마(부모 1.업무전반)
    p_top = ["1.업무전반"]
    p_deep = ["1.업무전반", "라.출장", "(1)시외"]
    chunks = [
        _rule(9, p_top + ["가.처리원칙"], ["팀장"]),
        _rule(14, p_deep + ["(가)팀장"], ["본부장"]),
        _rule(16, p_top + ["마.보고"], ["팀장"]),
    ]
    out = merge_sibling_rules(chunks, max_chars=1100)
    # 가 / (가)팀장 / 마 — 직속부모가 top, deep, top 으로 연속이 아니라 각자 단일 그룹
    assert len(out) == 3
    assert all(o.metadata.get("merged") is not True for o in out)


def test_cap_splits_same_parent_into_parts():
    parent = ["1.업무전반", "사.내규"]
    chunks = [_rule(20 + i, parent + [f"({i})항목{i}"], ["팀장"]) for i in range(6)]
    # 아주 작은 캡 → 같은 부모라도 여러 part 로 분할
    out = merge_sibling_rules(chunks, max_chars=90)
    assert len(out) >= 2
    merged = [o for o in out if o.metadata.get("merged")]
    assert all(len(o.content_text) <= 90 for o in merged)
    # 각 part 가 부모 헤더로 시작(부모명 재명시)
    assert all("사.내규" in o.content_text for o in merged)
    parts = [o.metadata.get("part_index") for o in merged]
    assert parts == sorted(parts) and parts[0] == 1


def test_cross_region_never_merges():
    a = _rule(9, ["1.업무전반", "가"], ["팀장"], region_id="r1")
    b = _rule(10, ["1.업무전반", "가"], ["팀장"], region_id="r2")
    out = merge_sibling_rules([a, b], max_chars=1100)
    assert len(out) == 2  # region_id 다르면 병합 안 함


def test_non_delegation_chunks_pass_through():
    """note/matrix_fact 등 비-delegation 청크는 위치 보존하며 그대로 통과해야 함."""
    other = RagChunk(
        source_file="doc.xlsx", sheet="S1", range="A1:A1",
        chunk_type="note", region_type="hierarchical_matrix", title="위임전결표",
        path=["주석"], fields={"주석": "※ 예외"}, facts=[], content_text="주석: ※ 예외",
        keywords=[], source={"file": "doc.xlsx", "sheet": "S1", "range": "A1:A1"},
        metadata={"region_id": "r1"},
        quality={"confidence": 0.9, "review_required": False, "parser_version": "x"},
    )
    r1 = _rule(14, ["1.업무전반", "라.출장", "(1)시외"] + ["(가)팀장"], ["본부장"])
    r2 = _rule(15, ["1.업무전반", "라.출장", "(1)시외"] + ["(나)팀원"], ["팀장"])
    out = merge_sibling_rules([other, r1, r2], max_chars=1100)
    assert out[0] is other  # 비-delegation 통과 + 위치(맨 앞) 보존
    merged = [c for c in out if c.metadata.get("merged")]
    assert len(merged) == 1 and merged[0].metadata["merged_count"] == 2
