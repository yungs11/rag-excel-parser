"""hierarchy_pack — greedy cap 분할 + part 번호 단위 테스트."""
from __future__ import annotations

from excel_parser_rag.chunking.hierarchy_pack import pack, assign_parts


def test_pack_splits_by_measure():
    # 각 item 길이 4("aaaa"), measure=합. cap=10 -> [2개],[2개],[1개] (구분자 없음 단순합)
    items = ["aaaa", "bbbb", "cccc", "dddd", "eeee"]
    out = pack(items, measure=lambda g: sum(len(x) for x in g), max_chars=10)
    assert [len(g) for g in out] == [2, 2, 1]


def test_pack_single_oversize_not_split():
    items = ["x" * 50, "y" * 3]
    out = pack(items, measure=lambda g: sum(len(x) for x in g), max_chars=10)
    assert out == [["x" * 50], ["y" * 3]]


def test_pack_empty():
    assert pack([], measure=lambda g: 0, max_chars=10) == []


def test_assign_parts_multi_only_default():
    # [2,1,2] -> 멀티만 카운트: total=2, 단일은 part_index=0
    subs = [["a", "b"], ["c"], ["d", "e"]]
    out = assign_parts(subs)
    assert [(len(g), pi, pt) for g, pi, pt in out] == [(2, 1, 2), (1, 0, 2), (2, 2, 2)]


def test_assign_parts_count_all():
    subs = [["a", "b"], ["c"]]
    out = assign_parts(subs, multi_only=False)
    assert [(pi, pt) for _g, pi, pt in out] == [(1, 2), (2, 2)]
