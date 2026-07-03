"""CGH (Cascade-Gated Hierarchy) 청킹 — spec §7 매트릭스 + 8 must-fix 반례 TDD.

계약: docs/specs/CGH-spec.md ★검증 계약(v2). merge_hierarchy_rows 가 진입점.
"""
from __future__ import annotations

import pytest

from excel_parser_rag.chunking.chunk_schema import validate_chunk_schema
from excel_parser_rag.chunking.hierarchy_tree import (
    detect_spine,
    merge_hierarchy_rows,
    norm_num,
    parent_of,
)


# ---------------------------------------------------------------- helpers
def _row(r, spine_val, *, sheet="WBS", spine="WBSID", label=None, extra=None):
    fields = {spine: spine_val}
    if label is not None:
        fields["TASK"] = label
    if extra:
        fields.update(extra)
    return {
        "id": f"doc::{sheet}::A{r}:G{r}::row", "source_file": "doc.xlsx", "sheet": sheet,
        "range": f"A{r}:G{r}", "chunk_type": "table_row", "region_type": "flat_table",
        "title": sheet, "path": [sheet],
        "fields": fields,
        "facts": [], "content_text": f"row {spine_val} {label or ''}".strip(), "keywords": [],
        "source": {"file": "doc.xlsx", "sheet": sheet, "range": f"A{r}:G{r}",
                   "start_row": r, "end_row": r, "start_col": 1, "end_col": 7},
        "metadata": {"workbook_title": "doc"},
        "quality": {"confidence": 0.86, "review_required": False,
                    "parser_version": "excel-parser-rag-v1"},
    }


def _nodes(out, ntype="hierarchy_node"):
    return [c for c in out if c.get("chunk_type") == ntype]


def _by_no(out, no):
    return [c for c in out if c.get("metadata", {}).get("node_no") == no
            or c.get("metadata", {}).get("parent_no") == no]


# ================================================================ MUST-FIX 1: norm_num
def test_norm_num_arabic_dotted():
    assert norm_num("1.1.1") == "1.1.1"
    assert norm_num("1") == "1"
    assert norm_num("1. GPU 리소스") == "1"
    assert norm_num("1.1. AI Agent") == "1.1"


def test_norm_num_no_digit_absorption_from_label():
    """'6.3. 1차 입찰'의 '1차' 숫자를 세그먼트로 흡수 금지 → 6.3.1 유령 방지.
    연속 세그먼트는 공백 없는 tight dot 일 때만 이어진다."""
    assert norm_num("6.3. 1차 입찰", in_spine=True) == "6.3"
    assert norm_num("6.3. 1차 입찰") == "6.3"
    assert norm_num("2.4. 3차 검토", in_spine=True) == "2.4"
    assert norm_num("가. 1번 항목", in_spine=True) == "1"  # 가 다음 공백 → 번호 종료


def test_norm_num_hangul_ordinal():
    assert norm_num("가.") == "1"
    assert norm_num("가.1") == "1.1"
    assert norm_num("나.2)") == "2.2"


def test_norm_num_compound_word_rejected():
    assert norm_num("가능성 평가") is None
    assert norm_num("착수") is None


def test_norm_num_decimal_ratio_rejected():
    # must-fix 1: 소수/비율/단위 강배제
    assert norm_num("0.30") is None
    assert norm_num("0.3007296") is None
    assert norm_num("1.5") is None          # 정수부>=1 단일소수
    assert norm_num("3.5") is None
    assert norm_num("100.0") is None


def test_norm_num_percent_and_units_rejected():
    assert norm_num("30%") is None
    assert norm_num("0.2%") is None
    assert norm_num("5억") is None
    assert norm_num("100원") is None
    assert norm_num("14시간") is None
    assert norm_num("3개") is None
    assert norm_num("5명") is None


# ================================================================ MUST-FIX 2: detect_spine gate
def test_detect_spine_finds_wbs_field():
    rows = [_row(10, "1", label="사업관리"), _row(11, "1.1", label="착수"),
            _row(12, "1.1.1", label="킥오프"), _row(13, "1.2", label="진행")]
    assert detect_spine(rows, min_ratio=0.6) == "WBSID"


def test_detect_spine_flat_id_plus_one_nested_no_op():
    # must-fix 2: 평면 ID열 + 단 1개의 depth>=2 행 → 오발화 차단
    rows = [_row(2, "1", label="a"), _row(3, "2", label="b"), _row(4, "3", label="c"),
            _row(5, "4", label="d"), _row(6, "4.1", label="e")]  # depth>=2 는 1건뿐
    assert detect_spine(rows, min_ratio=0.6) is None


def test_detect_spine_needs_two_distinct_deep_rows():
    rows = [_row(2, "1", label="a"), _row(3, "1.1", label="b"),
            _row(4, "1.2", label="c"), _row(5, "2", label="d")]  # depth>=2: 1.1, 1.2 = 2건
    assert detect_spine(rows, min_ratio=0.6) == "WBSID"


def test_detect_spine_version_column_guard():
    rows = [_row(r, v, spine="버전", label="x") for r, v in
            enumerate(["1.0", "1.1", "1.2", "2.0"], start=2)]
    assert detect_spine(rows, min_ratio=0.6) is None


def test_detect_spine_decimal_progress_no_op():
    rows = [_row(r, v, spine="진척율", label="x") for r, v in
            enumerate(["0.30", "0.15", "0.20", "0.50"], start=2)]
    assert detect_spine(rows, min_ratio=0.6) is None


def test_detect_spine_measurement_column_no_op():
    """공수(person-month) 측정열: 정수 다수 + 우연한 소수 2건(3.2, 100.2 합계).
    분기(자식 2+인 부모)가 없어 계층 아님 → None (인력관리대장 실측 오발화 회귀)."""
    vals = ["7", "1", "6", "7", "3", "3.2", "2", "6", "4", "5", "100.2"]
    rows = [_row(r, v, spine="전체공수_계획", label="x") for r, v in enumerate(vals, start=2)]
    assert detect_spine(rows, min_ratio=0.6) is None


# ================================================================ MUST-FIX 3: parent_of & phantom
def test_parent_of_longest_prefix():
    allnums = {"1", "1.1", "1.1.1"}
    assert parent_of("1.1.1", allnums) == "1.1"
    assert parent_of("1.1", allnums) == "1"
    assert parent_of("1", allnums) is None


def test_parent_of_skips_missing_prefix():
    # 결번 4.2.4 (4.2.3 없음) 이지만 4.2 존재 → 부모 4.2
    allnums = {"4", "4.2", "4.2.4"}
    assert parent_of("4.2.4", allnums) == "4.2"


def test_phantom_parent_synthesized():
    # 4.2 부재 + 4.2.4 존재 → 팬텀 4.2 합성
    rows = [_row(2, "4", label="루트"), _row(3, "4.2.4", label="고아자식"),
            _row(4, "4.2.5", label="고아자식2")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    nodes = _nodes(out)
    phantom = [c for c in nodes if c["metadata"].get("node_no") == "4.2"
               and c["metadata"].get("phantom")]
    assert phantom, "팬텀 4.2 미합성"
    # 팬텀이 4.2.4/4.2.5 를 자식으로
    assert set(phantom[0]["metadata"]["child_nos"]) == {"4.2.4", "4.2.5"}
    # 팬텀 자식이 된 실노드는 roots 에서 제외 (roots 에 4 만)
    roots = phantom[0]["metadata"].get("roots") or []
    if roots:
        assert "4.2.4" not in roots


def test_phantom_multi_level_gap():
    # 다단 결손: 5 존재, 5.1.1 존재, 5.1 결번 → 5.1 팬텀 합성
    rows = [_row(2, "5", label="루트"), _row(3, "5.1.1", label="깊은자식"),
            _row(4, "5.1.2", label="깊은자식2")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    nos = {c["metadata"].get("node_no") for c in _nodes(out)}
    assert "5.1" in nos  # 중간 결번 팬텀


def test_phantom_two_level_gap_schema_valid():
    # must-fix 3 + 7: ≥2 연속 결번 → 상위 팬텀의 직속 자식도 팬텀이라 실행 좌표를
    # 못 찾던 경로. 모든 발행 팬텀이 스키마 유효(source.file 비어있지 않음)해야.
    # 5 존재, 5.1.1.1 존재 → 5.1, 5.1.1 두 단계 팬텀 합성.
    rows = [_row(2, "5", label="루트"), _row(3, "5.1.1.1", label="아주깊은자식"),
            _row(4, "5.1.1.2", label="아주깊은자식2")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    phantoms = [c for c in _nodes(out) if c["metadata"].get("phantom")]
    pnos = {c["metadata"]["node_no"] for c in phantoms}
    assert {"5.1", "5.1.1"} <= pnos, f"다단 팬텀 미합성: {pnos}"
    for c in phantoms:
        assert c["source_file"], f"팬텀 {c['metadata']['node_no']} source_file 공백"
        assert c["source"].get("file"), f"팬텀 {c['metadata']['node_no']} source.file 공백"
        assert validate_chunk_schema(c) == [], \
            f"팬텀 {c['metadata']['node_no']} 스키마 위반: {validate_chunk_schema(c)}"


def test_phantom_two_branches_deep_gap_schema_valid():
    # 두 브랜치가 각각 다단 결번 → 다수 팬텀 발행. 전부 스키마 유효.
    rows = [_row(2, "6", label="루트"), _row(3, "6.1.1", label="a"),
            _row(4, "6.2.1", label="b"), _row(5, "6.2.2", label="c")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    for c in out:
        assert validate_chunk_schema(c) == [], \
            f"{c.get('chunk_type')} {c.get('metadata',{}).get('node_no')} 위반: {validate_chunk_schema(c)}"
    pnos = {c["metadata"]["node_no"] for c in _nodes(out) if c["metadata"].get("phantom")}
    assert {"6.1", "6.2"} <= pnos, f"브랜치 팬텀 미합성: {pnos}"


# ================================================================ MUST-FIX 4: mixed numbering
def test_mixed_numbering_conflict_discards_spine():
    # 한 spine 열에 arabic + hangul 혼용, 충돌 → no-op(폐기)
    rows = [_row(2, "1", spine="구분", label="a"), _row(3, "1.1", spine="구분", label="b"),
            _row(4, "가.", spine="구분", label="c"), _row(5, "가.1", spine="구분", label="d")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    # 충돌 폐기 → 원본 그대로 (hierarchy_node 발행 없음)
    assert _nodes(out) == []
    assert len(out) == len(rows)


def test_hangul_only_numbering_ok():
    rows = [_row(2, "가.", spine="구분", label="설계"),
            _row(3, "가.1", spine="구분", label="하위설계1"),
            _row(4, "가.2", spine="구분", label="하위설계2"),
            _row(5, "나.", spine="구분", label="구현")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    nodes = _nodes(out)
    assert any(c["metadata"].get("numbering_system") == "hangul" for c in nodes)


# ================================================================ MUST-FIX 5: pass-through
def test_pass_through_non_numbered_rows_preserved():
    rows = [
        _row(10, "1", label="사업관리"),
        _row(11, "1.1", label="착수"),
        _row(12, "1.1.1", label="킥오프"),
        _row(13, "1.1.2", label="계획"),
        # 비번호 table_row
        _row(14, "", label="비고행", extra={"메모": "특이사항"}),
    ]
    # 다른 chunk_type 도 섞어 넣음
    note = {"id": "doc::WBS::N1::note", "source_file": "doc.xlsx", "sheet": "WBS",
            "range": "A20:A20", "chunk_type": "note", "region_type": "flat_table",
            "title": "WBS", "path": ["WBS"], "fields": {}, "facts": [],
            "content_text": "주석 노트", "keywords": [],
            "source": {"file": "doc.xlsx", "sheet": "WBS", "range": "A20:A20",
                       "start_row": 20, "end_row": 20, "start_col": 1, "end_col": 1},
            "metadata": {}, "quality": {"confidence": 0.9, "review_required": False,
                                        "parser_version": "excel-parser-rag-v1"}}
    rows.append(note)
    out = merge_hierarchy_rows(rows, max_chars=1100)
    # note 그대로 보존
    assert note in out
    # 비번호 table_row 도 보존 (원본 객체 그대로)
    assert rows[4] in out


def test_no_input_chunk_dropped():
    rows = [_row(10, "1", label="a"), _row(11, "1.1", label="b"),
            _row(12, "1.1.1", label="c"), _row(13, "1.1.2", label="d")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    # 모든 leaf 의 원본 row 가 어떤 형태로든 표현되어야 (leaf 는 table_row in-place)
    leaf_rows = [c for c in out if c.get("chunk_type") == "table_row"]
    # 최소한 leaf 개수만큼 table_row 존재
    assert len(leaf_rows) >= 2


# ================================================================ MUST-FIX 6: rollup double-emit
def test_rollup_not_double_emitted():
    # '1' 은 번호 내부노드이면서 라벨이 '합계' → hierarchy_node 재발행 금지
    rows = [
        _row(10, "1", label="합계"),
        _row(11, "1.1", label="착수"),
        _row(12, "1.1.1", label="킥오프"),
        _row(13, "1.1.2", label="계획"),
    ]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    # node_no == '1' 인 hierarchy_node 는 없어야
    hnodes = [c for c in _nodes(out) if c["metadata"].get("node_no") == "1"]
    assert hnodes == [], "합계 라벨 내부노드가 hierarchy_node 로 재발행됨"
    # 대신 롤업 청크 1회
    rollup = [c for c in out if c.get("chunk_type") in ("total_row", "section_summary")]
    assert rollup, "롤업 canonical 청크 미발행"


def test_total_row_all_marker():
    rows = [
        _row(9, "ALL", label="전체 프로젝트"),
        _row(10, "1", label="사업관리"),
        _row(11, "1.1", label="착수"),
        _row(12, "1.1.1", label="킥오프"),
        _row(13, "1.1.2", label="계획"),
    ]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    total = [c for c in out if c.get("chunk_type") == "total_row"]
    assert total, "ALL total_row 미발행"


# ================================================================ MUST-FIX 7: schema 15-field
def test_all_emitted_chunks_valid_schema():
    rows = [
        _row(9, "ALL", label="전체"),
        _row(10, "1", label="사업관리"),
        _row(11, "1.1", label="착수"),
        _row(12, "1.1.1", label="킥오프"),
        _row(13, "1.1.2", label="계획"),
        _row(14, "1.2", label="진행"),
        _row(15, "1.2.1", label="주간보고"),
    ]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    for c in out:
        errs = validate_chunk_schema(c)
        assert errs == [], f"schema 위반 {c.get('chunk_type')}: {errs}"


def test_node_chunk_type_is_hierarchy_node():
    rows = [_row(10, "1", label="a"), _row(11, "1.1", label="b"),
            _row(12, "1.1.1", label="c"), _row(13, "1.1.2", label="d")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    nodes = _nodes(out)
    assert nodes, "hierarchy_node 미발행"
    for c in nodes:
        assert c["chunk_type"] == "hierarchy_node"
        assert c["quality"]["parser_version"] == "excel-parser-rag-v1"


def test_internal_node_contains_children_outline():
    rows = [_row(10, "1", label="사업관리"), _row(11, "1.1", label="착수"),
            _row(12, "1.2", label="진행"), _row(13, "1.3", label="완료")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    node1 = [c for c in _nodes(out) if c["metadata"].get("node_no") == "1"]
    assert node1
    txt = node1[0]["content_text"]
    assert "1.1" in txt and "1.2" in txt and "1.3" in txt


# ================================================================ MUST-FIX 8: leaf near-dup
def test_leaf_near_dup_suppressed():
    # 형제 leaf 공통 컨텍스트(단계/Activity/담당자) 는 breadcrumb/metadata 로,
    # 본문은 고유 라벨. leaf 는 table_row in-place + metadata 계층태그.
    common = {"단계": "사전 행정 처리", "담당자": "정을용"}
    rows = [
        _row(10, "1", label="GPU 리소스 산정", extra=common),
        _row(11, "1.1", label="사용량 추정", extra=common),
        _row(12, "1.1.1", label="시나리오 정의", extra=common),
        _row(13, "1.1.2", label="트래픽 가정", extra=common),
    ]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    leaves = [c for c in out if c.get("chunk_type") == "table_row"
              and c.get("metadata", {}).get("node_no")]
    assert leaves
    for lf in leaves:
        md = lf["metadata"]
        # 계층 태그 in-place
        assert "parent_no" in md and "depth" in md and "node_no" in md
        # breadcrumb 존재
        assert "breadcrumb" in md


# ================================================================ spec §7: 평면목록 no-op
def test_flat_list_no_op():
    rows = [_row(r, "", spine="ID", label=None, extra={"ID": f"SHT-SV-0{r}", "상태": "운영"})
            for r in range(2, 8)]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    assert out == rows
    assert _nodes(out) == []


# ================================================================ spec §7: 매트릭스 게이트
def test_matrix_gate_disables():
    matrix = {"id": "d::S::M1::mf", "source_file": "d.xlsx", "sheet": "WBS",
              "range": "A1:C1", "chunk_type": "matrix_fact", "region_type": "matrix",
              "title": "WBS", "path": ["WBS"], "fields": {}, "facts": [],
              "content_text": "matrix", "keywords": [],
              "source": {"file": "d.xlsx", "sheet": "WBS", "range": "A1:C1",
                         "start_row": 1, "end_row": 1, "start_col": 1, "end_col": 3},
              "metadata": {}, "quality": {"confidence": 0.9, "review_required": False,
                                          "parser_version": "excel-parser-rag-v1"}}
    rows = [matrix,
            _row(10, "1", label="a"), _row(11, "1.1", label="b"),
            _row(12, "1.1.1", label="c"), _row(13, "1.1.2", label="d")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    assert _nodes(out) == []  # matrix 게이트 → 비활성
    assert len(out) == len(rows)


# ================================================================ spec §7: 멀티시트 독립
def test_multi_sheet_independent():
    s1 = [_row(10, "1", sheet="WBS", label="a"), _row(11, "1.1", sheet="WBS", label="b"),
          _row(12, "1.1.1", sheet="WBS", label="c"), _row(13, "1.1.2", sheet="WBS", label="d")]
    s2 = [_row(2, "", sheet="자산", spine="ID", extra={"ID": f"X{r}", "상태": "운영"})
          for r in range(2, 6)]
    out = merge_hierarchy_rows(s1 + s2, max_chars=1100)
    # WBS 시트만 발화, 자산 시트는 no-op → 원본 보존
    for c in s2:
        assert c in out
    assert any(c["chunk_type"] == "hierarchy_node" and c["sheet"] == "WBS"
               for c in out)


# ================================================================ 깊은 트리 breadcrumb
def test_deep_tree_breadcrumb():
    rows = [_row(10, "1", label="L1"), _row(11, "1.1", label="L2"),
            _row(12, "1.1.1", label="L3"), _row(13, "1.1.1.1", label="L4a"),
            _row(14, "1.1.1.2", label="L4b")]
    out = merge_hierarchy_rows(rows, max_chars=1100)
    deep = [c for c in _nodes(out) if c["metadata"].get("node_no") == "1.1.1"]
    assert deep
    assert deep[0]["metadata"]["depth"] == 3
    # breadcrumb 조상 포함
    bc = deep[0]["metadata"].get("breadcrumb", "")
    assert "1" in bc and "1.1" in bc


# ================================================================ disabled
def test_disabled_when_max_chars_zero():
    rows = [_row(10, "1", label="a"), _row(11, "1.1", label="b"),
            _row(12, "1.1.1", label="c")]
    out = merge_hierarchy_rows(rows, max_chars=0)
    assert out == rows
