"""ColHierarchyTracker — 열별 체인 + merge-span 부모 보호 단위 테스트."""
from __future__ import annotations

from excel_parser_rag.parsers.hierarchy_table import ColHierarchyTracker


def test_deep_column_number_does_not_pop_shallow_category():
    # col2=무번호 카테고리(병합부모), col3=번호 소분류
    t = ColHierarchyTracker([2, 3])
    t.push(2, "경영 일반")          # 무번호 카테고리
    t.push(3, "1. 사업계획")        # col3 L0
    assert t.path == ["경영 일반", "1. 사업계획"]
    t.push(3, "가. 연간 사업계획")   # col3 L1 → 1.사업계획 밑
    assert t.path == ["경영 일반", "1. 사업계획", "가. 연간 사업계획"]
    t.push(3, "2. 조직개편")        # col3 L0 → col3 체인만 리셋, col2(경영일반)는 유지
    assert t.path == ["경영 일반", "2. 조직개편"]


def test_shallow_update_clears_deeper():
    t = ColHierarchyTracker([2, 3])
    t.push(2, "경영 일반"); t.push(3, "1. 사업계획")
    t.push(2, "예산")               # col2 갱신 → col3 무효화
    assert t.path == ["예산"]


def test_within_column_numbering_cascade():
    # 위임전결형: 한 열에 여러 번호 레벨(1.→라.→(1))
    t = ColHierarchyTracker([2, 3])
    t.push(2, "1. 업무전반")        # L0
    t.push(2, "라. 직원의 출장")     # L1 (가/나/다/라)
    t.push(2, "(1) 시외 출장")       # L2
    t.push(3, "(가) 팀장")          # col3 L3
    assert t.path == ["1. 업무전반", "라. 직원의 출장", "(1) 시외 출장", "(가) 팀장"]
    t.push(2, "마. 세무관서")        # col2 L1 → (1),라 pop, 1.업무전반 유지; col3 무효화
    assert t.path == ["1. 업무전반", "마. 세무관서"]


def test_top_and_last_item():
    t = ColHierarchyTracker([2, 3])
    assert t.top == "" and t.last_item == ""
    t.push(2, "경영 일반"); t.push(3, "가.")
    assert t.top == "경영 일반" and t.last_item == "가."


def test_single_column():
    t = ColHierarchyTracker([2])
    t.push(2, "1. 감사"); t.push(2, "가. 실시")
    assert t.path == ["1. 감사", "가. 실시"]
