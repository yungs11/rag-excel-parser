"""Matrix / hierarchical matrix 파싱 검증 (SoT §13, §14).

핵심: matrix_fact 개수 = marker 셀 개수, 행축x열축 content_text,
marker 정규화, 계층 path 복원.
"""

from __future__ import annotations

import fixture_builders as fb
from conftest import chunk_blob

ROW_LEVEL_TYPES = ("matrix_fact", "delegation_rule", "table_row")


class TestMatrixFacts:
    def test_matrix_fact_count_equals_marker_count(self, parse_chunks):
        """SoT §14.2 — marker/값 있는 본문 셀만 matrix_fact 가 된다."""
        chunks = parse_chunks("04_matrix_table.xlsx")
        facts = [c for c in chunks if c["chunk_type"] == "matrix_fact"]
        assert len(facts) == fb.MATRIX_04_MARKER_COUNT, (
            f"marker {fb.MATRIX_04_MARKER_COUNT}개 → matrix_fact "
            f"{fb.MATRIX_04_MARKER_COUNT}개 기대, 실제 {len(facts)}개: "
            f"{[(c['range'], c['content_text'][:40]) for c in facts]}"
        )

    def test_fact_axes_in_content_text(self, parse_chunks):
        """행축('계정 삭제')과 열축('본부장')이 content_text 에 함께 나타나야 함."""
        chunks = parse_chunks("04_matrix_table.xlsx")
        target = [
            c for c in chunks
            if c["chunk_type"] == "matrix_fact" and "계정 삭제" in chunk_blob(c)
        ]
        assert target, "계정 삭제 행의 matrix_fact 미생성"
        assert any("본부장" in c["content_text"] for c in target), (
            f"열축(본부장) 누락: {[c['content_text'] for c in target]}"
        )
        assert any(c["facts"] for c in target), "matrix_fact 의 facts 가 비어 있음"

    def test_marker_normalization(self, parse_chunks):
        """SoT §14.3 — '○' 은 applicable 로 정규화 (또는 한국어 라벨 '해당')."""
        chunks = parse_chunks("04_matrix_table.xlsx")
        target = [
            c for c in chunks
            if c["chunk_type"] == "matrix_fact" and "계정 생성" in chunk_blob(c)
        ]
        assert target, "계정 생성 행의 matrix_fact 미생성"
        ok = any(
            "applicable" in chunk_blob(c) or "해당" in c["content_text"]
            for c in target
        )
        assert ok, (
            f"'○' marker 정규화(applicable/해당) 미반영: "
            f"{[c['content_text'] for c in target]}"
        )

    def test_table_summary_exists(self, parse_chunks):
        """SoT Rule 5 — 표 전체 요약 chunk 도 함께 생성."""
        chunks = parse_chunks("04_matrix_table.xlsx")
        assert any(c["chunk_type"] == "table_summary" for c in chunks), (
            f"table_summary 미생성. 생성된 타입: {sorted({c['chunk_type'] for c in chunks})}"
        )

    def test_fact_source_rows_within_body(self, parse_chunks):
        chunks = parse_chunks("04_matrix_table.xlsx")
        for c in chunks:
            if c["chunk_type"] != "matrix_fact":
                continue
            start_row = c["source"].get("start_row")
            assert start_row is not None, f"matrix_fact source.start_row 누락: {c['source']}"
            assert 2 <= start_row <= 6, (
                f"matrix_fact 의 source 행이 본문(2~6행) 밖: {start_row} ({c['range']})"
            )


class TestHierarchyPath:
    def test_path_restored_through_levels(self, parse_chunks):
        """SoT §13 — '1. 업무전반 > 가. 직원의 출장 > (1) 시외 출장' path 복원."""
        chunks = parse_chunks("05_hierarchical_matrix.xlsx")
        target = [
            c for c in chunks
            if c["chunk_type"] in ROW_LEVEL_TYPES and "시외 출장" in chunk_blob(c)
        ]
        assert target, (
            f"시외 출장 행 chunk 미생성. 생성된 타입: "
            f"{sorted({c['chunk_type'] for c in chunks})}"
        )
        restored = False
        for c in target:
            joined = " > ".join(str(p) for p in c["path"])
            if "업무전반" in joined and "직원의 출장" in joined and "시외 출장" in joined:
                restored = True
                break
        assert restored, (
            f"계층 path 복원 실패. 실제 path: {[c['path'] for c in target]}"
        )

    def test_approver_in_content_text(self, parse_chunks):
        """시외 출장 marker(본부장 열)는 content_text 에 본부장이 드러나야 함."""
        chunks = parse_chunks("05_hierarchical_matrix.xlsx")
        target = [
            c for c in chunks
            if c["chunk_type"] in ("matrix_fact", "delegation_rule")
            and "시외 출장" in chunk_blob(c)
        ]
        assert target, "시외 출장 행의 matrix_fact/delegation_rule 미생성"
        assert any("본부장" in c["content_text"] for c in target), (
            f"열축(본부장) 이 content_text 에 없음: {[c['content_text'] for c in target]}"
        )

    def test_consultation_value_preserved(self, parse_chunks):
        """합의 컬럼 값('기획')이 어떤 chunk 에든 보존되어야 함."""
        chunks = parse_chunks("05_hierarchical_matrix.xlsx")
        target = [c for c in chunks if "시내 출장" in chunk_blob(c)]
        assert target, "시내 출장 행 chunk 미생성"
        assert any("기획" in chunk_blob(c) for c in target), (
            f"합의 값('기획') 유실: {[c['content_text'] for c in target]}"
        )

    def test_section_level_chunk_exists(self, parse_chunks):
        """SoT Rule 5 — 상위 항목(1. 업무전반)은 section_summary 또는 hierarchy_node 로 검색 가능해야 함."""
        chunks = parse_chunks("05_hierarchical_matrix.xlsx")
        sections = [
            c for c in chunks
            if c["chunk_type"] in ("section_summary", "hierarchy_node")
            and "업무전반" in chunk_blob(c)
        ]
        assert sections, (
            f"'1. 업무전반' 섹션 chunk 미생성. 생성된 타입: "
            f"{sorted({c['chunk_type'] for c in chunks})}"
        )
