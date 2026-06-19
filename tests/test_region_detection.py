"""Region 탐지/분류 검증 — region 개수, 경계(bbox), region_type (SoT §9, §10)."""

from __future__ import annotations

import pytest

import fixture_builders as fb
from conftest import main_region


class TestRegionCount:
    def test_flat_table_single_region(self, regions_for):
        pairs = regions_for("01_flat_table.xlsx")
        assert len(pairs) == 1, (
            f"표 1개 시트는 region 1개여야 함 (실제 {len(pairs)}: "
            f"{[r.range_a1 for r, _ in pairs]})"
        )
        region = pairs[0][0]
        bbox = (region.min_row, region.min_col, region.max_row, region.max_col)
        assert bbox == (1, 1, 5, 4), f"region 경계 불일치: {bbox} (기대 (1,1,5,4))"
        assert region.range_a1 == "A1:D5"

    def test_two_regions_one_sheet(self, regions_for):
        pairs = regions_for("08_multiple_regions_one_sheet.xlsx")
        assert len(pairs) == 2, (
            f"빈 행 3줄로 분리된 표 2개 → region 2개여야 함 (실제 {len(pairs)}: "
            f"{[r.range_a1 for r, _ in pairs]})"
        )
        spans = sorted((r.min_row, r.max_row) for r, _ in pairs)
        assert spans == fb.MULTI_REGION_08_SPANS, (
            f"region 행 구간 불일치: {spans} (기대 {fb.MULTI_REGION_08_SPANS})"
        )

    def test_note_row_attached_to_some_region(self, regions_for):
        """※ note 행은 어느 region 에든 포함되어야 함 (SoT §9.2-6, §16)."""
        pairs = regions_for("09_notes_and_footers.xlsx")
        assert any(r.contains(fb.NOTES_09_NOTE_ROW, 1) for r, _ in pairs), (
            f"※ note 행({fb.NOTES_09_NOTE_ROW}행)이 어느 region 에도 미포함: "
            f"{[r.range_a1 for r, _ in pairs]}"
        )

    def test_messy_main_table_detected(self, regions_for):
        """messy fixture 의 본문 marker 행이 region 으로 잡혀야 함."""
        pairs = regions_for("12_messy_real_world.xlsx")
        assert pairs
        # (1) 시외 출장 marker 위치 = (8행, 5열)
        assert any(r.contains(8, 5) for r, _ in pairs), (
            f"본문 marker 행(8행 5열)을 포함하는 region 없음: "
            f"{[r.range_a1 for r, _ in pairs]}"
        )


class TestRegionClassification:
    @pytest.mark.parametrize(
        "name,allowed",
        [
            ("01_flat_table.xlsx", {"flat_table"}),
            ("02_multi_header_table.xlsx", {"multi_header_table"}),
            ("03_merged_hierarchy_table.xlsx", {"hierarchical_table", "hierarchical_matrix"}),
            ("04_matrix_table.xlsx", {"matrix_table", "hierarchical_matrix"}),
            ("05_hierarchical_matrix.xlsx", {"hierarchical_matrix"}),
            ("06_form_document.xlsx", {"form", "key_value_block"}),
            ("07_code_mapping.xlsx", {"code_mapping_table"}),
        ],
    )
    def test_region_type(self, regions_for, name, allowed):
        region = main_region(regions_for(name))
        assert region.region_type in allowed, (
            f"{name}: region_type={region.region_type!r} (기대 {sorted(allowed)})"
        )

    @pytest.mark.parametrize(
        "name",
        [
            "01_flat_table.xlsx",
            "05_hierarchical_matrix.xlsx",
            "08_multiple_regions_one_sheet.xlsx",
            "12_messy_real_world.xlsx",
        ],
    )
    def test_confidence_and_warnings_shape(self, regions_for, name):
        for region, _ in regions_for(name):
            assert 0.0 <= region.confidence <= 1.0, (
                f"{name} {region.range_a1}: confidence 범위 위반 {region.confidence}"
            )
            assert isinstance(region.warnings, list)
