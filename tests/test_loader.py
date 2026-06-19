"""Loader / SheetCanvas 검증.

핵심: 병합 셀 logical_value 복원 (SoT §7, Rule 2), 숨김 행/열/시트 기록,
수식 처리, cell feature 키 완전성 (SoT §8).
"""

from __future__ import annotations

import fixture_builders as fb
from conftest import cell_text
from excel_parser_rag.canvas.cell_node import FEATURE_KEYS


class TestCanvasBasics:
    def test_single_sheet_canvas(self, canvases_for):
        canvases = canvases_for("01_flat_table.xlsx")
        assert len(canvases) == 1, f"시트 1개 파일은 canvas 1개여야 함 (실제 {len(canvases)})"
        canvas = canvases[0]
        assert canvas.sheet_name == fb.FLAT_01_SHEET
        assert canvas.workbook_name == "01_flat_table.xlsx"
        assert canvas.max_row >= 5, f"max_row={canvas.max_row} (데이터 마지막 행 5 이상이어야)"
        assert canvas.max_col >= 4, f"max_col={canvas.max_col} (헤더 4열 이상이어야)"

    def test_header_cells_loaded(self, canvases_for):
        canvas = canvases_for("01_flat_table.xlsx")[0]
        assert cell_text(canvas.get_cell(1, 1)) == "부서"
        assert cell_text(canvas.get_cell(1, 4)) == "영업이익"

    def test_numeric_raw_value_preserved(self, canvases_for):
        canvas = canvases_for("01_flat_table.xlsx")[0]
        assert canvas.get_cell(2, 2).raw_value == 100000000, (
            "숫자 셀 raw_value 가 원본 그대로 보존되어야 함"
        )

    def test_non_empty_cells_count(self, canvases_for):
        canvas = canvases_for("01_flat_table.xlsx")[0]
        count = len(canvas.non_empty_cells())
        # 헤더 4 + 본문 4행 x 4열 = 20
        assert count == 20, f"non_empty_cells {count}개 (기대 20개: 헤더 4 + 본문 16)"

    def test_row_density(self, canvases_for):
        canvas = canvases_for("01_flat_table.xlsx")[0]
        assert canvas.row_density(1) > 0.0
        assert canvas.row_density(99) == 0.0


class TestMergedCells:
    def test_vertical_merge_logical_value_restored(self, canvases_for):
        """SoT Rule 2 — 병합 멤버 셀: raw 는 비어 있고 logical_value 가 anchor 값."""
        canvas = canvases_for("03_merged_hierarchy_table.xlsx")[0]
        member = canvas.get_cell(3, 1)  # A3 — A2:A4 병합 멤버
        assert member.is_merged, "병합 범위 내 셀의 is_merged 가 True 여야 함"
        assert member.merge_anchor == "A2", f"merge_anchor={member.merge_anchor}"
        assert member.merge_range == "A2:A4", f"merge_range={member.merge_range}"
        assert "임직원 복리후생" in member.logical_value, (
            f"병합 멤버 logical_value 복원 실패: {member.logical_value!r}"
        )
        assert not member.raw_value, (
            f"병합 멤버의 raw_value 는 원본(빈 값) 유지여야 함: {member.raw_value!r}"
        )

    def test_merge_anchor_keeps_raw_value(self, canvases_for):
        canvas = canvases_for("03_merged_hierarchy_table.xlsx")[0]
        anchor = canvas.get_cell(2, 1)  # A2
        assert anchor.raw_value == "1. 임직원 복리후생"
        assert "임직원 복리후생" in cell_text(anchor)

    def test_merge_orientation(self, canvases_for):
        canvas03 = canvases_for("03_merged_hierarchy_table.xlsx")[0]
        assert canvas03.get_cell(3, 1).merge_orientation == "vertical"

        canvas02 = canvases_for("02_multi_header_table.xlsx")[0]
        member = canvas02.get_cell(1, 3)  # C1 — B1:C1 가로 병합 멤버
        assert member.merge_orientation == "horizontal"
        assert "2025" in member.logical_value, (
            f"가로 병합 헤더 logical_value 복원 실패: {member.logical_value!r}"
        )

    def test_merged_ranges_recorded(self, canvases_for):
        canvas = canvases_for("03_merged_hierarchy_table.xlsx")[0]
        compacted = {r.replace("$", "") for r in canvas.merged_ranges}
        assert "A2:A4" in compacted, f"merged_ranges 누락: {canvas.merged_ranges}"


class TestCellFeatures:
    def test_feature_keys_complete(self, canvases_for):
        """모든 non-empty 셀은 FEATURE_KEYS 전부를 가져야 함 (cell_node 계약)."""
        canvas = canvases_for("01_flat_table.xlsx")[0]
        cells = canvas.non_empty_cells()
        assert cells
        for cell in cells:
            missing = [k for k in FEATURE_KEYS if k not in cell.features]
            assert not missing, f"{cell.address} feature 키 누락: {missing}"

    def test_marker_feature(self, canvases_for):
        canvas = canvases_for("04_matrix_table.xlsx")[0]
        marker_cell = canvas.get_cell(2, 2)  # "계정 생성" x 팀장 = ○
        assert marker_cell.features.get("is_marker") is True, (
            f"'○' 셀이 marker 로 인식되어야 함: features={marker_cell.features}"
        )

    def test_note_feature(self, canvases_for):
        canvas = canvases_for("09_notes_and_footers.xlsx")[0]
        note_cell = canvas.get_cell(fb.NOTES_09_NOTE_ROW, 1)
        assert note_cell.features.get("looks_like_note") is True, (
            f"'※' 행이 note 로 인식되어야 함: features={note_cell.features}"
        )


class TestHiddenAndFormula:
    def test_hidden_sheet_skipped(self, canvases_for):
        canvases = canvases_for("10_hidden_rows_cols.xlsx")
        names = [c.sheet_name for c in canvases]
        assert fb.HIDDEN_10_HIDDEN_SHEET not in names, (
            f"parse_hidden_sheets=False 인데 숨김 시트가 로드됨: {names}"
        )
        assert fb.HIDDEN_10_SHEET in names

    def test_hidden_rows_cols_recorded(self, canvases_for):
        canvases = canvases_for("10_hidden_rows_cols.xlsx")
        canvas = [c for c in canvases if c.sheet_name == fb.HIDDEN_10_SHEET][0]
        assert fb.HIDDEN_10_HIDDEN_ROW in canvas.hidden_rows, (
            f"숨김 행 {fb.HIDDEN_10_HIDDEN_ROW} 미기록: hidden_rows={canvas.hidden_rows}"
        )
        assert fb.HIDDEN_10_HIDDEN_COL in canvas.hidden_cols, (
            f"숨김 열 {fb.HIDDEN_10_HIDDEN_COL}(C) 미기록: hidden_cols={canvas.hidden_cols}"
        )
        assert canvas.get_cell(fb.HIDDEN_10_HIDDEN_ROW, 1).hidden_row is True

    def test_formula_fixture_loaded(self, canvases_for):
        """openpyxl 생성 파일은 수식 캐시값이 없음 — 크래시 없이 로드되고
        수식 정보 보존 또는 빈 셀 처리 중 하나여야 함 (SoT §21.2 감점 조건)."""
        canvas = canvases_for("11_formula_values.xlsx")[0]
        b4 = canvas.get_cell(4, 2)  # =B2+B3
        assert (
            canvas.contains_formula
            or (b4.formula or "").startswith("=")
            or b4.data_type == "formula"
            or b4.is_empty
        ), f"수식 셀 처리 불명: formula={b4.formula!r} data_type={b4.data_type} is_empty={b4.is_empty}"
        # 수식이 아닌 셀은 정상 로드되어야 함
        assert canvas.get_cell(2, 2).raw_value == 3000000
