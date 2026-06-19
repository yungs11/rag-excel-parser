"""Header 탐지 검증 — header row depth, multi-header flatten, 컬럼 역할, override (SoT §11, §33.4, §34)."""

from __future__ import annotations

import fixture_builders as fb
from conftest import chunk_blob, main_region
from excel_parser_rag.textutil import compact


class TestHeaderRows:
    def test_flat_header_single_row(self, regions_for):
        region = main_region(regions_for("01_flat_table.xlsx"))
        assert region.header_rows == [1], f"header_rows={region.header_rows} (기대 [1])"
        assert set(region.body_rows) == {2, 3, 4, 5}, (
            f"body_rows={sorted(region.body_rows)} (기대 [2,3,4,5])"
        )

    def test_multi_header_two_rows(self, regions_for):
        region = main_region(regions_for("02_multi_header_table.xlsx"))
        assert region.header_rows == [1, 2], f"header_rows={region.header_rows} (기대 [1,2])"
        assert 3 in region.body_rows
        assert 1 not in region.body_rows and 2 not in region.body_rows

    def test_note_row_excluded_from_body(self, regions_for):
        """※ 행은 body 에서 제외 (SoT §12.1)."""
        region = main_region(regions_for("09_notes_and_footers.xlsx"))
        assert region.header_rows == [1]
        assert fb.NOTES_09_NOTE_ROW not in region.body_rows, (
            f"※ note 행이 body_rows 에 포함됨: {sorted(region.body_rows)}"
        )


class TestHeaderFlatten:
    def test_multi_header_flatten_column_names(self, parse_chunks):
        """SoT §11.4 — 다단 헤더는 '2025_매출' 식으로 합쳐진 컬럼명이 되어야 함."""
        chunks = parse_chunks("02_multi_header_table.xlsx")
        rows = [
            c for c in chunks
            if c["chunk_type"] == "table_row" and "AI팀" in chunk_blob(c)
        ]
        assert rows, "AI팀 행의 table_row chunk 미생성"
        keys = set()
        for c in rows:
            keys |= {str(k) for k in c["fields"].keys()}
        assert any("2025" in k and "매출" in k for k in keys), (
            f"상위+하위 헤더 합성 컬럼명(2025+매출) 미발견: {sorted(keys)}"
        )
        assert any("2026" in k and "매출" in k for k in keys), (
            f"상위+하위 헤더 합성 컬럼명(2026+매출) 미발견: {sorted(keys)}"
        )
        assert not any("Unnamed" in k for k in keys), f"Unnamed 컬럼 발생: {sorted(keys)}"


class TestMatrixColumns:
    def test_matrix_cols_named(self, regions_for):
        region = main_region(regions_for("04_matrix_table.xlsx"))
        assert region.matrix_cols, "matrix_cols 가 비어 있음 (marker 축 컬럼 미식별)"
        col_names = [compact(v) for v in region.matrix_cols.values()]
        for expected in ("팀장", "본부장", "대표이사"):
            assert any(expected in name for name in col_names), (
                f"matrix 컬럼 '{expected}' 미식별: {col_names}"
            )

    def test_hierarchical_matrix_columns(self, regions_for):
        region = main_region(regions_for("05_hierarchical_matrix.xlsx"))
        assert region.header_rows == [1], f"header_rows={region.header_rows} (기대 [1])"
        assert region.hierarchy_cols, "hierarchy_cols 가 비어 있음 (계층 컬럼 미식별)"
        named = [
            compact(v)
            for v in list(region.matrix_cols.values()) + list(region.metadata_cols.values())
        ]
        assert any("팀장" in n for n in named), f"전결권자 컬럼 미식별: {named}"
        assert any("합의" in n for n in named), f"합의 컬럼 미식별: {named}"


class TestOverride:
    def test_override_not_overwritten(self, fixture_paths):
        """pipeline 규약 — override 로 미리 지정된 region_type/header_rows 는
        classify_region/detect_headers 가 덮어쓰지 않는다."""
        from excel_parser_rag.config import ParserConfig, RegionOverride, SheetOverride
        from excel_parser_rag.pipeline import build_canvases, detect_and_classify

        config = ParserConfig()
        config.sheet_overrides[fb.MATRIX_04_SHEET] = SheetOverride(
            regions=[
                RegionOverride(
                    range="A1:D6",
                    region_type="matrix_table",
                    header_rows=[1],
                    hierarchy_cols=[1],
                )
            ]
        )
        canvases = build_canvases(fixture_paths["04_matrix_table.xlsx"], config)
        pairs = detect_and_classify(canvases, config)
        targets = [r for r, _ in pairs if r.contains(2, 1)]
        assert targets, f"override 범위의 region 미생성: {[r.range_a1 for r, _ in pairs]}"
        region = targets[0]
        assert region.region_type == "matrix_table", (
            f"override region_type 이 덮어써짐: {region.region_type}"
        )
        assert region.header_rows == [1], (
            f"override header_rows 가 덮어써짐: {region.header_rows}"
        )
