"""E2E + JSONL schema 검증.

- 12종 fixture 전부: ExcelRagParser.parse → validate_chunk_schema 위반 0건,
  content_text/source 필수, id 유일성.
- note/total 분리, code_mapping, messy real-world 시나리오.
- 실제 위임전결기준표 엑셀 존재 시 E2E 1건.
"""

from __future__ import annotations

import pytest

import fixture_builders as fb
from conftest import chunk_blob
from excel_parser_rag.chunking.chunk_schema import validate_chunk_schema


@pytest.mark.parametrize("name", fb.FIXTURE_NAMES)
def test_fixture_e2e_schema(parse_chunks, name):
    """모든 fixture: chunk 생성 + schema 위반 0건 + 필수 필드 (SoT Rule 3, §18)."""
    chunks = parse_chunks(name)
    assert chunks, f"{name}: chunk 가 1개도 생성되지 않음"

    ids = [c["id"] for c in chunks]
    dup = {i for i in ids if ids.count(i) > 1}
    assert len(ids) == len(set(ids)), f"{name}: 중복 id 존재: {sorted(dup)[:5]}"

    for c in chunks:
        violations = validate_chunk_schema(c)
        assert not violations, f"{name} chunk {c.get('id')!r}: schema 위반 {violations}"
        assert str(c["content_text"]).strip(), f"{name} {c['id']}: content_text 비어 있음"
        src = c["source"]
        assert src.get("file") == name, f"{name} {c['id']}: source.file={src.get('file')!r}"
        assert src.get("sheet"), f"{name} {c['id']}: source.sheet 누락"
        assert src.get("range"), f"{name} {c['id']}: source.range 누락"


def test_note_and_total_separated(parse_chunks):
    """SoT §12.2, §16 — 합계 행과 ※ 행은 일반 본문과 분리되어야 함."""
    chunks = parse_chunks("09_notes_and_footers.xlsx")
    notes = [c for c in chunks if c["chunk_type"] == "note"]
    assert any("해외출장" in c["content_text"] for c in notes), (
        f"※ note chunk 미생성 또는 본문 유실. note: {[c['content_text'] for c in notes]}, "
        f"생성 타입: {sorted({c['chunk_type'] for c in chunks})}"
    )
    total_marked = [
        c for c in chunks
        if c["chunk_type"] == "total_row" or c["metadata"].get("is_total") is True
    ]
    assert total_marked, (
        "합계 행이 total_row chunk 또는 metadata.is_total=True 로 분리되지 않음. "
        f"합계 관련 chunk: {[c['chunk_type'] for c in chunks if '합계' in chunk_blob(c)]}"
    )


def test_code_mapping_chunks(parse_chunks):
    """SoT §17.7 — 약어 매핑 표는 code_mapping chunk 로 추출."""
    chunks = parse_chunks("07_code_mapping.xlsx")
    mappings = [c for c in chunks if c["chunk_type"] == "code_mapping"]
    assert len(mappings) >= len(fb.CODE_07_MAPPINGS), (
        f"code_mapping {len(fb.CODE_07_MAPPINGS)}개 기대, 실제 {len(mappings)}개. "
        f"생성 타입: {sorted({c['chunk_type'] for c in chunks})}"
    )
    for code, meaning in fb.CODE_07_MAPPINGS.items():
        assert any(
            code in chunk_blob(c) and meaning in chunk_blob(c) for c in mappings
        ), f"약어 매핑 미추출: {code} → {meaning}"


def test_messy_real_world_scenario(parse_chunks):
    """messy fixture — 제목/빈 행/잔재가 섞여도 계층/note 가 살아남아야 함."""
    chunks = parse_chunks("12_messy_real_world.xlsx")
    assert len(chunks) >= 5, f"messy fixture chunk 수 부족: {len(chunks)}"
    # 계층 path: '1. 업무전반' 이 어느 chunk path 에든 복원
    assert any(
        any("업무전반" in str(p) for p in c["path"]) for c in chunks
    ), f"messy fixture 에서 섹션 path 미복원: {[c['path'] for c in chunks][:10]}"
    # ※ note 행
    assert any(
        c["chunk_type"] == "note" and "경미한" in c["content_text"] for c in chunks
    ), (
        "※ note chunk 미생성. note 후보: "
        f"{[c['content_text'] for c in chunks if '경미한' in chunk_blob(c)]}"
    )


def test_hierarchy_fixture_path(parse_chunks):
    """03 병합 계층 표 — 세로 병합 + 번호 패턴으로 path 복원 (SoT §13.5)."""
    chunks = parse_chunks("03_merged_hierarchy_table.xlsx")
    target = [c for c in chunks if "결혼" in chunk_blob(c)]
    assert target, "'(1) 결혼' 행 chunk 미생성"
    assert any(
        "복리후생" in " > ".join(str(p) for p in c["path"]) for c in target
    ), f"병합 셀 상위 path 미복원: {[c['path'] for c in target]}"


def test_stats_returned(fixture_paths):
    from excel_parser_rag import ExcelRagParser

    parser = ExcelRagParser()
    chunks, stats = parser.parse_with_stats(fixture_paths["01_flat_table.xlsx"])
    assert chunks
    assert isinstance(stats, dict) and stats, "build_stats 결과가 비어 있음"
    assert parser.last_stats == stats


# ---------------------------------------------------------------------------
# 실제 위임전결기준표 E2E (파일 존재 시)
# ---------------------------------------------------------------------------

REQUIRED_REAL_TYPES = ("delegation_rule", "matrix_fact", "code_mapping", "note", "table_summary")


def test_real_excel_e2e(real_excel_path):
    from excel_parser_rag import ExcelRagParser

    chunks = ExcelRagParser().parse(real_excel_path)
    assert len(chunks) > 500, f"실제 전결표 chunk 수 부족: {len(chunks)} (기대 > 500)"

    types = {c["chunk_type"] for c in chunks}
    for required in REQUIRED_REAL_TYPES:
        assert required in types, f"실제 전결표에서 {required} chunk 미생성. 생성 타입: {sorted(types)}"

    sheets = {c["sheet"] for c in chunks}
    assert "1.위임전결기준" in sheets, f"본표 시트 미처리: {sorted(sheets)}"

    ids = [c["id"] for c in chunks]
    assert len(ids) == len(set(ids)), "실제 전결표 chunk id 중복"

    bad = []
    for c in chunks:
        violations = validate_chunk_schema(c)
        if violations:
            bad.append((c.get("id"), violations))
    assert not bad, f"schema 위반 {len(bad)}건 (앞 5건): {bad[:5]}"
