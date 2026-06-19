"""kordoc 백엔드 통합 테스트 (합성 xlsx + kordoc md 페어)."""
from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_parser_rag.backends import get_backend
from excel_parser_rag.backends.base import BackendError
from excel_parser_rag.config import ParserConfig
from excel_parser_rag.chunking.chunk_schema import validate_chunk_schema


def _make_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "테스트표"
    ws.merge_cells("A1:C1")
    ws.append([])  # 빈 행은 자동으로 안 들어가므로 직접 좌표 기입
    ws["A3"], ws["B3"], ws["C3"] = "항목", "담당", "비고"
    ws["A4"], ws["B4"], ws["C4"] = "작업A", "홍길동", "메모1"
    ws["A5"], ws["B5"], ws["C5"] = "작업B", "김철수", "메모2"
    # marker 매트릭스 행
    ws["A6"], ws["D6"] = "승인건", "○"
    ws["D3"] = "팀장"
    wb.save(path)


MD = """## S1

<table>
<tr><th colspan="4">테스트표</th></tr>
<tr><td>항목</td><td>담당</td><td>비고</td><td>팀장</td></tr>
<tr><td>작업A</td><td>홍길동</td><td>메모1</td><td></td></tr>
<tr><td>작업B</td><td>김철수</td><td>메모2</td><td></td></tr>
<tr><td>승인건</td><td></td><td></td><td>○</td></tr>
</table>
"""


@pytest.fixture()
def doc(tmp_path):
    xlsx = tmp_path / "테스트표.xlsx"
    md = tmp_path / "테스트표.md"
    _make_xlsx(xlsx)
    md.write_text(MD, encoding="utf-8")
    return xlsx, md


def test_kordoc_backend_basic(doc):
    xlsx, md = doc
    cfg = ParserConfig(backend="kordoc", kordoc_md_path=str(md))
    chunks, stats = get_backend("kordoc").parse(xlsx, cfg)

    assert stats["backend"] == "kordoc"
    assert len(chunks) >= 3
    # 모든 청크 좌표 + 스키마
    assert all(c["range"] for c in chunks)
    for c in chunks:
        assert not validate_chunk_schema(c), c.get("id")
    # 시트제목이 섹션으로 새지 않고 title 로 분리
    assert all(c["title"] == "테스트표" for c in chunks)


def test_kordoc_backend_rows_and_coords(doc):
    xlsx, md = doc
    cfg = ParserConfig(backend="kordoc", kordoc_md_path=str(md))
    chunks, _ = get_backend("kordoc").parse(xlsx, cfg)
    rows = {c["range"]: c for c in chunks if c["chunk_type"] == "table_row"}
    # 작업A 행이 원본 A4:* 로 좌표 복원
    a = next(c for c in chunks if c["fields"].get("항목") == "작업A")
    assert a["fields"] == {"항목": "작업A", "담당": "홍길동", "비고": "메모1"}
    assert a["source"]["start_row"] == 4


def test_kordoc_backend_compact_matrix(doc):
    xlsx, md = doc
    cfg = ParserConfig(backend="kordoc", kordoc_md_path=str(md))
    chunks, _ = get_backend("kordoc").parse(xlsx, cfg)
    mf = [c for c in chunks if c["chunk_type"] == "matrix_fact"]
    assert mf, "marker 행이 matrix_fact 로"
    m = mf[0]
    # compact: 마커를 "해당: <컬럼헤더>" 로 접음 (○ 노이즈 없음)
    assert m["fields"].get("해당") == "팀장"
    assert "○" not in m["fields"].get("해당", "")


def test_kordoc_backend_missing_md_errors(tmp_path):
    xlsx = tmp_path / "없는md.xlsx"
    _make_xlsx(xlsx)
    cfg = ParserConfig(backend="kordoc")  # md 미지정
    with pytest.raises(BackendError):
        get_backend("kordoc").parse(xlsx, cfg)
