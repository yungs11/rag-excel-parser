"""pytest 공용 fixture.

- fixture xlsx 12종을 tmp_path_factory 에 세션 스코프로 생성
- 파이프라인 단계별 헬퍼 (canvases_for / regions_for / parse_chunks)
- 다른 에이전트 모듈이 미완성일 수 있으므로 excel_parser_rag 의 파이프라인 모듈은
  fixture 내부에서 lazy import 한다 (collection 단계 실패 방지).
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any, Dict

import pytest

TESTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = TESTS_DIR.parent
for _p in (str(PROJECT_ROOT), str(TESTS_DIR)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import fixture_builders  # noqa: E402  (tests 디렉토리 sys.path 등록 후 import)

REAL_EXCEL_PATH = PROJECT_ROOT / "2-1. 위임전결기준표(2026.04.17. 개정).xlsx"


# ---------------------------------------------------------------------------
# 테스트 모듈 공용 헬퍼 함수 (`from conftest import ...`)
# ---------------------------------------------------------------------------

def cell_text(cell) -> str:
    """CellNode 의 의미값을 우선순위(logical > normalized > display > raw)로 반환."""
    for value in (cell.logical_value, cell.normalized_value, cell.display_value):
        if value:
            return str(value)
    return "" if cell.raw_value is None else str(cell.raw_value)


def chunk_blob(chunk: Dict[str, Any]) -> str:
    """chunk dict 전체를 substring 검색 가능한 문자열로 직렬화."""
    return json.dumps(chunk, ensure_ascii=False, default=str)


def main_region(pairs, sheet_name: str | None = None):
    """(Region, SheetCanvas) 쌍 목록에서 가장 큰(본문) region 을 반환."""
    candidates = [
        (region, canvas)
        for region, canvas in pairs
        if sheet_name is None or canvas.sheet_name == sheet_name
    ]
    assert candidates, "탐지된 region 이 하나도 없음"
    return max(candidates, key=lambda rc: rc[0].row_count * rc[0].col_count)[0]


# ---------------------------------------------------------------------------
# 세션 스코프 fixture
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def fixture_paths(tmp_path_factory) -> Dict[str, Path]:
    """12종 fixture xlsx 를 생성하고 {파일명: 경로} 를 반환."""
    base = tmp_path_factory.mktemp("xlsx_fixtures")
    paths: Dict[str, Path] = {}
    for name, builder in fixture_builders.FIXTURE_BUILDERS.items():
        paths[name] = builder(base / name)
    return paths


@pytest.fixture(scope="session")
def canvases_for(fixture_paths):
    """이름 → build_canvases 결과(list[SheetCanvas]) 캐시 접근자."""
    from excel_parser_rag.config import ParserConfig
    from excel_parser_rag.pipeline import build_canvases

    cache: Dict[str, Any] = {}

    def _get(name: str):
        if name not in cache:
            cache[name] = build_canvases(fixture_paths[name], ParserConfig())
        return cache[name]

    return _get


@pytest.fixture(scope="session")
def regions_for(fixture_paths):
    """이름 → detect_and_classify 결과(list[(Region, SheetCanvas)]) 캐시 접근자."""
    from excel_parser_rag.config import ParserConfig
    from excel_parser_rag.pipeline import build_canvases, detect_and_classify

    cache: Dict[str, Any] = {}

    def _get(name: str):
        if name not in cache:
            config = ParserConfig()
            canvases = build_canvases(fixture_paths[name], config)
            cache[name] = detect_and_classify(canvases, config)
        return cache[name]

    return _get


@pytest.fixture(scope="session")
def parse_chunks(fixture_paths):
    """이름 → ExcelRagParser().parse 결과(list[dict]) 캐시 접근자."""
    from excel_parser_rag import ExcelRagParser

    cache: Dict[str, Any] = {}

    def _parse(name: str):
        if name not in cache:
            cache[name] = ExcelRagParser().parse(fixture_paths[name])
        return cache[name]

    return _parse


@pytest.fixture(scope="session")
def real_excel_path() -> Path:
    if not REAL_EXCEL_PATH.exists():
        pytest.skip("실제 위임전결기준표 엑셀 파일이 없어 스킵")
    return REAL_EXCEL_PATH
