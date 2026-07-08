"""위임전결 골든 스냅샷 회귀 가드 (merge=0, packing-immune).

병합을 끄면(delegation_merge_max_chars=0) 모든 delegation 행이 자체 청크(start_row 키)라
content 길이/병합 재packing 과 무관하게 키·개수가 안정. 계층 path 만 순수 검사한다.
"""
from __future__ import annotations

import json
import pathlib
from collections import Counter

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.config import ParserConfig

WIJUM = pathlib.Path("/Users/xxx/workspace/7.excel-parser/test_doc_excel/2-1. 위임전결기준표(2026.04.17. 개정).xlsx")
GOLDEN = pathlib.Path(__file__).parent / "fixtures" / "wijum_golden_nomerge.json"


@pytest.mark.skipif(not WIJUM.exists(), reason="위임전결 없음")
def test_wijum_golden_no_regression():
    golden = json.loads(GOLDEN.read_text(encoding="utf-8"))
    cfg = ParserConfig(); cfg.backend = "openpyxl"
    cfg.delegation_merge_max_chars = 0  # 병합 비활성 → 행당 자체 청크 → packing/컨텐츠 길이 무관
    ch, _ = get_backend("openpyxl").parse(WIJUM, cfg)
    wd = [c for c in ch if c["sheet"] == golden["sheet"]]
    from excel_parser_rag.pipeline import build_canvases, detect_and_classify
    pairs = detect_and_classify(build_canvases(WIJUM, cfg), cfg)
    assert [rg for rg, cv in pairs if cv.sheet_name == golden["sheet"]
            and rg.region_type == "hierarchical_matrix" and len(rg.hierarchy_cols or []) > 1], \
        "다중 계층열 region 없음"
    # (a) 타입별 개수 완전 동일 (merge=0 → 길이 무관, r387 교정은 골든에 이미 반영).
    counts = dict(Counter(c["chunk_type"] for c in wd))
    assert counts == golden["counts"], f"타입별 개수 회귀: {counts} vs {golden['counts']}"
    # (b) 모든 행 prefix-containment (merge=0 → 전 행 자체 키, 소실 없음). header:값 은 path 불변.
    cur = {}
    for c in wd:
        if c["chunk_type"] not in ("delegation_rule", "matrix_fact", "hierarchy_node", "note", "table_row"):
            continue
        r = (c.get("source") or {}).get("start_row")
        if r:
            cur.setdefault(f"{r}:{c['chunk_type']}", c.get("path") or [])
    regressions = []
    for k, gstr in golden["paths"].items():
        g = gstr.split(" > ") if gstr else []
        n = cur.get(k)
        if n is None:
            regressions.append((k, "청크 소실")); continue
        if g == n[:len(g)]:
            continue
        if (len(n) >= len(g) and g and g[:-1] == n[:len(g) - 1]
                and n[len(g) - 1] and (g[-1] in n[len(g) - 1] or n[len(g) - 1] in g[-1])):
            continue
        regressions.append((k, f"{g} → {n[:len(g) + 1]}"))
    assert not regressions, f"위임전결 회귀 {len(regressions)}건: {regressions[:8]}"
    # (c) r387 계층교정 확증(matrix_fact — 병합 무관).
    r387 = cur.get("387:matrix_fact") or []
    assert any("4 1 내지 3항" in e for e in r387), f"r387 ④ 소실: {r387}"
    assert not any("3 책준사업" in e for e in r387[:-1]), f"r387 ③ 아래 잘못중첩: {r387}"
