"""위임전결 골든 스냅샷 회귀 가드 — 타입별 개수 동일 + 모든 행 prefix-containment(조상 보존)."""
from __future__ import annotations

import json
import pathlib
from collections import Counter

import pytest

from excel_parser_rag.backends import get_backend
from excel_parser_rag.config import ParserConfig

WIJUM = pathlib.Path("/Users/xxx/workspace/7.excel-parser/test_doc_excel/2-1. 위임전결기준표(2026.04.17. 개정).xlsx")
GOLDEN = pathlib.Path(__file__).parent / "fixtures" / "wijum_golden.json"


@pytest.mark.skipif(not WIJUM.exists(), reason="위임전결 없음")
def test_wijum_golden_no_regression():
    golden = json.loads(GOLDEN.read_text(encoding="utf-8"))
    cfg = ParserConfig(); cfg.backend = "openpyxl"
    ch, _ = get_backend("openpyxl").parse(WIJUM, cfg)
    wd = [c for c in ch if c["sheet"] == golden["sheet"]]
    # 골든이 실제로 다중열 코드 경로를 exercise 하는지(hier_cols>1) 확인 — 아니면 가드가 공허함.
    from excel_parser_rag.pipeline import build_canvases, detect_and_classify
    pairs = detect_and_classify(build_canvases(WIJUM, cfg), cfg)
    multi = [rg for rg, cv in pairs if cv.sheet_name == golden["sheet"]
             and rg.region_type == "hierarchical_matrix" and len(rg.hierarchy_cols or []) > 1]
    assert multi, "골든 시트에 다중 계층열 region 없음 — 가드가 다중열 변경을 exercise 못함"
    cur = {}
    for c in wd:
        if c["chunk_type"] not in ("delegation_rule", "matrix_fact", "hierarchy_node", "note", "table_row"):
            continue
        r = (c.get("source") or {}).get("start_row")
        if r:
            cur.setdefault(f"{r}:{c['chunk_type']}", c.get("path") or [])

    # r387: rewire 가 사전 버그를 교정한 유일한 행. D열 circled ③(r384)·④(r387)은 형제인데
    # 구코드는 ④ 를 ③ 의 자식으로 잘못 중첩했다. 컬럼-앵커가 ④ 를 ③ 의 형제((사) 밑)로 교정
    # → path 에서 '③...' 조상이 빠지고 delegation_rule 이 1개 늘어난다. 검증된 개선(회귀 아님).
    IMPROVED_ROW = "387:"

    # (a) 타입별 개수: delegation_rule 만 +1(r387 교정), 나머지 완전 동일.
    counts = dict(Counter(c["chunk_type"] for c in wd))
    counts_expected = dict(golden["counts"])
    counts_expected["delegation_rule"] += 1
    assert counts == counts_expected, f"타입별 개수 회귀: {counts} vs {counts_expected}"

    # (b) r387 외 전 행 prefix-containment(조상 전부 순서 보존 — 중간치환·조상소실·re-root·단축 다 잡힘).
    regressions = []
    for k, gstr in golden["paths"].items():
        if k.startswith(IMPROVED_ROW):
            continue  # 검증된 교정행 — 아래서 별도 확증
        g = gstr.split(" > ") if gstr else []
        n = cur.get(k)
        if n is None:
            regressions.append((k, "청크 소실")); continue
        if g == n[:len(g)]:
            continue
        # 유일 예외: leaf-text 확장(조상 완전 일치 + leaf 상호 부분문자열).
        if (len(n) >= len(g) and g and g[:-1] == n[:len(g) - 1]
                and n[len(g) - 1] and (g[-1] in n[len(g) - 1] or n[len(g) - 1] in g[-1])):
            continue
        regressions.append((k, f"{g} → {n[:len(g) + 1]}"))
    assert not regressions, f"위임전결 회귀 {len(regressions)}건: {regressions[:8]}"

    # (c) r387 교정 확증: ④('4 1 내지 3항') 가 존재하고, 그 조상에 형제 ③('3 책준사업')가 없어야 함.
    r387 = cur.get("387:matrix_fact") or []
    assert any("4 1 내지 3항" in e for e in r387), f"r387 ④ 소실: {r387}"
    assert not any("3 책준사업" in e for e in r387[:-1]), f"r387 여전히 ③ 아래 잘못 중첩: {r387}"
