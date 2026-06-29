"""게이트 검증 요약 — 추출 실패(ref/header_leak/empty_header) + 나란히2표.

설계: docs/superpowers/specs/2026-06-29-excel-gate-postparse-design.md
백엔드(openpyxl/kordoc) 무관하게 동작: 원시 셀(openpyxl)로 구조/참조,
실제 파싱 chunks 로 헤더누수를 판정한다.
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List

import openpyxl
from openpyxl.utils import get_column_letter

from ..config import ParserConfig
from ..pipeline import build_canvases, detect_and_classify

ERROR_RE = re.compile(r"#(REF|VALUE|DIV/0|N/A|NAME\?|NULL|NUM)!?")


def _header_labels(region, canvas) -> Dict[int, str]:
    """region 헤더행의 {col: label}. header_rows 없으면 빈 dict."""
    out: Dict[int, str] = {}
    for hr in (region.header_rows or []):
        for col in range(region.min_col, region.max_col + 1):
            cell = canvas.cells.get((hr, col))
            # CellNode 필드명: display_value / normalized_value / logical_value (cell_node.py).
            # 병합·복원 셀까지 잡으려면 logical_value 우선.
            val = "" if cell is None else ("" if cell.is_empty else str(getattr(cell, "logical_value", "") or cell.display_value or cell.normalized_value or "").strip())
            if val and col not in out:
                out[col] = val
    return out


def compute_gate_summary(input_path, chunks: List[Dict[str, Any]]) -> Dict[str, Any]:
    path = Path(input_path)
    cfg = ParserConfig()
    canvases = build_canvases(path, cfg)
    region_pairs = detect_and_classify(canvases, cfg)

    # region 을 시트별로 묶기
    by_sheet: Dict[str, list] = defaultdict(list)
    for region, canvas in region_pairs:
        by_sheet[canvas.sheet_name].append((region, canvas))

    # 원시 워크북(참조오류 스캔용)
    wb = openpyxl.load_workbook(path, data_only=True)

    sheets_out: List[Dict[str, Any]] = []
    for ws in wb.worksheets:
        findings: List[Dict[str, Any]] = []

        # 1) ref_error — 모든 셀 스캔
        ref_cells: List[str] = []
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and ERROR_RE.search(cell.value):
                    ref_cells.append(f"{get_column_letter(cell.column)}{cell.row}")
        if ref_cells:
            findings.append({"code": "ref_error", "cells": ref_cells[:20],
                             "detail": f"참조 오류가 값에 포함됨: {', '.join(ref_cells[:5])}"})

        # 2)~3) side_by_side / empty_header — region 헤더 기반
        for region, canvas in by_sheet.get(ws.title, []):
            labels = _header_labels(region, canvas)
            if not labels:
                continue
            counts = Counter(labels.values())
            dups = [lab for lab, n in counts.items() if n > 1]
            if dups:
                dup_cells = [f"{get_column_letter(col)}{region.header_rows[0]}"
                             for col, lab in labels.items() if lab in dups]
                findings.append({"code": "side_by_side", "cells": sorted(dup_cells),
                                 "detail": f"헤더 라벨 중복({', '.join(dups)}) — 나란히 놓인 두 표로 판단"})
            # empty_header: 사용 열에 헤더 라벨이 비어있는 칸
            # 임계치: 전체 region 열 중 빈 헤더 비율이 50% 초과인 경우만 플래그
            # (trailing blank columns 오탐 방지 — 뒤쪽 빈 열은 무시)
            total_cols = region.max_col - region.min_col + 1
            empty_cols = [get_column_letter(col) + str(region.header_rows[0])
                          for col in range(region.min_col, region.max_col + 1) if col not in labels]
            empty_ratio = len(empty_cols) / total_cols if total_cols > 0 else 0
            if empty_cols and empty_ratio > 0.5 and len(labels) < 2:
                findings.append({"code": "empty_header", "cells": empty_cols[:20],
                                 "detail": f"헤더 컬럼명이 비어있음: {', '.join(empty_cols[:5])}"})

        # 4) header_leak — chunk 의 field[k]==k (헤더행이 데이터로 추출됨)
        for c in chunks:
            if c.get("sheet") != ws.title:
                continue
            fields = c.get("fields") or {}
            if not isinstance(fields, dict) or len(fields) < 2:
                continue
            same = sum(1 for k, v in fields.items()
                       if isinstance(v, str) and v.strip() == str(k).strip() and v.strip() != "")
            if same >= max(2, (len(fields) + 1) // 2):
                src = c.get("source") or {}
                row = src.get("start_row")
                loc = [f"row{row}"] if row else []
                findings.append({"code": "header_leak", "cells": loc,
                                 "detail": "헤더행이 데이터로 추출됨(헤더=값)"})
                break  # 시트당 1건이면 충분

        sheets_out.append({"sheet": ws.title, "ok": not findings, "findings": findings})

    wb.close()
    return {"ok": all(s["ok"] for s in sheets_out), "sheets": sheets_out}
