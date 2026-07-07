import pathlib
import re

import pytest

from excel_parser_rag.gate.excel_gate import compute_gate_summary, _numbering_restart
from excel_parser_rag.pipeline import parse_excel_for_rag

ROOT = pathlib.Path("/Users/xxx/workspace")
EXCEL = ROOT / "7.excel-parser/test_doc_excel"
MARK = ROOT / "excel-parser-markitdown/test_doc_excel"


def _summ(path):
    chunks, _ = parse_excel_for_rag(str(path))
    return compute_gate_summary(path, chunks)


def _codes(summary, sheet_substr):
    # 음성 단언(예: side_by_side not in ...)에서 시트명 오타로 trivially-pass 되는 걸 막는다.
    match = [s for s in summary["sheets"] if sheet_substr in s["sheet"]]
    assert match, f"No sheet matching {sheet_substr!r}"
    return {f["code"] for f in match[0]["findings"]}


def test_side_by_side_blocks_beoplyeong():
    s = _summ(EXCEL / "신한자산신탁_외부테이터_필요사이트 정리.xlsx")
    assert s["ok"] is False
    assert "side_by_side" in _codes(s, "법령리스트")
    # 중복 라벨 셀 좌표가 보고된다
    cells = [c for f in next(x for x in s["sheets"] if "법령리스트" in x["sheet"])["findings"]
             if f["code"] == "side_by_side" for c in f["cells"]]
    assert "A2" in cells and "C2" in cells


def test_ref_error_blocks_wbs():
    s = _summ(MARK / "251210_중소형그룹사_AX추진지원_WBS_v0.1_sys.xlsx")
    assert s["ok"] is False
    wbs_codes = set().union(*[{f["code"] for f in sh["findings"]} for sh in s["sheets"]])
    assert "ref_error" in wbs_codes


def test_aws_passes():
    s = _summ(pathlib.Path("/Users/xxx/Downloads/aws_cost_estimate.xlsx"))
    assert s["ok"] is True


def test_external_sheet1_passes():
    s = _summ(EXCEL / "신한자산신탁_외부테이터_필요사이트 정리.xlsx")
    assert _codes(s, "외부데이터소스 현황") == set()


def test_jasan_access_passes():
    s = _summ(MARK / "신한자산신탁_자산목록_v20251013.xlsx")
    assert _codes(s, "접근제어 적용 대상") == set()


def test_wijum_passes_for_now():
    s = _summ(EXCEL / "2-1. 위임전결기준표(2026.04.17. 개정).xlsx")
    # 향후 고도화 전까지 통과(매트릭스 미차단)
    assert _codes(s, "위임전결") == set()


# ── ambiguous_hierarchy: 계층열에서 번호가 재시작되어 상하위 모호한 표 게이트 ──

def test_numbering_restart_helper():
    assert _numbering_restart([(2, 0), (3, 0)]) == (2, 3)        # 직무전결형(깊은 열이 같은 0으로 재시작)
    assert _numbering_restart([(1, 1), (2, 1)]) == (1, 2)        # 가.나.다.가 최상위인 문서의 재시작
    assert _numbering_restart([(2, 0), (3, 3), (4, 5)]) is None  # 위임전결형(최소레벨 strictly 증가)
    assert _numbering_restart([(2, 0)]) is None
    assert _numbering_restart([]) is None


def test_ambiguous_hierarchy_fires_jikmu():
    path = EXCEL / "직무전결기준표(2026.05.04).xlsx"
    if not path.exists():
        pytest.skip("직무전결 파일 없음")
    s = _summ(path)
    assert "ambiguous_hierarchy" in _codes(s, "신한DS")
    cells = [c for f in next(x for x in s["sheets"] if "신한DS" in x["sheet"])["findings"]
             if f["code"] == "ambiguous_hierarchy" for c in f["cells"]]
    # 헤더행(3)이 아니라 실제 번호 body 셀(B5='1. 경영 관리' / C6='1. 사업계획')을 지목
    assert "B5" in cells and "C6" in cells
    for c in cells:
        m = re.match(r"^[A-Z]+([0-9]+)$", c)
        assert m and int(m.group(1)) > 3, f"헤더행 셀 지목됨: {c}"


def test_ambiguous_hierarchy_no_false_positive_wijum():
    # 위임전결(2-1): 번호가 컬럼따라 깊어짐(최소레벨 증가) → 미발화(오탐 0)
    s = _summ(EXCEL / "2-1. 위임전결기준표(2026.04.17. 개정).xlsx")
    for sh in s["sheets"]:
        assert "ambiguous_hierarchy" not in {f["code"] for f in sh["findings"]}, sh["sheet"]


# ── side_by_side 정밀화 회귀 (index중복 OR ≥2 distinct 라벨블록 비겹침 반복만) ──

def test_side_by_side_flags_nac():
    # NAC연계: [시스템,방식] 하위컬럼이 업무망/인터넷망 두 표로 좌우 반복 → 차단
    s = _summ(MARK / "신한자산신탁_자산목록_v20251013.xlsx")
    assert "side_by_side" in _codes(s, "NAC연계")


def test_no_false_positive_jeonche_jasan():
    # 전체자산: 한 표에 HOSTNAME류 동명 컬럼 → side_by_side 아님
    s = _summ(MARK / "신한자산신탁_자산목록_v20251013.xlsx")
    assert "side_by_side" not in _codes(s, "전체자산")


def test_no_false_positive_access_matrix():
    # 접근제어_조사: 사람별 열 매트릭스(같은 라벨 인접 반복) → side_by_side 아님
    s = _summ(MARK / "신한자산신탁_자산목록_v20251013.xlsx")
    assert "side_by_side" not in _codes(s, "접근제어_조사")


def test_jasan_file_blocks_due_to_nac_only():
    # 파일 단위 차단(사용자 결정): NAC연계 때문에 파일 ok=False, 차단 시트는 NAC연계뿐.
    s = _summ(MARK / "신한자산신탁_자산목록_v20251013.xlsx")
    assert s["ok"] is False
    blocked = [sh["sheet"] for sh in s["sheets"] if not sh["ok"]]
    assert blocked == ["NAC연계"]
