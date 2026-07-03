"""추출된 표를 원본 구조(병합/크기/제목)를 유지하며 엑셀로 저장."""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from .models import MatchResult, Table

_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

TABLE_START_ROW_WITH_TITLE = 3  # 1행: 제목, 2행: 여백, 3행부터 표
TABLE_START_ROW_NO_TITLE = 1    # 제목 없이 표만 → 1행부터


def hwpunit_to_pt(v: int) -> float:
    """HWPUNIT(1/7200 inch) → 포인트(1/72 inch). 행 높이용."""
    return v / 100


def hwpunit_to_col_width(v: int) -> float:
    """HWPUNIT → 엑셀 열 너비(문자 폭 단위) 근사 변환.

    1) HWPUNIT → 인치 → 96dpi 픽셀
    2) 엑셀 열 너비 = (픽셀 - 5) / 7  (기본 폰트 기준 근사식)
    """
    px = v / 7200 * 96
    return max((px - 5) / 7, 2.0)


def _safe_sheet_name(name: str, used: set[str]) -> str:
    """엑셀 시트명 제약(31자, 금지문자) 처리 및 중복 방지."""
    name = re.sub(r"[\\/*?:\[\]]", "_", name).strip() or "표"
    name = name[:28]
    candidate, n = name, 1
    while candidate in used:
        n += 1
        candidate = f"{name}_{n}"
    used.add(candidate)
    return candidate


def write_table(ws, table: Table, *, include_title: bool = True) -> None:
    """워크시트 하나에 표 하나를 기록 (병합 + 값 + 크기, include_title 시 1행에 제목)."""
    start_row = TABLE_START_ROW_WITH_TITLE if include_title else TABLE_START_ROW_NO_TITLE

    if include_title:
        title = table.title()
        if title:
            tc = ws.cell(row=1, column=1, value=title)
            tc.font = Font(bold=True, size=12)

    col_widths: dict[int, int] = {}
    row_heights: dict[int, int] = {}

    for c in table.cells:
        r = start_row + c.row
        col = 1 + c.col

        cell = ws.cell(row=r, column=col, value=c.text if c.text else None)
        cell.border = _BORDER
        cell.alignment = Alignment(wrap_text=True, vertical="center")

        if c.row_span > 1 or c.col_span > 1:
            end_r, end_c = r + c.row_span - 1, col + c.col_span - 1
            ws.merge_cells(start_row=r, start_column=col, end_row=end_r, end_column=end_c)
            # 병합 영역 전체에 테두리 적용
            for rr in range(r, end_r + 1):
                for cc in range(col, end_c + 1):
                    ws.cell(row=rr, column=cc).border = _BORDER

        # 병합되지 않은 셀 크기만 열너비/행높이 산정에 사용 (병합 셀 크기로 왜곡 방지)
        if c.col_span == 1 and c.width > 0:
            col_widths[col] = max(col_widths.get(col, 0), c.width)
        if c.row_span == 1 and c.height > 0:
            row_heights[r] = max(row_heights.get(r, 0), c.height)

    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = hwpunit_to_col_width(w)
    for r, h in row_heights.items():
        ws.row_dimensions[r].height = hwpunit_to_pt(h)


def write_workbook(
    results: list[MatchResult],
    out_path: str | Path,
    *,
    include_title: bool = True,
) -> Path:
    """매칭된 표들을 시트별로 담은 엑셀 파일 생성. 표 1개 = 시트 1개."""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)
    used_names: set[str] = set()

    for res in results:
        sheet_name = _safe_sheet_name(f"{res.rule_name}_{res.table.index}", used_names)
        ws = wb.create_sheet(sheet_name)
        write_table(ws, res.table, include_title=include_title)

    if not wb.sheetnames:  # 매칭된 표가 없어도 빈 파일 대신 안내 시트 생성
        ws = wb.create_sheet("결과없음")
        ws.cell(row=1, column=1, value="선별 조건에 맞는 표를 찾지 못했습니다.")

    wb.save(out_path)
    return out_path
