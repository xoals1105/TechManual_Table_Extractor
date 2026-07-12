"""투명 가로선 기반 시각적 셀 병합 단위 테스트 (한글 설치 불필요).

사용법: .venv\\Scripts\\python tools\\test_visual_merge.py
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import Workbook  # noqa: E402

from src.excel_writer import write_table  # noqa: E402
from src.hwpx_parser import parse_tables  # noqa: E402
from src.visual_merge import apply_visual_merges  # noqa: E402

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HH = "http://www.hancom.co.kr/hwpml/2011/head"
W, H = 9000, 850


def _border_fill(bf_id: int, *, top="SOLID", bottom="SOLID", left="SOLID", right="SOLID") -> str:
    return (
        f'<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        f'<hh:leftBorder type="{left}" width="0.12 mm" color="#000000"/>'
        f'<hh:rightBorder type="{right}" width="0.12 mm" color="#000000"/>'
        f'<hh:topBorder type="{top}" width="0.12 mm" color="#000000"/>'
        f'<hh:bottomBorder type="{bottom}" width="0.12 mm" color="#000000"/>'
        f"</hh:borderFill>"
    )


def _header_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<hh:head xmlns:hh="{HH}">'
        "<hh:borderFills>"
        + _border_fill(10)
        + _border_fill(11, bottom="NONE")
        + _border_fill(12, top="NONE", bottom="NONE")
        + _border_fill(13, top="NONE")
        + "</hh:borderFills>"
        "</hh:head>"
    )


def _cell(row: int, col: int, text: str, bf_id: int) -> str:
    return (
        f'<hp:tc borderFillIDRef="{bf_id}">'
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{W}" height="{H}"/>'
        f"<hp:subList><hp:p><hp:run><hp:t>{text}</hp:t></hp:run></hp:p></hp:subList>"
        f"</hp:tc>"
    )


def _section_xml() -> str:
    rows = []
    data = [
        ["A1", "B1", "C1", "D1"],
        ["A2", "B2", "ㅇㅇ", "D2"],
        ["A3", "B3", "ㄴㄴ", "D3"],
        ["A4", "B4", "ㅁㅁ", "D4"],
        ["A5", "B5", "C5", "D5"],
        ["A6", "B6", "C6", "D6"],
        ["A7", "B7", "C7", "D7"],
    ]
    border_ids = [
        [10, 10, 10, 10],
        [10, 10, 11, 10],
        [10, 10, 12, 10],
        [10, 10, 13, 10],
        [10, 10, 10, 10],
        [10, 10, 10, 10],
        [10, 10, 10, 10],
    ]
    for r, (texts, bf_row) in enumerate(zip(data, border_ids)):
        cells = "".join(_cell(r, c, t, bf) for c, (t, bf) in enumerate(zip(texts, bf_row)))
        rows.append(f"<hp:tr>{cells}</hp:tr>")

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="{HP}">'
        f'<hp:p><hp:run><hp:tbl rowCnt="7" colCnt="4">{"".join(rows)}</hp:tbl></hp:run></hp:p>'
        "</hs:sec>"
    )


def _build_sample_hwpx(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/header.xml", _header_xml())
        z.writestr("Contents/section0.xml", _section_xml())


def main() -> None:
    sample = Path("data/test_visual_merge.hwpx")
    _build_sample_hwpx(sample)

    raw = parse_tables(sample)[0]
    table = apply_visual_merges(raw)
    assert table.n_rows == 7 and table.n_cols == 4

    col2 = [c for c in table.cells if c.col == 2 and c.col_span == 1]
    visual = [c for c in col2 if c.row == 1 and c.row_span == 3]
    assert len(visual) == 1, f"3열 병합 실패: {[(c.row, c.row_span, c.text) for c in col2]}"
    assert visual[0].text == "ㅇㅇ\nㄴㄴ\nㅁㅁ", f"병합 텍스트 오류: {visual[0].text!r}"

    separate = [c for c in col2 if c.row in (0, 4, 5, 6)]
    assert len(separate) == 4, "가로선 있는 행은 분리 유지되어야 함"

    # 기존 rowSpan 병합이 깨지지 않는지 (샘플 hwpx 기존 테스트와 동일 구조는 아니지만 회귀 방지용)
    raw_existing = parse_tables(Path("data/sample_manual.hwpx")) if Path("data/sample_manual.hwpx").exists() else None
    if raw_existing:
        first = apply_visual_merges(raw_existing[0])
        merged = [(c.row, c.col, c.row_span, c.col_span) for c in first.cells
                  if c.row_span > 1 or c.col_span > 1]
        assert (1, 3, 2, 1) in merged and (3, 0, 1, 2) in merged, f"기존 병합 유지 실패: {merged}"

    out = Path("output/test_visual_merge.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    write_table(ws, table, include_title=False)
    wb.save(out)

    print("시각적 병합 OK:", visual[0].text.replace("\n", " / "))
    print("엑셀 저장 OK:", out)


if __name__ == "__main__":
    main()
