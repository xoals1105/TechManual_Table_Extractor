"""파서/선별기 동작 검증용 최소 구조의 샘플 .hwpx 파일 생성 스크립트.

실제 한글 문서와 동일한 OWPML 네임스페이스/표 구조(section0.xml)를 사용하므로
parse_tables() 검증에 사용할 수 있다. (한글 프로그램에서 열리는 완전한 문서는 아님)

사용법: python tools/make_sample_hwpx.py [출력경로]
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"

# 셀 크기 (HWPUNIT: 7200 = 1인치)
W, H = 9000, 850


def para(text: str) -> str:
    return f'<hp:p><hp:run><hp:t>{text}</hp:t></hp:run></hp:p>'


def cell(row: int, col: int, text: str, rowspan: int = 1, colspan: int = 1,
         width: int = W, height: int = H) -> str:
    return (
        f'<hp:tc>'
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
        f'<hp:cellSpan colSpan="{colspan}" rowSpan="{rowspan}"/>'
        f'<hp:cellSz width="{width}" height="{height}"/>'
        f'<hp:subList>{para(text)}</hp:subList>'
        f'</hp:tc>'
    )


def table(rows: list[list], row_cnt: int, col_cnt: int, caption: str = "") -> str:
    cap = ""
    if caption:
        cap = f'<hp:caption><hp:subList>{para(caption)}</hp:subList></hp:caption>'
    trs = "".join(f'<hp:tr>{"".join(cells)}</hp:tr>' for cells in rows)
    return (
        f'<hp:p><hp:run>'
        f'<hp:tbl rowCnt="{row_cnt}" colCnt="{col_cnt}">{cap}{trs}</hp:tbl>'
        f'</hp:run></hp:p>'
    )


def build_section() -> str:
    body: list[str] = []

    body.append(para("제3장 정비"))
    body.append(para("3.1 수리부속 현황"))
    body.append(para("표 3-1 부품 목록"))
    # 표 1: 부품목록표 규칙에 매칭 (헤더 4개 일치 + 제목 키워드 + 열 4개, 병합 셀 포함)
    body.append(table(
        rows=[
            [cell(0, 0, "품명"), cell(0, 1, "규격"), cell(0, 2, "수량"), cell(0, 3, "비고")],
            [cell(1, 0, "볼트"), cell(1, 1, "M8x20"), cell(1, 2, "4"), cell(1, 3, "공용", rowspan=2)],
            [cell(2, 0, "너트"), cell(2, 1, "M8"), cell(2, 2, "4")],
            [cell(3, 0, "합계", colspan=2, width=W * 2), cell(3, 2, "8"), cell(3, 3, "-")],
        ],
        row_cnt=4, col_cnt=4,
    ))

    body.append(para("3.2 점검 기준"))
    # 표 2: 정비주기표 규칙에 매칭 (캡션 사용)
    body.append(table(
        rows=[
            [cell(0, 0, "점검항목"), cell(0, 1, "주기"), cell(0, 2, "방법")],
            [cell(1, 0, "오일 상태"), cell(1, 1, "주간"), cell(1, 2, "육안 점검")],
            [cell(2, 0, "벨트 장력"), cell(2, 1, "월간"), cell(2, 2, "장력계 측정")],
        ],
        row_cnt=3, col_cnt=3,
        caption="표 3-2 정비 주기",
    ))

    body.append(para("부록 A 개정 이력"))
    # 표 3: 어떤 규칙에도 매칭되지 않아야 함
    body.append(table(
        rows=[
            [cell(0, 0, "개정번호"), cell(0, 1, "개정일자")],
            [cell(1, 0, "1"), cell(1, 1, "2026-01-15")],
        ],
        row_cnt=2, col_cnt=2,
    ))

    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="{HP}">'
        + "".join(body) +
        '</hs:sec>'
    )


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/sample_manual.hwpx")
    out.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/section0.xml", build_section())
    print(f"샘플 생성 완료: {out}")


if __name__ == "__main__":
    main()
