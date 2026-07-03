"""COM 리더 검증용 테스트 .hwp 문서 생성 (한글 프로그램 필요).

제목 문단 + 4열 표(헤더: 품명/규격/수량/비고)를 가진 .hwp를 만든다.
사용법: python tools/make_test_hwp.py [출력경로]
"""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client as win32


def insert_text(hwp, text: str) -> None:
    act, ps = hwp.HAction, hwp.HParameterSet
    act.GetDefault("InsertText", ps.HInsertText.HSet)
    ps.HInsertText.Text = text
    act.Execute("InsertText", ps.HInsertText.HSet)


def main() -> None:
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("data/test_com_doc.hwp")
    out = out.resolve()
    out.parent.mkdir(parents=True, exist_ok=True)

    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        hwp.XHwpWindows.Item(0).Visible = False

        insert_text(hwp, "3.1 수리부속 현황")
        hwp.HAction.Run("BreakPara")
        insert_text(hwp, "표 3-1 부품 목록")
        hwp.HAction.Run("BreakPara")

        act, ps = hwp.HAction, hwp.HParameterSet
        act.GetDefault("TableCreate", ps.HTableCreation.HSet)
        ps.HTableCreation.Rows = 3
        ps.HTableCreation.Cols = 4
        ps.HTableCreation.WidthType = 2   # 단에 맞춤
        ps.HTableCreation.HeightType = 0
        act.Execute("TableCreate", ps.HTableCreation.HSet)

        cells = ["품명", "규격", "수량", "비고",
                 "볼트", "M8x20", "4", "공용",
                 "너트", "M8", "4", "-"]
        for i, text in enumerate(cells):
            if text:
                insert_text(hwp, text)
            if i < len(cells) - 1:
                hwp.HAction.Run("TableRightCell")

        if not hwp.SaveAs(str(out), "HWP", ""):
            raise RuntimeError("저장 실패")
        print(f"테스트 hwp 생성 완료: {out}")
    finally:
        hwp.Quit()


if __name__ == "__main__":
    main()
