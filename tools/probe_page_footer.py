"""다중 구역 꼬리말 + 표 문서 생성 후 COM 읽기 검증."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import win32com.client as win32

from src.hwp_com_reader import _footer_at_table

OUT = Path("data/test_mp_footer.hwpx").resolve()


def build() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        hwp.XHwpWindows.Item(0).Visible = False
        hwp.HAction.Run("FileNew")
        act, ps = hwp.HAction, hwp.HParameterSet

        # footer A from body
        hwp.HAction.Run("MoveDocBegin")
        hwp.HAction.Run("Footer")
        act.GetDefault("InsertText", ps.HInsertText.HSet)
        ps.HInsertText.Text = "FOOTER_A"
        act.Execute("InsertText", ps.HInsertText.HSet)
        hwp.HAction.Run("CloseEx")

        act.GetDefault("TableCreate", ps.HTableCreation.HSet)
        ps.HTableCreation.Rows = 2
        ps.HTableCreation.Cols = 2
        act.Execute("TableCreate", ps.HTableCreation.HSet)
        for i, t in enumerate(["T1A", "T1B", "T1C", "T1D"]):
            act.GetDefault("InsertText", ps.HInsertText.HSet)
            ps.HInsertText.Text = t
            act.Execute("InsertText", ps.HInsertText.HSet)
            if i < 3:
                hwp.HAction.Run("TableRightCell")
        hwp.HAction.Run("Cancel")
        hwp.HAction.Run("MoveDocEnd")
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPage")
        try:
            hwp.HAction.Run("BreakSection")
            print("BreakSection ok")
        except Exception as e:
            print("BreakSection fail", e)

        # footer B
        hwp.HAction.Run("Footer")
        try:
            hwp.HAction.Run("SelectAll")
            hwp.HAction.Run("Delete")
        except Exception:
            pass
        act.GetDefault("InsertText", ps.HInsertText.HSet)
        ps.HInsertText.Text = "FOOTER_B"
        act.Execute("InsertText", ps.HInsertText.HSet)
        hwp.HAction.Run("CloseEx")

        act.GetDefault("TableCreate", ps.HTableCreation.HSet)
        ps.HTableCreation.Rows = 2
        ps.HTableCreation.Cols = 2
        act.Execute("TableCreate", ps.HTableCreation.HSet)
        for i, t in enumerate(["T2A", "T2B", "T2C", "T2D"]):
            act.GetDefault("InsertText", ps.HInsertText.HSet)
            ps.HInsertText.Text = t
            act.Execute("InsertText", ps.HInsertText.HSet)
            if i < 3:
                hwp.HAction.Run("TableRightCell")

        print("PageCount", hwp.PageCount)
        hwp.SaveAs(str(OUT), "HWPX", "")
        print("saved", OUT)
    finally:
        hwp.Quit()


def verify() -> None:
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        hwp.XHwpWindows.Item(0).Visible = False
        print("open", hwp.Open(str(OUT), "HWPX", "forceopen:true"))
        ids = []
        ctrl = hwp.HeadCtrl
        while ctrl:
            ids.append(ctrl.CtrlID)
            if ctrl.CtrlID in ("foot", "head", "secd"):
                print(ctrl.CtrlID, repr(ctrl.UserDesc))
            ctrl = ctrl.Next
        print("ids", ids)

        tables = []
        ctrl = hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == "tbl":
                tables.append(ctrl)
            ctrl = ctrl.Next
        for i, tc in enumerate(tables):
            text = _footer_at_table(hwp, tc)
            print(f"table{i} footer=", repr(text))
    finally:
        hwp.Quit()


if __name__ == "__main__":
    build()
    verify()
