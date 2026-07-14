"""2구역·구역별 꼬리말·표 문서 생성 후 COM 꼬리말 추출 검증."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import win32com.client as win32

from src.hwp_com_reader import _enter_footer, _footer_at_table, _read_list_paragraphs

OUT = Path("data/test_2sec_footer.hwpx").resolve()


def _put(hwp, text: str) -> None:
    act, ps = hwp.HAction, hwp.HParameterSet
    act.GetDefault("InsertText", ps.HInsertText.HSet)
    ps.HInsertText.Text = text
    act.Execute("InsertText", ps.HInsertText.HSet)


def _table(hwp, label: str) -> None:
    act, ps = hwp.HAction, hwp.HParameterSet
    act.GetDefault("TableCreate", ps.HTableCreation.HSet)
    ps.HTableCreation.Rows = 1
    ps.HTableCreation.Cols = 1
    act.Execute("TableCreate", ps.HTableCreation.HSet)
    _put(hwp, label)
    hwp.HAction.Run("CloseEx")
    hwp.HAction.Run("BreakPara")


def _section_break(hwp) -> None:
    act, ps = hwp.HAction, hwp.HParameterSet
    act.GetDefault("InsertBreak", ps.HInsertBreak.HSet)
    # 3 ≈ 구역 나누기 (버전별 차이는 아래에서 secd 개수로 확인)
    ps.HInsertBreak.BreakType = 3
    act.Execute("InsertBreak", ps.HInsertBreak.HSet)


def _set_footer(hwp, text: str) -> None:
    # 새 문서에는 꼬리말이 없을 수 있어 InsertFooter → enter 순으로 시도
    if not _enter_footer(hwp):
        for cmd in ("InsertFooter", "Footer"):
            try:
                hwp.HAction.Run(cmd)
            except Exception:
                continue
            try:
                hwp.HAction.Run("HeaderFooterModify")
            except Exception:
                pass
            if hwp.GetPos()[0] != 0:
                break
        if hwp.GetPos()[0] == 0 and not _enter_footer(hwp):
            raise RuntimeError("꼬리말 진입 실패")
    try:
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Delete")
    except Exception:
        pass
    _put(hwp, text)
    print("  footer set", repr(_read_list_paragraphs(hwp)), "pos", hwp.GetPos())
    hwp.HAction.Run("CloseEx")


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

        _table(hwp, "T1")
        _set_footer(hwp, "FOOTER_SEC1")

        _section_break(hwp)
        _table(hwp, "T2")
        _set_footer(hwp, "FOOTER_SEC2")

        print("PageCount", hwp.PageCount)
        secd = foot = tbl = 0
        c = hwp.HeadCtrl
        while c is not None:
            if c.CtrlID == "secd":
                secd += 1
            elif c.CtrlID == "foot":
                foot += 1
                print(" foot", repr(c.UserDesc))
            elif c.CtrlID == "tbl":
                tbl += 1
            c = c.Next
        print("counts secd", secd, "foot", foot, "tbl", tbl)

        hwp.SaveAs(str(OUT), "HWPX", "")
        print("saved", OUT)
    finally:
        try:
            hwp.Clear(1)
        except Exception:
            pass
        try:
            hwp.Quit()
        except Exception:
            pass


def verify() -> None:
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        hwp.XHwpWindows.Item(0).Visible = False
        assert hwp.Open(str(OUT), "HWPX", "forceopen:true")
        tables = []
        c = hwp.HeadCtrl
        while c is not None:
            if c.CtrlID == "tbl":
                tables.append(c)
            c = c.Next
        results = []
        for i, tc in enumerate(tables):
            text = _footer_at_table(hwp, tc)
            print(f"table{i} footer=", repr(text))
            results.append(text)
        if len(results) >= 2:
            if results[0] == "FOOTER_SEC1" and results[1] == "FOOTER_SEC2":
                print("PASS: 구역별 꼬리말 일치")
            else:
                print("FAIL: expected FOOTER_SEC1 / FOOTER_SEC2")
                sys.exit(1)
    finally:
        try:
            hwp.Clear(1)
        except Exception:
            pass
        try:
            hwp.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    build()
    verify()
