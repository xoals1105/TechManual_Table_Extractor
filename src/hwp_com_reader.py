"""DRM/암호화 문서 대응: 한글 프로그램(COM)으로 문서를 직접 열어 표를 읽는다.

디스크의 파일이 DRM으로 암호화되어 있어도, 열람 권한이 있는 PC에서는
한글이 문서를 복호화해서 연다. 이 모듈은 파일을 복호화하려 하지 않고,
한글이 열어 놓은 문서에서 표마다 HWPML(XML) 블록을 내보내 파싱한다.

요구 사항: 한글(Hancom Office) 설치 + pywin32
"""
from __future__ import annotations

import re
from pathlib import Path

from lxml import etree

from .models import Cell, Table

_XML_DECL = re.compile(r"^\s*<\?xml[^>]*\?>")


def read_tables_via_com(path: str | Path, visible: bool = False) -> list[Table]:
    """한글로 문서를 열고 모든 표를 문서 순서대로 읽어 반환."""
    try:
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError("pywin32가 설치되어 있지 않습니다.") from e

    path = Path(path).resolve()
    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
    except Exception:
        try:
            hwp = win32.Dispatch("HWPFrame.HwpObject")
        except Exception as e:
            raise RuntimeError(
                "한글 자동화 개체(HWPFrame.HwpObject)를 실행할 수 없습니다.\n"
                "이 PC에 '한글(정품, 편집 가능 버전)'이 설치되어 있어야 합니다.\n"
                "한컴오피스 뷰어(hwpviewer)만 설치된 PC에서는 동작하지 않습니다."
            ) from e
    try:
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        try:
            hwp.XHwpWindows.Item(0).Visible = visible
        except Exception:
            pass

        if not hwp.Open(str(path), "HWP", "forceopen:true"):
            raise RuntimeError(
                f"한글에서 문서를 열지 못했습니다: {path}\n"
                "DRM 문서라면 이 PC/계정에 열람 권한이 있는지 확인하세요."
            )

        tables: list[Table] = []
        idx = 0
        ctrl = hwp.HeadCtrl
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                table = _read_one_table(hwp, ctrl, idx)
                if table is not None:
                    tables.append(table)
                    idx += 1
            ctrl = ctrl.Next
        return tables
    finally:
        try:
            hwp.Clear(1)  # 문서 닫기 (저장 안 함)
        except Exception:
            pass
        try:
            hwp.Quit()
        except Exception:
            pass


def _read_one_table(hwp, ctrl, idx: int) -> Table | None:
    """표 컨트롤 하나를 선택하여 HWPML XML로 내보낸 뒤 파싱."""
    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
    hwp.FindCtrl()  # 표 개체 선택

    xml = hwp.GetTextFile("HWPML2X", "saveblock")
    if not xml or "<" not in xml:
        raise RuntimeError(
            f"표 #{idx}의 내용을 내보내지 못했습니다. "
            "DRM 정책이 내용 내보내기를 차단했을 수 있습니다."
        )

    table = _parse_hml_table(xml, idx)
    if table is None:
        return None
    table.preceding_texts = _preceding_paragraphs(hwp, ctrl)
    return table


# ---------------------------------------------------------------
# HWPML(HML) 파싱 — CELL 속성: RowAddr/ColAddr/RowSpan/ColSpan/Width/Height
# ---------------------------------------------------------------

def _nearest_ancestor(el, tag: str):
    p = el.getparent()
    while p is not None:
        if p.tag == tag:
            return p
        p = p.getparent()
    return None


def _para_text(p_el) -> str:
    return "".join(t for ch in p_el.iter("CHAR") for t in ch.itertext()).strip()


def _parse_hml_table(xml: str, idx: int) -> Table | None:
    cleaned = _XML_DECL.sub("", xml, count=1)
    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.fromstring(cleaned.encode("utf-8", errors="replace"), parser=parser)
    if root is None:
        return None

    # 최상위 TABLE (중첩 표 제외)
    tbl_el = next(
        (t for t in root.iter("TABLE") if _nearest_ancestor(t, "TABLE") is None),
        None,
    )
    if tbl_el is None:
        return None

    table = Table(
        index=idx,
        section=0,
        caption="",
        preceding_texts=[],
        n_rows=int(tbl_el.get("RowCount", 0)),
        n_cols=int(tbl_el.get("ColCount", 0)),
    )

    for cap in tbl_el.iter("CAPTION"):
        if _nearest_ancestor(cap, "TABLE") is tbl_el:
            table.caption = " ".join(
                _para_text(p) for p in cap.iter("P") if _para_text(p)
            ).strip()
            break

    for cell_el in tbl_el.iter("CELL"):
        if _nearest_ancestor(cell_el, "TABLE") is not tbl_el:
            continue
        lines = []
        for p in cell_el.iter("P"):
            # 중첩 표 내부 문단 제외
            if _nearest_ancestor(p, "CELL") is not cell_el:
                continue
            line = _para_text(p)
            if line:
                lines.append(line)
        table.cells.append(Cell(
            row=int(cell_el.get("RowAddr", 0)),
            col=int(cell_el.get("ColAddr", 0)),
            row_span=int(cell_el.get("RowSpan", 1)),
            col_span=int(cell_el.get("ColSpan", 1)),
            text="\n".join(lines),
            width=int(cell_el.get("Width", 0)),
            height=int(cell_el.get("Height", 0)),
        ))

    if not table.cells:
        return None

    # RowCount/ColCount 속성이 없을 때 셀 좌표로 보정
    if table.n_rows == 0:
        table.n_rows = max(c.row + c.row_span for c in table.cells)
    if table.n_cols == 0:
        table.n_cols = max(c.col + c.col_span for c in table.cells)

    return table


def _preceding_paragraphs(hwp, ctrl, n: int = 3) -> list[str]:
    """표 직전 문단 텍스트를 최대 n개 수집 (실패해도 빈 목록 반환)."""
    texts: list[str] = []
    try:
        hwp.SetPosBySet(ctrl.GetAnchorPos(0))
        for _ in range(n):
            hwp.HAction.Run("MoveUp")
            hwp.HAction.Run("MoveParaBegin")
            hwp.HAction.Run("MoveSelParaEnd")
            t = (hwp.GetTextFile("TEXT", "saveblock") or "").strip()
            hwp.HAction.Run("Cancel")
            if t and t not in texts:
                texts.append(t)
        texts.reverse()  # 문서 순서(위→아래)로 정렬
    except Exception:
        pass
    return texts
