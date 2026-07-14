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

from .border_info import parse_hml_border_fills, resolve_hml_cell_borders
from .footer_utils import select_footer_text
from .models import Cell, Table

_MAX_FOOTER_PARAS = 50

_XML_DECL = re.compile(r"^\s*<\?xml[^>]*\?>")
_FOOTER_BLOCK = re.compile(r"<FOOTER\b[\s\S]*?</FOOTER>", re.IGNORECASE)
_SECTION_SPLIT = re.compile(r"(?=<SECTION[\s>])", re.IGNORECASE)
_CHAR_TEXT = re.compile(r"<CHAR[^>]*>([^<]*)</CHAR>", re.IGNORECASE)
_P_BLOCK = re.compile(r"<P\b[\s\S]*?</P>", re.IGNORECASE)
_APPLY_PAGE = re.compile(r'ApplyPageType="([^"]*)"', re.IGNORECASE)


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
        tbl_ctrls: list = []
        idx = 0
        ctrl = hwp.HeadCtrl
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                table = _read_one_table(hwp, ctrl, idx)
                if table is not None:
                    tables.append(table)
                    tbl_ctrls.append(ctrl)
                    idx += 1
            ctrl = ctrl.Next

        # 꼬리말: Goto(찾아가기) 없이 문서 전체 HWPML2X의 <FOOTER>를 파싱한다.
        try:
            full_xml = hwp.GetTextFile("HWPML2X", "") or ""
            _assign_footers_from_hml(hwp, tables, tbl_ctrls, full_xml)
        except Exception:
            pass
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


def _footer_element_text_hml(footer_xml: str) -> str:
    """HWPML <FOOTER> 블록에서 문단별 텍스트를 줄바꿈으로 이어 붙인다."""
    lines: list[str] = []
    for p in _P_BLOCK.findall(footer_xml):
        text = "".join(_CHAR_TEXT.findall(p)).strip()
        if text:
            lines.append(text)
    if not lines:
        text = "".join(_CHAR_TEXT.findall(footer_xml)).strip()
        if text:
            lines.append(text)
    return "\n".join(lines)


def _parse_hml_section_footers(xml: str) -> list[list[tuple[str, str]]]:
    """전체 HWPML2X에서 섹션별 (applyPageType, text) 꼬리말 목록을 추출."""
    parts = _SECTION_SPLIT.split(xml)
    result: list[list[tuple[str, str]]] = []
    for part in parts:
        if not part.lstrip().upper().startswith("<SECTION"):
            continue
        foots: list[tuple[str, str]] = []
        for block in _FOOTER_BLOCK.findall(part):
            apply = "BOTH"
            m = _APPLY_PAGE.search(block)
            if m:
                apply = m.group(1).upper()
            text = _footer_element_text_hml(block).strip()
            if text:
                foots.append((apply, text))
        result.append(foots)
    return result


def _table_section_indices(hwp) -> list[int]:
    """HeadCtrl 순서 기준으로 각 표의 구역(0-based) 인덱스를 반환."""
    sec_i = -1
    indices: list[int] = []
    ctrl = hwp.HeadCtrl
    while ctrl is not None:
        if ctrl.CtrlID == "secd":
            sec_i += 1
        elif ctrl.CtrlID == "tbl":
            indices.append(max(sec_i, 0))
        ctrl = ctrl.Next
    return indices


def _assign_footers_from_hml(hwp, tables: list[Table], tbl_ctrls: list, full_xml: str) -> None:
    """Goto 없이 HWPML <FOOTER> + 구역 상속으로 표 footer_text를 채운다."""
    if not tables or not full_xml:
        return

    section_foots = _parse_hml_section_footers(full_xml)
    if not section_foots and _FOOTER_BLOCK.search(full_xml):
        # SECTION 래퍼가 약한 경우: 문서 순서 FOOTER만 모은다
        foots: list[tuple[str, str]] = []
        for block in _FOOTER_BLOCK.findall(full_xml):
            apply = "BOTH"
            m = _APPLY_PAGE.search(block)
            if m:
                apply = m.group(1).upper()
            text = _footer_element_text_hml(block).strip()
            if text:
                foots.append((apply, text))
        section_foots = [foots] if foots else []

    # 이전 구역과 동일(꼬리말 없음) → 상속
    effective: list[list[tuple[str, str]]] = []
    inherited: list[tuple[str, str]] = []
    for foots in section_foots:
        if foots:
            inherited = list(foots)
            effective.append(list(foots))
        else:
            effective.append(list(inherited))

    table_secs = _table_section_indices(hwp)
    for i, table in enumerate(tables):
        sec = table_secs[i] if i < len(table_secs) else 0
        if sec >= len(effective):
            sec = len(effective) - 1 if effective else 0
        candidates = effective[sec] if effective else []
        page_no = 1
        if i < len(tbl_ctrls):
            try:
                page_no = _page_at_table_end(hwp, tbl_ctrls[i])
            except Exception:
                page_no = 1
        table.section = sec
        table.footer_text = select_footer_text(candidates, page_no)


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

    border_fills = parse_hml_border_fills(root)

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
            borders=resolve_hml_cell_borders(cell_el, border_fills),
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


def _goto_table_last_cell(hwp, ctrl) -> None:
    """표 컨트롤의 마지막 셀(끝 페이지 쪽)으로 캐럿을 이동한다."""
    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
    hwp.FindCtrl()
    hwp.HAction.Run("ShapeObjTableSelCell")
    for _ in range(500):
        before = hwp.GetPos()
        try:
            hwp.HAction.Run("TableColEnd")
        except Exception:
            pass
        try:
            hwp.HAction.Run("TableLowerCell")
        except Exception:
            pass
        if hwp.GetPos() == before:
            break


def _exit_table_to_body(hwp) -> None:
    """표의 마지막 셀에서 같은(끝) 쪽 본문으로 빠져나온다.

    CloseEx는 표 앵커(시작 쪽)로 돌아가므로 사용하지 않는다.
    마지막 셀에서 MoveDown 하면 표 바로 다음 본문(보통 끝 쪽)으로 이동한다.
    """
    try:
        hwp.HAction.Run("Cancel")
    except Exception:
        pass
    for _ in range(8):
        if hwp.GetPos()[0] == 0:
            return
        before = hwp.GetPos()
        try:
            hwp.HAction.Run("MoveDown")
        except Exception:
            break
        if hwp.GetPos() == before:
            break
    # 그래도 표 안이면 오른쪽으로 한 번 더
    for _ in range(3):
        if hwp.GetPos()[0] == 0:
            return
        try:
            hwp.HAction.Run("MoveRight")
        except Exception:
            break



def _move_to_current_list_begin(hwp) -> None:
    """현재 리스트 인스턴스의 첫 문단으로 이동한다.

    SetPos(list_id, 0, 0)은 문서 전체에서 같은 목록 ID의 첫 인스턴스
    (대개 첫 페이지/첫 구역 꼬리말)로 점프하므로 사용하지 않는다.
    """
    list_id = hwp.GetPos()[0]
    if list_id == 0:
        return
    for _ in range(_MAX_FOOTER_PARAS):
        before = hwp.GetPos()
        try:
            hwp.HAction.Run("MovePrevPara")
        except Exception:
            break
        after = hwp.GetPos()
        if after[0] != list_id:
            # 리스트를 벗어났으면 한 문단 복귀
            try:
                hwp.HAction.Run("MoveNextPara")
            except Exception:
                pass
            break
        if after == before:
            break


def _read_list_paragraphs(hwp, list_id: int | None = None) -> str:
    """현재 캐럿이 있는 리스트(꼬리말 등) 문단을 순서·줄바꿈 유지해 반환.

    list_id가 주어지면 현재 리스트와 같을 때만 읽는다.
    SetPos(list_id,0,0)은 다른 쪽/구역의 첫 꼬리말로 점프할 수 있어 쓰지 않는다.
    """
    cur_id = hwp.GetPos()[0]
    if cur_id == 0:
        return ""
    if list_id is not None and cur_id != list_id:
        return ""

    _move_to_current_list_begin(hwp)
    lines: list[str] = []
    for _ in range(_MAX_FOOTER_PARAS):
        if hwp.GetPos()[0] != cur_id:
            break
        hwp.HAction.Run("MoveParaBegin")
        hwp.HAction.Run("MoveSelParaEnd")
        text = (hwp.GetTextFile("TEXT", "saveblock") or "").strip()
        hwp.HAction.Run("Cancel")
        if text:
            lines.append(text.replace("\r\n", "\n").replace("\r", "\n"))
        pos = hwp.GetPos()
        hwp.HAction.Run("MoveNextPara")
        if hwp.GetPos() == pos:
            break
        if hwp.GetPos()[0] != cur_id:
            break
    return "\n".join(lines)


def _enter_footer(hwp) -> bool:
    """꼬리말 편집 영역으로 들어간다. 성공 여부 반환.

    Goto(14)는 '다음 꼬리말 컨트롤'을 찾는다. 캐럿을 본문 앞으로 둔 뒤
    호출하면 문서 순서의 꼬리말을 순회할 수 있다.
    (현재 쪽 꼬리말만 보장하지는 않으므로, 표 매칭은 컨트롤 순서를 쓴다.)
    """
    act, ps = hwp.HAction, hwp.HParameterSet
    act.GetDefault("Goto", ps.HGotoE.HSet)
    ps.HGotoE.HSet.SetItem("DialogResult", 14)  # 꼬리말
    ps.HGotoE.SetSelectionIndex = 5
    act.Execute("Goto", ps.HGotoE.HSet)
    try:
        hwp.HAction.Run("HeaderFooterModify")
    except Exception:
        pass
    if hwp.GetPos()[0] != 0:
        return True
    try:
        hwp.HAction.Run("Footer")
        try:
            hwp.HAction.Run("HeaderFooterModify")
        except Exception:
            pass
    except Exception:
        return False
    return hwp.GetPos()[0] != 0


def _footer_kind_from_desc(user_desc: str) -> str:
    """꼬리말 컨트롤 UserDesc → BOTH / ODD / EVEN."""
    desc = user_desc or ""
    if "홀수" in desc:
        return "ODD"
    if "짝수" in desc:
        return "EVEN"
    return "BOTH"


def _collect_footer_texts_in_order(hwp) -> list[str]:
    """문서에 등장하는 꼬리말 컨트롤 텍스트를 순서대로 수집 (한 바퀴면 종료)."""
    texts: list[str] = []
    first_list_id: int | None = None
    try:
        hwp.HAction.Run("MoveDocBegin")
    except Exception:
        pass

    for _ in range(40):
        if not _enter_footer(hwp):
            break
        list_id = hwp.GetPos()[0]
        text = _read_list_paragraphs(hwp, list_id)
        try:
            hwp.HAction.Run("CloseEx")
        except Exception:
            pass

        if first_list_id is None:
            first_list_id = list_id
            texts.append(text)
        elif list_id == first_list_id:
            break
        else:
            texts.append(text)

        for _ in range(3):
            try:
                hwp.HAction.Run("MoveRight")
            except Exception:
                break
    return texts


def _page_at_table_end(hwp, ctrl) -> int:
    """표 마지막 셀(끝 쪽) 기준 쪽 번호. KeyIndicator 실패 시 1."""
    try:
        _goto_table_last_cell(hwp, ctrl)
        _exit_table_to_body(hwp)
        if hwp.GetPos()[0] != 0:
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
        ki = hwp.KeyIndicator()
        # (ok, ?, page, ...) — page는 보통 인덱스 2
        if isinstance(ki, (tuple, list)) and len(ki) > 2:
            page = int(ki[2])
            if page >= 1:
                return page
    except Exception:
        pass
    return 1


def _select_footer_from_candidates(
    candidates: list[tuple[str, str]],
    page_no: int,
) -> str:
    """구역 내 (kind, text) 후보에서 쪽 번호에 맞는 꼬리말을 고른다."""
    if not candidates:
        return ""
    odd = [t for k, t in candidates if k == "ODD" and t]
    even = [t for k, t in candidates if k == "EVEN" and t]
    both = [t for k, t in candidates if k == "BOTH" and t]
    if page_no % 2 == 1 and odd:
        return odd[-1]
    if page_no % 2 == 0 and even:
        return even[-1]
    if both:
        return both[-1]
    if odd:
        return odd[-1]
    if even:
        return even[-1]
    return candidates[-1][1]


def _footer_map_for_tables(hwp, tbl_ctrls: list) -> list[str]:
    """각 표에 대응하는 꼬리말 텍스트 목록 (표 컨트롤 순서와 동일).

    구역(secd) 단위로 꼬리말을 모은 뒤, 그 구역의 모든 표에 적용한다.
    foot 컨트롤이 표보다 뒤에 있어도(한글 문서에서 흔함) 동일 구역이면 적용한다.
    꼬리말 없는 구역은 직전 구역 꼬리말을 상속한다(이전 구역과 동일).
    """
    foot_texts = _collect_footer_texts_in_order(hwp)
    if not foot_texts:
        return [""] * len(tbl_ctrls)

    # 1패스: 구역별 꼬리말·표 인덱스 수집
    sections: list[dict] = []
    current: dict = {"foots": [], "tbl_indices": []}
    foot_i = 0
    tbl_i = 0
    ctrl = hwp.HeadCtrl
    first_secd = True
    while ctrl is not None:
        cid = ctrl.CtrlID
        if cid == "secd":
            if not first_secd:
                sections.append(current)
                current = {"foots": [], "tbl_indices": []}
            first_secd = False
        elif cid == "foot":
            kind = _footer_kind_from_desc(getattr(ctrl, "UserDesc", "") or "")
            text = foot_texts[foot_i] if foot_i < len(foot_texts) else ""
            foot_i += 1
            current["foots"].append((kind, text))
        elif cid == "tbl":
            current["tbl_indices"].append(tbl_i)
            tbl_i += 1
        ctrl = ctrl.Next
    sections.append(current)

    result = [""] * len(tbl_ctrls)
    inherited: list[tuple[str, str]] = []
    for sec in sections:
        foots = sec["foots"] or list(inherited)
        if sec["foots"]:
            inherited = list(sec["foots"])
        for ti in sec["tbl_indices"]:
            if ti >= len(tbl_ctrls):
                continue
            page_no = _page_at_table_end(hwp, tbl_ctrls[ti])
            result[ti] = _select_footer_from_candidates(foots, page_no)
    return result


def _footer_at_table(hwp, ctrl) -> str:
    """표가 끝나는 쪽의 꼬리말 텍스트를 반환."""
    try:
        return _footer_map_for_tables(hwp, [ctrl])[0]
    except Exception:
        try:
            hwp.HAction.Run("CloseEx")
        except Exception:
            pass
        return ""


def enrich_footers_via_com(path: str | Path, tables: list[Table], visible: bool = False) -> int:
    """한글 COM으로 각 표의 끝 페이지 꼬리말을 읽어 footer_text를 덮어쓴다.

    표 컨트롤 문서 순서가 tables 순서와 같다고 가정한다.
    성공한 표 개수를 반환한다. 한글이 없거나 열기 실패 시 0.
    """
    if not tables:
        return 0
    try:
        import win32com.client as win32
    except ImportError:
        return 0

    path = Path(path).resolve()
    hwp = None
    updated = 0
    try:
        try:
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        except Exception:
            hwp = win32.Dispatch("HWPFrame.HwpObject")
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        try:
            hwp.XHwpWindows.Item(0).Visible = visible
        except Exception:
            pass

        fmt = "HWPX" if path.suffix.lower() == ".hwpx" else "HWP"
        if not hwp.Open(str(path), fmt, "forceopen:true"):
            return 0

        tbl_ctrls = []
        ctrl = hwp.HeadCtrl
        while ctrl is not None:
            if ctrl.CtrlID == "tbl":
                tbl_ctrls.append(ctrl)
            ctrl = ctrl.Next

        # Goto 없이 HWPML FOOTER 파싱 (찾아가기 팝업 방지)
        full_xml = hwp.GetTextFile("HWPML2X", "") or ""
        _assign_footers_from_hml(hwp, tables, tbl_ctrls, full_xml)
        for table in tables:
            if table.footer_text:
                updated += 1
        return updated
    except Exception:
        return 0
    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)
            except Exception:
                pass
            try:
                hwp.Quit()
            except Exception:
                pass
