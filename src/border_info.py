"""HWP/HWPX 표 셀 테두리(borderFill) 파싱."""
from __future__ import annotations

import zipfile

from lxml import etree

from .models import CellBorders

HH_NS = "http://www.hancom.co.kr/hwpml/2011/head"


def _hh_tag(name: str) -> str:
    return f"{{{HH_NS}}}{name}"


def is_invisible_border(border_type: str | None) -> bool:
    """테두리가 화면에 보이지 않는지 판별 (NONE / 빈 값)."""
    if not border_type:
        return True
    return border_type.strip().upper() in ("NONE", "NIL", "")


def borders_from_sides(
    *,
    left: str | None = "SOLID",
    right: str | None = "SOLID",
    top: str | None = "SOLID",
    bottom: str | None = "SOLID",
) -> CellBorders:
    return CellBorders(
        left=left or "NONE",
        right=right or "NONE",
        top=top or "NONE",
        bottom=bottom or "NONE",
    )


def parse_hwpx_border_fills(header_root: etree._Element) -> dict[str, CellBorders]:
    """header.xml 의 borderFill 정의를 id → CellBorders 로 변환."""
    fills: dict[str, CellBorders] = {}
    for bf in header_root.iter(_hh_tag("borderFill")):
        bf_id = bf.get("id")
        if bf_id is None:
            continue
        fills[bf_id] = borders_from_sides(
            left=_hwpx_side_type(bf, "leftBorder"),
            right=_hwpx_side_type(bf, "rightBorder"),
            top=_hwpx_side_type(bf, "topBorder"),
            bottom=_hwpx_side_type(bf, "bottomBorder"),
        )
    return fills


def _hwpx_side_type(bf_el: etree._Element, side_name: str) -> str:
    el = bf_el.find(_hh_tag(side_name))
    if el is None:
        return "NONE"
    return el.get("type", "NONE")


def load_hwpx_border_fills(z: zipfile.ZipFile) -> dict[str, CellBorders]:
    """HWPX ZIP 에서 header.xml 테두리 정의를 읽는다."""
    names = [n for n in z.namelist() if n.replace("\\", "/").endswith("Contents/header.xml")]
    if not names:
        return {}
    root = etree.fromstring(z.read(names[0]))
    return parse_hwpx_border_fills(root)


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]
    return tag


def parse_hml_border_fills(root: etree._Element) -> dict[str, CellBorders]:
    """HWPML2X XML 의 BORDERFILL 정의를 id → CellBorders 로 변환."""
    fills: dict[str, CellBorders] = {}
    for el in root.iter():
        if _local_name(el.tag).upper() != "BORDERFILL":
            continue
        bf_id = el.get("Id") or el.get("id") or el.get("ID")
        if bf_id is None:
            continue
        sides = {"left": "SOLID", "right": "SOLID", "top": "SOLID", "bottom": "SOLID"}
        for child in el:
            name = _local_name(child.tag).upper()
            side_key = {
                "LEFTBORDER": "left",
                "RIGHTBORDER": "right",
                "TOPBORDER": "top",
                "BOTTOMBORDER": "bottom",
            }.get(name)
            if side_key is None:
                continue
            sides[side_key] = child.get("Type") or child.get("type") or "NONE"
        fills[str(bf_id)] = borders_from_sides(**sides)
    return fills


def resolve_hml_cell_borders(
    cell_el: etree._Element,
    border_fills: dict[str, CellBorders],
) -> CellBorders | None:
    bf_id = cell_el.get("BorderFill") or cell_el.get("borderFill")
    if bf_id is None:
        return None
    return border_fills.get(str(bf_id))
