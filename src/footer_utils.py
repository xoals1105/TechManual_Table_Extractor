"""꼬리말 텍스트 추출·페이지별 선택 유틸."""
from __future__ import annotations

from lxml import etree

HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def hp_tag(name: str) -> str:
    return f"{{{HP_NS}}}{name}"


def run_text(run_el) -> str:
    """런 내부 텍스트·줄바꿈을 문서 순서대로 이어 붙인다."""
    parts: list[str] = []
    for child in run_el:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag == "t" and child.text:
            parts.append(child.text)
        elif tag == "lineBreak":
            parts.append("\n")
        elif tag == "tab":
            parts.append("\t")
    return "".join(parts)


def paragraph_text(p_el) -> str:
    """문단의 런 텍스트를 순서대로 연결한다."""
    return "".join(run_text(run) for run in p_el.findall(hp_tag("run")))


def footer_element_text(footer_el) -> str:
    """hp:footer 요소에서 문단별 줄바꿈을 유지해 텍스트를 반환."""
    lines: list[str] = []
    for sub in footer_el.findall(hp_tag("subList")):
        for p in sub.findall(hp_tag("p")):
            text = paragraph_text(p).strip()
            if text:
                lines.append(text)
    return "\n".join(lines)


def parse_section_footers(root: etree._Element) -> list[tuple[str, str]]:
    """섹션 XML에서 (applyPageType, 텍스트) 꼬리말 목록을 반환."""
    footers: list[tuple[str, str]] = []
    for footer_el in root.iter(hp_tag("footer")):
        text = footer_element_text(footer_el).strip()
        if not text:
            continue
        apply_type = (footer_el.get("applyPageType") or "BOTH").upper()
        footers.append((apply_type, text))
    return footers


def section_page_layout(root: etree._Element) -> tuple[int, bool]:
    """본문 영역 높이(HWPUNIT)와 첫 쪽 꼬리말 감춤 여부를 반환."""
    sec_pr = root.find(f".//{hp_tag('secPr')}")
    if sec_pr is None:
        return 0, False

    page_pr = sec_pr.find(hp_tag("pagePr"))
    margin = page_pr.find(hp_tag("margin")) if page_pr is not None else None
    page_height = int(page_pr.get("height", 84186)) if page_pr is not None else 84186
    top = int(margin.get("top", 5668)) if margin is not None else 5668
    bottom = int(margin.get("bottom", 4252)) if margin is not None else 4252
    header = int(margin.get("header", 4252)) if margin is not None else 4252
    footer = int(margin.get("footer", 4252)) if margin is not None else 4252
    content_height = page_height - top - bottom - header - footer

    visibility = sec_pr.find(hp_tag("visibility"))
    hide_first = False
    if visibility is not None:
        hide_first = visibility.get("hideFirstFooter", "0") == "1"
    return max(content_height, 0), hide_first


def paragraph_bottom_vert(p_el: etree._Element) -> int:
    max_bottom = 0
    for seg in p_el.iter(hp_tag("lineseg")):
        vert = int(seg.get("vertpos", 0))
        size = int(seg.get("vertsize", 0))
        max_bottom = max(max_bottom, vert + size)
    return max_bottom


def estimate_page_number(
    top_paragraphs: list[etree._Element],
    host_p: etree._Element | None,
    content_height: int,
) -> int:
    """표가 속한 문단까지의 레이아웃 힌트로 쪽 번호를 추정한다."""
    if host_p is None:
        return 1
    if content_height <= 0:
        return 1

    page = 1
    for p in top_paragraphs:
        if p.get("pageBreak", "0") == "1" and p is not host_p:
            page += 1
        if p is host_p:
            break

    bottom = paragraph_bottom_vert(host_p)
    while bottom > content_height:
        page += 1
        bottom -= content_height
    return max(page, 1)


def select_footer_text(
    footers: list[tuple[str, str]],
    page_no: int,
    *,
    hide_first_footer: bool = False,
) -> str:
    """쪽 번호와 applyPageType에 맞는 꼬리말 텍스트를 선택한다."""
    if not footers:
        return ""
    if hide_first_footer and page_no == 1:
        return ""

    odd = [text for kind, text in footers if kind == "ODD"]
    even = [text for kind, text in footers if kind == "EVEN"]
    both = [text for kind, text in footers if kind == "BOTH"]

    if page_no % 2 == 1 and odd:
        return odd[-1]
    if page_no % 2 == 0 and even:
        return even[-1]
    if both:
        return both[-1]
    return ""
