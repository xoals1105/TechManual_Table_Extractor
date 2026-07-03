"""HWPX(ZIP + OWPML XML) 파일에서 표를 추출하는 파서.

HWPX 구조:
  - .hwpx 파일은 ZIP 압축 파일이며, 본문은 Contents/section0.xml, section1.xml ... 에 있음
  - 표는 <hp:tbl> 요소, 행은 <hp:tr>, 셀은 <hp:tc>
  - 셀 주소는 <hp:cellAddr colAddr rowAddr>, 병합은 <hp:cellSpan colSpan rowSpan>,
    크기는 <hp:cellSz width height> (단위: HWPUNIT = 1/7200 inch)
"""
from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from .models import Cell, Table

HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def _tag(name: str) -> str:
    return f"{{{HP_NS}}}{name}"


def _text_of(elem, *, skip_nested_table: bool = False) -> str:
    """요소 내부의 모든 <hp:t> 텍스트를 이어붙여 반환.

    skip_nested_table=True 이면 내부에 중첩된 <hp:tbl> 안의 텍스트는 제외한다
    (셀 텍스트 추출 시 표 안의 표 내용이 섞이는 것을 방지).
    """
    parts: list[str] = []
    for t in elem.iter(_tag("t")):
        if skip_nested_table:
            anc = t.getparent()
            inside_nested = False
            while anc is not None and anc is not elem:
                if anc.tag == _tag("tbl"):
                    inside_nested = True
                    break
                anc = anc.getparent()
            if inside_nested:
                continue
        if t.text:
            parts.append(t.text)
    # 문단 단위 줄바꿈은 유지하지 않고 공백으로 연결 (엑셀 셀에는 한 줄로)
    return " ".join(p.strip() for p in parts if p.strip()).strip()


def _cell_text(tc) -> str:
    """셀 내부 문단들을 줄바꿈으로 연결한 텍스트."""
    lines: list[str] = []
    for p in tc.iter(_tag("p")):
        # 중첩 표 내부의 문단은 제외
        anc = p.getparent()
        inside_nested = False
        while anc is not None and anc is not tc:
            if anc.tag == _tag("tbl"):
                inside_nested = True
                break
            anc = anc.getparent()
        if inside_nested:
            continue
        line = _text_of(p, skip_nested_table=True)
        if line:
            lines.append(line)
    return "\n".join(lines)


def _ancestor_top_paragraph(elem, top_paragraphs: set) -> "etree._Element | None":
    """표 요소가 속한 섹션 최상위 문단(<hp:p>)을 찾는다."""
    anc = elem.getparent()
    while anc is not None:
        if anc in top_paragraphs:
            return anc
        anc = anc.getparent()
    return None


def parse_tables(hwpx_path: str | Path) -> list[Table]:
    """HWPX 파일의 모든 표를 문서 순서대로 파싱하여 반환."""
    hwpx_path = Path(hwpx_path)
    tables: list[Table] = []
    table_index = 0

    with zipfile.ZipFile(hwpx_path) as z:
        section_names = sorted(
            n for n in z.namelist()
            if n.startswith("Contents/section") and n.endswith(".xml")
        )
        if not section_names:
            raise ValueError(f"본문 섹션을 찾을 수 없습니다 (hwpx 파일이 맞는지 확인): {hwpx_path}")

        for sec_no, sec_name in enumerate(section_names):
            root = etree.fromstring(z.read(sec_name))
            # 섹션 최상위 문단 목록 (표 직전 문단 탐색용)
            top_paragraphs = [child for child in root if child.tag == _tag("p")]
            top_para_set = set(top_paragraphs)

            for tbl in root.iter(_tag("tbl")):
                table = Table(
                    index=table_index,
                    section=sec_no,
                    caption="",
                    preceding_texts=[],
                    n_rows=int(tbl.get("rowCnt", 0)),
                    n_cols=int(tbl.get("colCnt", 0)),
                )
                table_index += 1

                cap = tbl.find(_tag("caption"))
                if cap is not None:
                    table.caption = _text_of(cap)

                host_p = _ancestor_top_paragraph(tbl, top_para_set)
                if host_p is not None:
                    i = top_paragraphs.index(host_p)
                    preceding = []
                    for p in top_paragraphs[max(0, i - 3):i]:
                        # 다른 표를 담은 문단은 제외
                        if p.find(f".//{_tag('tbl')}") is not None:
                            continue
                        preceding.append(_text_of(p))
                    table.preceding_texts = preceding

                for tr in tbl.findall(_tag("tr")):
                    for tc in tr.findall(_tag("tc")):
                        addr = tc.find(_tag("cellAddr"))
                        span = tc.find(_tag("cellSpan"))
                        size = tc.find(_tag("cellSz"))
                        if addr is None:
                            continue
                        table.cells.append(Cell(
                            row=int(addr.get("rowAddr", 0)),
                            col=int(addr.get("colAddr", 0)),
                            row_span=int(span.get("rowSpan", 1)) if span is not None else 1,
                            col_span=int(span.get("colSpan", 1)) if span is not None else 1,
                            text=_cell_text(tc),
                            width=int(size.get("width", 0)) if size is not None else 0,
                            height=int(size.get("height", 0)) if size is not None else 0,
                        ))

                tables.append(table)

    return tables
