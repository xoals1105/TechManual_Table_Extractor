"""표 끝 페이지·꼬리말 선택 단위 검증 (한글 불필요)."""
from __future__ import annotations

import sys
from pathlib import Path

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.footer_utils import (  # noqa: E402
    estimate_page_number,
    estimate_table_end_page,
    select_footer_text,
)

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"


def main() -> None:
    xml = f"""<?xml version="1.0"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="{HP}">
  <hp:p pageBreak="0"><hp:run><hp:t>a</hp:t></hp:run>
    <hp:linesegarray><hp:lineseg vertpos="0" vertsize="1000"/></hp:linesegarray></hp:p>
  <hp:p pageBreak="1"><hp:run><hp:tbl rowCnt="2" colCnt="1">
    <hp:tr><hp:tc><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="100" height="30000"/><hp:subList/></hp:tc></hp:tr>
    <hp:tr><hp:tc><hp:cellAddr colAddr="0" rowAddr="1"/><hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="100" height="30000"/><hp:subList/></hp:tc></hp:tr>
  </hp:tbl></hp:run>
  <hp:linesegarray><hp:lineseg vertpos="0" vertsize="60000"/></hp:linesegarray></hp:p>
</hs:sec>"""
    root = etree.fromstring(xml.encode())
    paras = [c for c in root if c.tag == f"{{{HP}}}p"]
    host = paras[1]
    tbl = host.find(f".//{{{HP}}}tbl")
    content_h = 40000
    start = estimate_page_number(paras, host, content_h)
    end = estimate_table_end_page(paras, host, tbl, content_h)
    assert start == 2, start
    assert end >= start + 1, (start, end)
    assert select_footer_text([("ODD", "o"), ("EVEN", "e")], 3) == "o"
    assert select_footer_text([("ODD", "o"), ("EVEN", "e")], 4) == "e"
    assert select_footer_text([("BOTH", "b")], 2) == "b"
    print("footer end-page estimate OK:", f"start={start}, end={end}")


if __name__ == "__main__":
    main()
