"""XML로 2구역 꼬리말 테스트 HWPX를 만든다."""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from src.hwpx_parser import parse_tables  # noqa: E402

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
SRC = ROOT / "data" / "test_table_footer3.hwpx"
DST = ROOT / "data" / "test_2sec_footer.hwpx"


def main() -> None:
    with zipfile.ZipFile(SRC) as z:
        sec0 = z.read("Contents/section0.xml")
        hpf = z.read("Contents/content.hpf").decode("utf-8")
        others = {
            n: z.read(n)
            for n in z.namelist()
            if n not in ("Contents/section0.xml", "Contents/content.hpf")
        }

    root = etree.fromstring(sec0)
    sec1_root = etree.fromstring(sec0)

    for t in root.iter(f"{HP}t"):
        if t.text and "꼬리말" in t.text:
            t.text = "FOOTER_SEC1"

    for t in sec1_root.iter(f"{HP}t"):
        if t.text and "꼬리말" in t.text:
            t.text = "FOOTER_SEC2"
        elif t.text in ("A", "B", "C", "D"):
            t.text = "T2" + t.text

    if "section1" not in hpf:
        hpf = hpf.replace(
            'id="section0" href="Contents/section0.xml" media-type="application/xml"/>',
            'id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
            '<opf:item id="section1" href="Contents/section1.xml" media-type="application/xml"/>',
        )
        hpf = hpf.replace(
            '<opf:itemref idref="section0"/>',
            '<opf:itemref idref="section0"/><opf:itemref idref="section1"/>',
        )

    with zipfile.ZipFile(DST, "w") as zout:
        for n, data in others.items():
            zout.writestr(n, data)
        zout.writestr(
            "Contents/section0.xml",
            etree.tostring(root, xml_declaration=True, encoding="UTF-8"),
        )
        zout.writestr(
            "Contents/section1.xml",
            etree.tostring(sec1_root, xml_declaration=True, encoding="UTF-8"),
        )
        zout.writestr("Contents/content.hpf", hpf.encode("utf-8"))

    print("wrote", DST)
    tables = parse_tables(DST)
    print("xml tables", len(tables))
    for i, t in enumerate(tables):
        cell0 = t.cells[0].text if t.cells else None
        print(i, "section", t.section, "footer", repr(t.footer_text), "cell0", cell0)


if __name__ == "__main__":
    main()
