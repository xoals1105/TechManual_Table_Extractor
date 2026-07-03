"""hwp_com_reader의 HML(XML) 파서 단위 검증 (한글 설치 불필요).

한글 COM의 GetTextFile("HWPML2X", "saveblock") 결과와 동일한 구조의
샘플 XML로 파싱 → 선별 → 엑셀 저장까지 검증한다.

사용법: .venv\\Scripts\\python tools\\test_hml_parser.py
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.hwp_com_reader import _parse_hml_table  # noqa: E402
from src.table_matcher import select_tables  # noqa: E402
from src.excel_writer import write_workbook  # noqa: E402
from src.models import MatchResult  # noqa: E402

SAMPLE_HML = """<?xml version="1.0" encoding="UTF-16"?>
<HWPML Version="2.91" SubVersion="10.0.0.0" Style="export">
 <BODY>
  <SECTION Id="0">
   <P ParaShape="0" Style="0">
    <TEXT CharShape="0">
     <TABLE BorderFill="3" CellSpacing="0" ColCount="4" RowCount="4" PageBreak="Cell" RepeatHeader="true">
      <SHAPEOBJECT InstId="1" ZOrder="0" NumberingType="Table">
       <SIZE Width="42000" Height="8000" WidthRelTo="Absolute" HeightRelTo="Absolute" Protect="false"/>
       <CAPTION Side="Top" FullSize="false" Width="8504" Gap="850" LastWidth="42000">
        <PARALIST LineWrap="Break" VertAlign="Top" LinkListID="0" LinkListIDNext="0" TextDirection="0">
         <P ParaShape="1" Style="0"><TEXT CharShape="1"><CHAR>표 3-1 부품 목록</CHAR></TEXT></P>
        </PARALIST>
       </CAPTION>
      </SHAPEOBJECT>
      <ROW>
       <CELL Name="A1" ColAddr="0" RowAddr="0" ColSpan="1" RowSpan="1" Width="10500" Height="850" Header="false" HasMargin="false" Protect="false" Editable="false" Dirty="false" BorderFill="4">
        <PARALIST><P ParaShape="2" Style="0"><TEXT CharShape="2"><CHAR>품명</CHAR></TEXT></P></PARALIST>
       </CELL>
       <CELL Name="B1" ColAddr="1" RowAddr="0" ColSpan="1" RowSpan="1" Width="10500" Height="850" BorderFill="4">
        <PARALIST><P><TEXT><CHAR>규격</CHAR></TEXT></P></PARALIST>
       </CELL>
       <CELL Name="C1" ColAddr="2" RowAddr="0" ColSpan="1" RowSpan="1" Width="10500" Height="850" BorderFill="4">
        <PARALIST><P><TEXT><CHAR>수량</CHAR></TEXT></P></PARALIST>
       </CELL>
       <CELL Name="D1" ColAddr="3" RowAddr="0" ColSpan="1" RowSpan="1" Width="10500" Height="850" BorderFill="4">
        <PARALIST><P><TEXT><CHAR>비고</CHAR></TEXT></P></PARALIST>
       </CELL>
      </ROW>
      <ROW>
       <CELL ColAddr="0" RowAddr="1" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>볼트</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="1" RowAddr="1" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>M8x20</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="2" RowAddr="1" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>4</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="3" RowAddr="1" ColSpan="1" RowSpan="2" Width="10500" Height="1700"><PARALIST><P><TEXT><CHAR>공용</CHAR></TEXT></P></PARALIST></CELL>
      </ROW>
      <ROW>
       <CELL ColAddr="0" RowAddr="2" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>너트</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="1" RowAddr="2" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>M8</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="2" RowAddr="2" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>4</CHAR></TEXT></P></PARALIST></CELL>
      </ROW>
      <ROW>
       <CELL ColAddr="0" RowAddr="3" ColSpan="2" RowSpan="1" Width="21000" Height="850"><PARALIST><P><TEXT><CHAR>합계</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="2" RowAddr="3" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>8</CHAR></TEXT></P></PARALIST></CELL>
       <CELL ColAddr="3" RowAddr="3" ColSpan="1" RowSpan="1" Width="10500" Height="850"><PARALIST><P><TEXT><CHAR>-</CHAR></TEXT></P></PARALIST></CELL>
      </ROW>
     </TABLE>
    </TEXT>
   </P>
  </SECTION>
 </BODY>
</HWPML>
"""


def main() -> None:
    table = _parse_hml_table(SAMPLE_HML, 0)
    assert table is not None, "파싱 결과가 None"
    assert table.n_rows == 4 and table.n_cols == 4, f"행/열 오류: {table.n_rows}x{table.n_cols}"
    assert table.caption == "표 3-1 부품 목록", f"캡션 오류: {table.caption!r}"
    assert table.header_row() == ["품명", "규격", "수량", "비고"], table.header_row()
    merged = [(c.row, c.col, c.row_span, c.col_span) for c in table.cells
              if c.row_span > 1 or c.col_span > 1]
    assert (1, 3, 2, 1) in merged and (3, 0, 1, 2) in merged, f"병합 오류: {merged}"
    print("HML 파싱 OK:", f"{table.n_rows}x{table.n_cols},",
          f"캡션='{table.caption}',", f"셀 {len(table.cells)}개, 병합 {merged}")

    rule = {"name": "부품목록표",
            "expected_headers": ["품명", "규격", "수량", "비고"],
            "title_keywords": ["부품 목록"], "expected_cols": 4, "threshold": 60}
    results = select_tables([table], [rule])
    assert len(results) == 1, "선별 실패"

    out = Path("output/test_hml_parser.xlsx")
    write_workbook(results, out)
    print("엑셀 저장 OK:", out)


if __name__ == "__main__":
    main()
