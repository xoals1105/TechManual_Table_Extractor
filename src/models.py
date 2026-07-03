"""표 데이터 모델 정의."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Cell:
    row: int            # 0-based 행 주소
    col: int            # 0-based 열 주소
    row_span: int
    col_span: int
    text: str
    width: int          # HWPUNIT (1/7200 inch)
    height: int         # HWPUNIT


@dataclass
class Table:
    index: int                      # 문서 내 표 순번 (0-based)
    section: int                    # 섹션 번호
    caption: str                    # <hp:caption> 텍스트
    preceding_texts: list[str]      # 표 직전 문단 텍스트 (최대 3개)
    cells: list[Cell] = field(default_factory=list)
    n_rows: int = 0
    n_cols: int = 0

    def header_row(self) -> list[str]:
        """첫 행(row=0) 셀 텍스트를 열 순서대로 반환."""
        return [c.text.strip() for c in sorted(self.cells, key=lambda c: c.col) if c.row == 0]

    def title(self) -> str:
        """엑셀에 기록할 표 제목: 캡션 우선, 없으면 직전 문단 중 마지막 비어있지 않은 것."""
        if self.caption.strip():
            return self.caption.strip()
        for text in reversed(self.preceding_texts):
            if text.strip():
                return text.strip()
        return ""


@dataclass
class MatchResult:
    table: Table
    rule_name: str
    score: float
    reasons: list[str]
