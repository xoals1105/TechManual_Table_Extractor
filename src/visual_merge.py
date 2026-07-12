"""투명/없음 가로 테두리로 분리된 셀을 시각적 병합으로 합친다."""
from __future__ import annotations

from .border_info import is_invisible_border
from .models import Cell, Table


def apply_visual_merges(table: Table) -> Table:
    """인접 행 사이 가로선이 없는 셀을 열 단위로 세로 병합한다."""
    if not any(c.borders for c in table.cells):
        return table

    cells_to_remove: set[int] = set()
    cell_updates: dict[int, Cell] = {}

    for col in range(table.n_cols):
        row = 0
        while row < table.n_rows:
            top_idx = _cell_index_at(table, row, col, cells_to_remove, cell_updates)
            if top_idx is None:
                row += 1
                continue

            top = cell_updates.get(top_idx, table.cells[top_idx])
            group = [top_idx]
            r = top.row + top.row_span

            while r < table.n_rows:
                bottom_idx = _cell_index_at(table, r, col, cells_to_remove, cell_updates)
                if bottom_idx is None:
                    break
                prev = cell_updates.get(group[-1], table.cells[group[-1]])
                bottom = cell_updates.get(bottom_idx, table.cells[bottom_idx])
                if not _can_merge_vertical(prev, bottom):
                    break
                group.append(bottom_idx)
                r = bottom.row + bottom.row_span

            if len(group) > 1:
                merged = _merge_group([cell_updates.get(i, table.cells[i]) for i in group])
                cell_updates[group[0]] = merged
                for idx in group[1:]:
                    cells_to_remove.add(idx)
                row = merged.row + merged.row_span
            else:
                row = top.row + top.row_span

    new_cells = [
        cell_updates[i] if i in cell_updates else c
        for i, c in enumerate(table.cells)
        if i not in cells_to_remove
    ]
    return Table(
        index=table.index,
        section=table.section,
        caption=table.caption,
        preceding_texts=list(table.preceding_texts),
        footer_text=table.footer_text,
        cells=new_cells,
        n_rows=table.n_rows,
        n_cols=table.n_cols,
    )


def _cell_index_at(
    table: Table,
    row: int,
    col: int,
    removed: set[int],
    updates: dict[int, Cell],
) -> int | None:
    for i, raw in enumerate(table.cells):
        if i in removed:
            continue
        c = updates.get(i, raw)
        if c.row <= row < c.row + c.row_span and c.col <= col < c.col + c.col_span:
            return i
    return None


def _can_merge_vertical(upper: Cell, lower: Cell) -> bool:
    if upper.col != lower.col or upper.col_span != lower.col_span:
        return False
    if upper.row + upper.row_span != lower.row:
        return False
    if upper.borders is None or lower.borders is None:
        return False
    return (
        is_invisible_border(upper.borders.bottom)
        and is_invisible_border(lower.borders.top)
    )


def _merge_group(group: list[Cell]) -> Cell:
    anchor = group[0]
    texts = [c.text.strip() for c in group if c.text and c.text.strip()]
    total_height = sum(c.height for c in group if c.height > 0)
    return Cell(
        row=anchor.row,
        col=anchor.col,
        row_span=sum(c.row_span for c in group),
        col_span=anchor.col_span,
        text="\n".join(texts),
        width=anchor.width,
        height=total_height or anchor.height,
        borders=anchor.borders,
    )
