"""Tables: make them, read them, fill them, shape them.

A table is found by its number in the document (1 = the first), and its
cells by row and column, both numbered from 1. Each cell's text is also a
paragraph of the document, so a result gives the paragraphs a table covers:
word_format_text and the other paragraph tools reach into it with those.

Measured on both platforms (docs/PROBES.md): a new table has the Normal
Table style, which draws no lines, so a table made without a style gets
its borders turned on (Borders.Enable, https://learn.microsoft.com/en-us/office/vba/api/word.borders.enable).
An English name such as "Table Grid" is not taken for a style on a Word in
another language; the built-in table styles that have a WdBuiltinStyle
constant ("Light Grid", "Medium Shading 1 Accent 1"...) are set by it, and
any other by the name word_list_styles gives.
"""

from __future__ import annotations

import re
from contextlib import contextmanager
from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, annotations, call, follow, mac_impl
from ._props import bgr, hex_of
from ._word import ALIGNMENTS, BUILTIN_STYLES, flag, index_of, paragraph, resolve_style

CELL_END = "\r\x07"  # what ends each cell, and each row, in a table's text
WD_WITHIN_TABLE = 12  # WdInformation
WD_COLOR_AUTOMATIC = -16777216
WD_STYLE_TYPE_TABLE = 3
WD_PREFERRED_WIDTH_AUTO = 1
WD_PREFERRED_WIDTH_POINTS = 3
AUTOFIT = {"fixed": 0, "content": 1, "window": 2}  # WdAutoFitBehavior
ROW_ALIGNMENTS = {"left": 0, "center": 1, "right": 2}  # WdRowAlignment
VERTICAL = {"top": 0, "center": 1, "bottom": 3}  # WdCellVerticalAlignment
TABLE_STYLE_CONSTANTS = range(-178, -158)  # wdStyleTableLightShading ... MediumList1Accent1
MAX_ROWS_PER_READ = 100

# -- arguments ---------------------------------------------------------------------

class TableLook(Input):
    """What a table looks like as a whole: shared by inserting and changing one."""

    style: str | None = Field(None, max_length=255, description=(
        "A table style: a built-in English name that has a constant (Light Grid, Light "
        "List, Light Shading, Medium Shading 1, Medium Grid 1, Colorful List, Medium "
        "Shading 1 Accent 1...), or any table style's name as word_list_styles gives "
        "it (\"Grid Table 4 - Accent 1\" and the like go by their local name)."))
    header_rows: int | None = Field(None, ge=0, le=10, description=(
        "How many rows at the top are headings: repeated at the top of each page, "
        "and shown as headings by the style."))
    borders: bool | None = Field(None, description=(
        "Grid lines around every cell. A new table without a style gets them."))
    column_widths: list[float] | None = Field(None, min_length=1, max_length=63, description=(
        "Width of each column in points, left to right (72 points = 1 inch)."))
    autofit: Literal["window", "content", "fixed"] | None = Field(None, description=(
        "window: as wide as the page's text. content: each column as wide as its "
        "text. fixed: the widths stay as set."))
    alignment: Literal["left", "center", "right"] | None = Field(None, description=(
        "Where the table sits between the margins."))


def _text(value: str) -> str:
    return value.replace("\r\n", "\n").replace("\n", "\v")


def _cell_text(raw: str) -> str:
    """A cell's text as a tool gives it: paragraphs and line breaks as \\n."""
    return raw.replace("\r", "\n").replace("\v", "\n")


class InsertTableInput(TableLook):
    rows: list[list[str]] = Field(..., min_length=1, max_length=500, description=(
        "The cells, row by row: [[\"Name\", \"Score\"], [\"Ann\", \"92\"]]. A short row is "
        "filled with empty cells; \\n starts a new line within a cell."))
    after: int | None = Field(None, ge=0, description=(
        "Put the table after this paragraph (0 = at the very start). Neither "
        "after nor before: at the end."))
    before: int | None = Field(None, ge=1, description="Put it before this paragraph.")
    header_rows: int = Field(1, ge=0, le=10, description=(
        "How many rows at the top are headings (default 1): repeated on each page, "
        "shown as headings by the style, and bold when no style is given."))
    column_alignments: list[Literal["left", "center", "right", "justify"] | None] | None = (
        Field(None, max_length=63, description=(
            "Alignment of the text in each column, left to right, e.g. "
            "[null, \"right\"] for numbers on the right in the second.")))

    @model_validator(mode="after")
    def _shape(self):
        if self.after is not None and self.before is not None:
            raise ValueError("Give `after` or `before`, not both.")
        columns = max(len(r) for r in self.rows)
        if not 1 <= columns <= 63:
            raise ValueError("A table has 1 to 63 columns.")
        if self.header_rows > len(self.rows):
            raise ValueError("header_rows is more than the rows given.")
        for name in ("column_widths", "column_alignments"):
            value = getattr(self, name)
            if value is not None and len(value) != columns:
                raise ValueError(f"{name} has {len(value)} entries; the table has "
                                 f"{columns} columns.")
        return self


class TableInput(Input):
    table: int = Field(..., ge=1, description="The table, numbered from 1 in the document.")


class GetTableInput(TableInput):
    start_row: int = Field(1, ge=1, description="First row to read.")


class SetTableInput(TableInput, TableLook):
    pass


class CellBlock(TableInput):
    row: int = Field(1, ge=1, description="Top row of the block.")
    column: int = Field(1, ge=1, description="Left column of the block.")


class SetCellsInput(CellBlock):
    values: list[list[str | None]] = Field(..., min_length=1, max_length=500, description=(
        "The text for a block of cells, row by row, starting at `row`/`column`; "
        "null leaves a cell as it is."))


class FormatCellsInput(TableInput):
    from_row: int = Field(1, ge=1, description="First row (default 1).")
    to_row: int | None = Field(None, ge=1, description="Last row (default: the last).")
    from_column: int = Field(1, ge=1, description="First column (default 1).")
    to_column: int | None = Field(None, ge=1, description="Last column (default: the last).")
    shading: str | None = Field(None, pattern=r"^(#?[0-9A-Fa-f]{6}|none)$", description=(
        "Background color #RRGGBB, or none."))
    bold: bool | None = None
    italic: bool | None = None
    color: str | None = Field(None, pattern=r"^#?[0-9A-Fa-f]{6}$", description=(
        "Text color #RRGGBB."))
    size: float | None = Field(None, ge=1, le=1638, description="Font size in points.")
    alignment: Literal["left", "center", "right", "justify", "distribute"] | None = Field(
        None, description="Horizontal alignment of the text.")
    vertical_alignment: Literal["top", "center", "bottom"] | None = None

    @model_validator(mode="after")
    def _something(self):
        if all(getattr(self, n) is None for n in (
                "shading", "bold", "italic", "color", "size", "alignment",
                "vertical_alignment")):
            raise ValueError("Give at least one of shading, bold, italic, color, size, "
                             "alignment, vertical_alignment.")
        return self


class InsertRowsInput(TableInput):
    after_row: int | None = Field(None, ge=0, description=(
        "Put the rows after this row (0 = above the first). Default: at the bottom."))
    count: int | None = Field(None, ge=1, le=500, description=(
        "How many empty rows. Not needed with `values`."))
    values: list[list[str]] | None = Field(None, min_length=1, max_length=500, description=(
        "The new rows' cells, row by row."))

    @model_validator(mode="after")
    def _how_many(self):
        if self.count is None and self.values is None:
            raise ValueError("Give `values` (the new rows) or `count` (empty rows).")
        if self.count is not None and self.values is not None and self.count != len(self.values):
            raise ValueError("count and the number of rows in values differ; give one.")
        return self


class InsertColumnsInput(TableInput):
    after_column: int | None = Field(None, ge=0, description=(
        "Put the columns after this column (0 = at the left). Default: at the right."))
    count: int | None = Field(None, ge=1, le=62, description="How many empty columns.")
    values: list[list[str]] | None = Field(None, min_length=1, max_length=62, description=(
        "The new columns' cells, a column at a time from the top: [[\"Unit\", \"kg\", "
        "\"m\"]] is one column."))

    @model_validator(mode="after")
    def _how_many(self):
        if self.count is None and self.values is None:
            raise ValueError("Give `values` (the new columns) or `count` (empty columns).")
        if self.count is not None and self.values is not None and self.count != len(self.values):
            raise ValueError("count and the number of columns in values differ; give one.")
        return self


class DeleteSpanInput(TableInput):
    start: int = Field(..., ge=1, description="First one to delete.")
    end: int | None = Field(None, ge=1, description="Last one (default: just `start`).")


class MergeCellsInput(TableInput):
    from_row: int = Field(..., ge=1)
    from_column: int = Field(..., ge=1)
    to_row: int = Field(..., ge=1)
    to_column: int = Field(..., ge=1)
    text: str | None = Field(None, description=(
        "The merged cell's text. Without it, the cells' texts are kept, one "
        "paragraph each."))

    @model_validator(mode="after")
    def _order(self):
        if self.to_row < self.from_row or self.to_column < self.from_column:
            raise ValueError("to_row/to_column come before from_row/from_column.")
        if (self.to_row, self.to_column) == (self.from_row, self.from_column):
            raise ValueError("That is one cell: nothing to merge.")
        return self


# -- shared ----------------------------------------------------------------------

def _place(count: int, after, before) -> int:
    """The paragraph the table goes in front of (count + 1: the end)."""
    if before is not None:
        if before > count:
            raise ValueError(f"There is no paragraph {before}: the document has {count}.")
        return before
    if after is not None:
        if after > count:
            raise ValueError(f"There is no paragraph {after}: the document has {count}.")
        return after + 1
    return count + 1


def _in_table_error(index: int) -> ValueError:
    return ValueError(
        f"Paragraph {index} is inside a table, and a table there would be a table in a "
        "cell. Put the new table before the table's first paragraph or after its last "
        "(word_list_tables gives them).")


def _rows_from_tokens(text: str, cells_per_row: list[int]) -> list[list[str]]:
    """A table's text cut into rows of cells: each row is its cells, then a row end."""
    tokens = text.split(CELL_END)
    out, i = [], 0
    for n in cells_per_row:
        out.append([_cell_text(t) for t in tokens[i:i + n]])
        i += n + 1
    return out


def _block(n_rows, n_cols, row, column, values):
    last_row = row + len(values) - 1
    if last_row > n_rows:
        raise ValueError(f"The table has {n_rows} rows, and the values reach row "
                         f"{last_row}. Add rows first with word_insert_table_rows.")
    widest = max(len(v) for v in values)
    if column + widest - 1 > n_cols:
        raise ValueError(f"The table has {n_cols} columns, and the values reach column "
                         f"{column + widest - 1}.")


def _unhead(is_heading, unhead, after_row, first, n, total) -> None:
    """Make rows added below the heading rows plain rows.

    A new row copies the row it is added from, heading format and bold
    included; rows added under the headings are data, so they lose both.
    Rows added inside the heading block, or above it, stay headings.
    """
    if after_row == 0:
        return
    try:
        below = first + n
        if below <= total and is_heading(below):
            return
        if not is_heading(after_row):
            return
    except Exception:  # noqa: BLE001 - rows out of reach: the rows were added anyway
        return
    for r in range(first, first + n):
        unhead(r)


def _check_cells(n_rows, n_cols, to_row, to_column):
    if to_row > n_rows:
        raise ValueError(f"The table has {n_rows} rows.")
    if to_column > n_cols:
        raise ValueError(f"The table has {n_cols} columns.")


# -- Windows -----------------------------------------------------------------------

def _table(doc, number: int):
    count = doc.Tables.Count
    if not 1 <= number <= count:
        raise ValueError(f"There is no table {number}: the document has {count} "
                         "tables (numbered from 1; word_list_tables lists them).")
    return doc.Tables.Item(number)


def _span(doc, tb) -> dict:
    rng = tb.Range
    return {"first": index_of(doc, rng), "last": doc.Range(0, rng.End).Paragraphs.Count}


def _grid(tb, start_row: int = 1, limit: int = MAX_ROWS_PER_READ) -> list[list[str]]:
    n_rows = tb.Rows.Count
    if tb.Uniform:
        n_cols = tb.Columns.Count
        rows = _rows_from_tokens(tb.Range.Text, [n_cols] * n_rows)
    else:
        # Rows of a table with merged cells hold different numbers of cells,
        # and Columns(i) cannot be reached (measured), so the cells are walked.
        by_row: dict[int, list[str]] = {}
        cells = tb.Range.Cells
        for i in range(1, cells.Count + 1):
            c = cells.Item(i)
            by_row.setdefault(c.RowIndex, []).append(
                _cell_text(c.Range.Text[:-len(CELL_END)]))
        rows = [by_row.get(r, []) for r in range(1, n_rows + 1)]
    return rows[start_row - 1:start_row - 1 + limit]


@contextmanager
def _columns_reachable(what: str):
    """Answer Word's refusal to reach a column of a table with merged cells.

    Table.Uniform is not asked first: it read False on Windows for a table
    no cell of which had been merged (live check), while Columns(i) worked.
    """
    try:
        yield
    except ValueError:
        raise
    except Exception:  # noqa: BLE001 - Word's refusal, answered
        raise ValueError(f"This table has merged cells, so {what}. Nothing after "
                         "that was changed.") from None


ROWS_UNREACHABLE = ("Word cannot reach its rows one by one (a cell is merged down a "
                    "column)")


def _describe(doc, tb, number: int, cells: bool = True) -> dict:
    out = {"table": number, "rows": tb.Rows.Count, "columns": tb.Columns.Count,
           "style": tb.Style.NameLocal if tb.Style is not None else None,
           "paragraphs": _span(doc, tb)}
    heading = 0
    for r in range(1, min(out["rows"], 10) + 1):
        try:
            if tb.Rows.Item(r).HeadingFormat in (True, -1):
                heading += 1
                continue
        except Exception:  # noqa: BLE001 - rows of merged cells cannot be reached
            pass
        break
    out["header_rows"] = heading
    if cells:
        out["cells"] = _grid(tb)
        if out["rows"] > MAX_ROWS_PER_READ:
            out["more"] = (f"Rows {MAX_ROWS_PER_READ + 1} to {out['rows']} follow; "
                           f"word_get_table with start_row={MAX_ROWS_PER_READ + 1}.")
    return out


def mark_tables(doc, paragraphs: list[dict]) -> None:
    """Give each paragraph that is in a table the table's number."""
    if not paragraphs or not doc.Tables.Count:
        return
    hi = paragraphs[-1]["index"]
    for n in range(1, doc.Tables.Count + 1):
        span = _span(doc, doc.Tables.Item(n))
        if span["first"] > hi:
            break
        for p in paragraphs:
            if span["first"] <= p["index"] <= span["last"]:
                p["table"] = n


def _number_of(doc, tb) -> int:
    return doc.Range(0, tb.Range.Start + 1).Tables.Count


def _set_style(doc, tb, name: str) -> list[str]:
    style = resolve_style(doc, name)
    if style.Type != WD_STYLE_TYPE_TABLE:
        raise ValueError(f"'{name}' is not a table style. word_list_styles lists the "
                         "table styles (type \"table\").")
    key = " ".join(name.lower().split())
    tb.Style = BUILTIN_STYLES[key] if key in BUILTIN_STYLES else style.NameLocal
    return []


def _apply_look(doc, tb, style, header_rows, borders, column_widths, autofit,
                alignment) -> list[str]:
    warnings = []
    if style is not None:
        if header_rows is None:
            header_rows = _describe(doc, tb, 0, cells=False)["header_rows"] or None
        _set_style(doc, tb, style)
    if header_rows is not None:
        with _columns_reachable(ROWS_UNREACHABLE):
            for r in range(1, min(tb.Rows.Count, max(header_rows, 10)) + 1):
                tb.Rows.Item(r).HeadingFormat = r <= header_rows
        tb.ApplyStyleHeadingRows = header_rows > 0
    if borders is not None:
        tb.Borders.Enable = bool(borders)
    if column_widths is not None and autofit is None:
        autofit = "fixed"  # a table fitted to the window stretches the widths given
    if autofit is not None:
        tb.AutoFitBehavior(AUTOFIT[autofit])
    if column_widths is not None:
        if len(column_widths) != tb.Columns.Count:
            raise ValueError(f"column_widths has {len(column_widths)} entries; the table "
                             f"has {tb.Columns.Count} columns.")
        tb.PreferredWidthType = WD_PREFERRED_WIDTH_AUTO  # the columns' widths add up
        for i, width in enumerate(column_widths, start=1):
            with _columns_reachable("its columns have no single width to set. Set widths "
                                    "before merging"):
                col = tb.Columns.Item(i)
            col.PreferredWidthType = WD_PREFERRED_WIDTH_POINTS
            col.PreferredWidth = width
            col.Width = width
    if alignment is not None:
        tb.Rows.Alignment = ROW_ALIGNMENTS[alignment]
    return warnings


def _insert_table_impl(rows, after, before, header_rows, style, borders, column_widths,
                       autofit, alignment, column_alignments):
    from . import text

    doc = backend.word.document()
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    count = doc.Paragraphs.Count
    at = _place(count, after, before)
    if at <= count and paragraph(doc, at).Range.Information(WD_WITHIN_TABLE):
        raise _in_table_error(at)
    after_table = at > 1 and bool(paragraph(doc, at - 1).Range.Information(WD_WITHIN_TABLE))
    # The table goes into an empty paragraph of its own: made in a paragraph
    # with text, it would take that text into its first cell. Right after
    # another table, one more empty paragraph keeps the two apart; two
    # tables with nothing between them are one table.
    blank = [{"text": "", "style": "Normal"}] * (2 if after_table else 1)
    index = text._insert_paragraphs_impl(blank, None,
                                         at if at <= count else None)["inserted"][-1]["index"]
    start = paragraph(doc, index).Range.Start
    tb = doc.Tables.Add(doc.Range(start, start), n_rows, n_cols)
    _drop_paragraph_after(doc, tb)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row, start=1):
            if value:
                tb.Cell(r, c).Range.Text = _text(value)
    if borders is None and style is None:
        borders = True
    if autofit is None:
        autofit = "fixed" if column_widths else "window"
    warnings = _apply_look(doc, tb, style, header_rows, borders, column_widths, autofit,
                           alignment)
    if header_rows and style is None:
        doc.Range(tb.Rows.Item(1).Range.Start,
                  tb.Rows.Item(header_rows).Range.End).Font.Bold = True
    for c, how in enumerate(column_alignments or [], start=1):
        if how:
            for r in range(1, n_rows + 1):
                tb.Cell(r, c).Range.ParagraphFormat.Alignment = ALIGNMENTS[how]
    follow(doc, tb.Range)
    result = {"success": True, **_describe(doc, tb, _number_of(doc, tb)),
              "document_paragraphs": doc.Paragraphs.Count}
    if warnings:
        result["warnings"] = warnings
    return result


def _drop_paragraph_after(doc, tb) -> None:
    """Delete the empty paragraph a new table leaves after it, when it may go.

    A table made in an empty paragraph leaves that paragraph after it
    (measured on both platforms). It stays at the end of the document, which
    must end in a paragraph, and before another table, which it keeps apart.
    """
    end = tb.Range.End
    if end + 1 >= doc.Content.End:
        return
    if doc.Range(end, end + 1).Text != "\r":
        return
    if doc.Range(end + 1, end + 1).Information(WD_WITHIN_TABLE):
        return
    doc.Range(end, end + 1).Delete()


def _list_tables_impl():
    doc = backend.word.document()
    out = []
    for n in range(1, doc.Tables.Count + 1):
        tb = doc.Tables.Item(n)
        info = _describe(doc, tb, n, cells=False)
        first = _grid(tb, 1, 1)
        info["first_row"] = first[0] if first else []
        out.append(info)
    return {"success": True, "tables": out}


def _get_table_impl(table, start_row):
    doc = backend.word.document()
    tb = _table(doc, table)
    info = _describe(doc, tb, table, cells=False)
    if start_row > info["rows"]:
        raise ValueError(f"The table has {info['rows']} rows.")
    info["start_row"] = start_row
    info["cells"] = _grid(tb, start_row)
    last = start_row + len(info["cells"]) - 1
    if last < info["rows"]:
        info["more"] = f"Rows {last + 1} to {info['rows']} follow; pass start_row={last + 1}."
    return {"success": True, **info}


def _set_table_impl(table, style, header_rows, borders, column_widths, autofit, alignment):
    doc = backend.word.document()
    tb = _table(doc, table)
    warnings = _apply_look(doc, tb, style, header_rows, borders, column_widths, autofit,
                           alignment)
    follow(doc, tb.Range)
    result = {"success": True, **_describe(doc, tb, table, cells=False)}
    if warnings:
        result["warnings"] = warnings
    return result


def _set_cells_impl(table, row, column, values):
    doc = backend.word.document()
    tb = _table(doc, table)
    _block(tb.Rows.Count, tb.Columns.Count, row, column, values)
    written, warnings = 0, []
    for r, line in enumerate(values, start=row):
        for c, value in enumerate(line, start=column):
            if value is None:
                continue
            try:
                cell = tb.Cell(r, c)
            except Exception:  # noqa: BLE001 - merged away
                warnings.append(f"There is no cell at row {r}, column {c} (merged).")
                continue
            cell.Range.Text = _text(value)
            got = _cell_text(cell.Range.Text[:-len(CELL_END)])
            if got != _cell_text(_text(value)):
                warnings.append(f"Row {r}, column {c} reads back as {got!r}.")
            written += 1
    follow(doc, tb.Cell(row, column).Range)
    result = {"success": True, "table": table, "written": written,
              "cells": _grid(tb, row, len(values))}
    if warnings:
        result["warnings"] = warnings
    return result


def _cells_in(tb, from_row, to_row, from_column, to_column):
    n_rows, n_cols = tb.Rows.Count, tb.Columns.Count
    to_row, to_column = to_row or n_rows, to_column or n_cols
    _check_cells(n_rows, n_cols, to_row, to_column)
    missing = []
    for r in range(from_row, to_row + 1):
        for c in range(from_column, to_column + 1):
            try:
                cell = tb.Cell(r, c)
            except Exception:  # noqa: BLE001 - merged away
                missing.append((r, c))
                continue
            yield r, c, cell
    if missing:
        yield None, None, missing


def _format_cells_impl(table, from_row, to_row, from_column, to_column, shading, bold, italic,
                       color, size, alignment, vertical_alignment):
    doc = backend.word.document()
    tb = _table(doc, table)
    done, warnings = 0, []
    for r, c, cell in _cells_in(tb, from_row, to_row, from_column, to_column):
        if r is None:
            warnings.append(f"{len(cell)} positions have no cell of their own (merged).")
            continue
        if shading is not None:
            cell.Shading.BackgroundPatternColor = (WD_COLOR_AUTOMATIC if shading == "none"
                                                   else bgr(shading))
        font = cell.Range.Font
        if bold is not None:
            font.Bold = bold
        if italic is not None:
            font.Italic = italic
        if color is not None:
            font.Color = bgr(color)
        if size is not None:
            font.Size = size
        if alignment is not None:
            cell.Range.ParagraphFormat.Alignment = ALIGNMENTS[alignment]
        if vertical_alignment is not None:
            cell.VerticalAlignment = VERTICAL[vertical_alignment]
        done += 1
    first = tb.Cell(from_row, from_column)
    follow(doc, first.Range)
    sample = {"shading": hex_of(first.Shading.BackgroundPatternColor),
              "bold": flag(first.Range.Font.Bold)}
    result = {"success": True, "table": table, "cells": done,
              "first_cell_reads": sample}
    if warnings:
        result["warnings"] = warnings
    return result


def _insert_rows_impl(table, after_row, count, values):
    doc = backend.word.document()
    tb = _table(doc, table)
    n = count or len(values)
    rows, n_cols = tb.Rows.Count, tb.Columns.Count
    after_row = rows if after_row is None else after_row
    if after_row > rows:
        raise ValueError(f"The table has {rows} rows.")
    if values and max(len(v) for v in values) > n_cols:
        raise ValueError(f"A new row has more cells than the table's {n_cols} columns.")
    first = after_row + 1
    with _columns_reachable(ROWS_UNREACHABLE):
        for i in range(n):
            if after_row < rows:
                tb.Rows.Add(tb.Rows.Item(after_row + i + 1))
            else:
                tb.Rows.Add()
        _unhead(lambda r: tb.Rows.Item(r).HeadingFormat in (True, -1),
                lambda r: (setattr(tb.Rows.Item(r), "HeadingFormat", False),
                           setattr(tb.Rows.Item(r).Range.Font, "Bold", False)),
                after_row, first, n, tb.Rows.Count)
    for r, line in enumerate(values or [], start=first):
        for c, value in enumerate(line, start=1):
            tb.Cell(r, c).Range.Text = _text(value)
    follow(doc, tb.Cell(first, 1).Range)
    return {"success": True, "table": table, "inserted_rows": [first, first + n - 1],
            "rows": tb.Rows.Count, "cells": _grid(tb, first, n),
            "paragraphs": _span(doc, tb)}


def _delete_rows_impl(table, start, end):
    doc = backend.word.document()
    tb = _table(doc, table)
    end = end or start
    rows = tb.Rows.Count
    if end < start or end > rows:
        raise ValueError(f"Rows {start} to {end} are not all in the table: it has {rows}.")
    if start == 1 and end == rows:
        raise ValueError("That is every row. Use word_delete_table to delete the table.")
    with _columns_reachable(ROWS_UNREACHABLE):
        for r in range(end, start - 1, -1):
            tb.Rows.Item(r).Delete()
    follow(doc, tb.Range)
    return {"success": True, "table": table, "rows": tb.Rows.Count,
            "paragraphs": _span(doc, tb)}


def _insert_columns_impl(table, after_column, count, values):
    doc = backend.word.document()
    tb = _table(doc, table)
    n = count or len(values)
    cols = tb.Columns.Count
    after_column = cols if after_column is None else after_column
    if after_column > cols:
        raise ValueError(f"The table has {cols} columns.")
    if cols + n > 63:
        raise ValueError("A table has at most 63 columns.")
    with _columns_reachable("Word adds columns only to tables without them. Add "
                            "columns before merging"):
        for i in range(n):
            if after_column < cols:
                tb.Columns.Add(tb.Columns.Item(after_column + i + 1))
            else:
                tb.Columns.Add()
    first = after_column + 1
    for c, column in enumerate(values or [], start=first):
        if len(column) > tb.Rows.Count:
            raise ValueError(f"A new column has {len(column)} cells; the table has "
                             f"{tb.Rows.Count} rows. The columns were added, empty.")
        for r, value in enumerate(column, start=1):
            tb.Cell(r, c).Range.Text = _text(value)
    follow(doc, tb.Range)
    return {"success": True, "table": table, "inserted_columns": [first, first + n - 1],
            "columns": tb.Columns.Count, "cells": _grid(tb), "paragraphs": _span(doc, tb)}


def _delete_columns_impl(table, start, end):
    doc = backend.word.document()
    tb = _table(doc, table)
    end = end or start
    cols = tb.Columns.Count
    if end < start or end > cols:
        raise ValueError(f"Columns {start} to {end} are not all in the table: it has {cols}.")
    if start == 1 and end == cols:
        raise ValueError("That is every column. Use word_delete_table to delete the table.")
    with _columns_reachable("Word cannot reach its columns one by one"):
        for c in range(end, start - 1, -1):
            tb.Columns.Item(c).Delete()
    follow(doc, tb.Range)
    return {"success": True, "table": table, "columns": tb.Columns.Count,
            "cells": _grid(tb), "paragraphs": _span(doc, tb)}


def _merge_cells_impl(table, from_row, from_column, to_row, to_column, text):
    doc = backend.word.document()
    tb = _table(doc, table)
    _check_cells(tb.Rows.Count, tb.Columns.Count, to_row, to_column)
    cell = tb.Cell(from_row, from_column)
    cell.Merge(tb.Cell(to_row, to_column))
    cell = tb.Cell(from_row, from_column)
    if text is not None:
        cell.Range.Text = _text(text)
    follow(doc, cell.Range)
    return {"success": True, "table": table,
            "merged_cell": _cell_text(cell.Range.Text[:-len(CELL_END)]),
            "cells": _grid(tb),
            "paragraphs": _span(doc, tb)}


def _delete_table_impl(table):
    doc = backend.word.document()
    tb = _table(doc, table)
    span = _span(doc, tb)
    tb.Delete()
    return {"success": True, "deleted": table, "was_paragraphs": span,
            "tables": doc.Tables.Count, "paragraphs": doc.Paragraphs.Count}


# -- Mac -------------------------------------------------------------------------------

def _mac_table(doc, number: int):
    from appscript import k

    count = int(doc.count(each=k.table))
    if not 1 <= number <= count:
        raise ValueError(f"There is no table {number}: the document has {count} "
                         "tables (numbered from 1; word_list_tables lists them).")
    return doc.tables[number]


def _mac_span(doc, tb) -> dict:
    from appscript import k

    rng = tb.text_object
    start, end = rng.start_of_content.get(), rng.end_of_content.get()
    return {"first": int(doc.create_range(start=0, end_=start + 1).count(each=k.paragraph)),
            "last": int(doc.create_range(start=0, end_=end).count(each=k.paragraph))}


def _mac_clean(content) -> str:
    return content[:-len(CELL_END)] if isinstance(content, str) and \
        content.endswith(CELL_END) else (content or "")


def _mac_true(value) -> bool:
    """`get range information` answers text: 'true' or 'false' (measured)."""
    return value is True or str(value).lower() == "true"


def _mac_size(tb) -> tuple[int, int]:
    return int(tb.number_of_rows.get()), int(tb.number_of_columns.get())


def _mac_grid(tb, start_row: int = 1, limit: int = MAX_ROWS_PER_READ) -> list[list[str]]:
    from appscript import k

    n_rows, n_cols = _mac_size(tb)
    if tb.uniform.get():
        rows = _rows_from_tokens(tb.text_object.content.get() or "", [n_cols] * n_rows)
        return rows[start_row - 1:start_row - 1 + limit]
    # With merged cells the rows' cell counts cannot be read (a cell merged
    # down a column makes `rows` unreachable, -1728) and the text keeps an
    # empty place for the cells under a merged one, so each cell is asked
    # for. Those under a merged cell answer -1708 and are left out, as
    # Windows leaves them out of a table's cells (measured).
    out = []
    for r in range(start_row, min(n_rows, start_row + limit - 1) + 1):
        row = []
        for c in range(1, n_cols + 1):
            try:
                row.append(_cell_text(_mac_clean(
                    tb.get_cell_from_table(row=r, column=c).text_object.content.get())))
            except Exception:  # noqa: BLE001 - no cell there
                continue
        out.append(row)
    return out


def _mac_describe(doc, tb, number: int, cells: bool = True) -> dict:
    n_rows, n_cols = _mac_size(tb)
    try:
        style = tb.style.name_local.get()
    except Exception:  # noqa: BLE001
        style = None
    out = {"table": number, "rows": n_rows, "columns": n_cols, "style": style,
           "paragraphs": _mac_span(doc, tb)}
    heading = 0
    for r in range(1, min(n_rows, 10) + 1):
        try:
            if tb.rows[r].heading_format.get() is True:
                heading += 1
                continue
        except Exception:  # noqa: BLE001
            pass
        break
    out["header_rows"] = heading
    if cells:
        out["cells"] = _mac_grid(tb)
        if n_rows > MAX_ROWS_PER_READ:
            out["more"] = (f"Rows {MAX_ROWS_PER_READ + 1} to {n_rows} follow; "
                           f"word_get_table with start_row={MAX_ROWS_PER_READ + 1}.")
    return out


def mac_mark_tables(doc, paragraphs: list[dict]) -> None:
    from appscript import k

    count = int(doc.count(each=k.table))
    if not paragraphs or not count:
        return
    hi = paragraphs[-1]["index"]
    for n in range(1, count + 1):
        span = _mac_span(doc, doc.tables[n])
        if span["first"] > hi:
            break
        for p in paragraphs:
            if span["first"] <= p["index"] <= span["last"]:
                p["table"] = n


def _mac_number_of(doc, tb) -> int:
    from appscript import k

    start = tb.text_object.start_of_content.get()
    for n in range(1, int(doc.count(each=k.table)) + 1):
        if doc.tables[n].text_object.start_of_content.get() == start:
            return n
    return 0


def _mac_table_style(doc, name: str):
    """What to set a table's style to: a Word for Mac enumerator, or the style."""
    from appscript import k

    key = " ".join(name.lower().split())
    if BUILTIN_STYLES.get(key) in TABLE_STYLE_CONSTANTS:
        # "medium shading 1 accent 1" is "style table medium shading1 accent1".
        return getattr(k, "style_table_" + re.sub(r" (\d)", r"\1", key).replace(" ", "_"))
    style = doc.Word_styles[name]
    try:
        kind = style.style_type.get()
    except Exception:  # noqa: BLE001
        raise ValueError(
            f"No table style called '{name}' in this document. Use a built-in name that "
            "has a constant (Light Grid, Medium Shading 1 Accent 1...), or a table "
            "style's name as word_list_styles gives it.") from None
    if kind != k.style_type_table:
        raise ValueError(f"'{name}' is not a table style. word_list_styles lists the "
                         "table styles (type \"table\").")
    return style


def _mac_set_style(doc, tb, name: str) -> None:
    value = _mac_table_style(doc, name)
    tb.style.set(value)
    # Word for Mac takes a name it does not know without error and leaves the
    # style as it was (measured with "Table Grid"), so it is read back.
    want = (doc.Word_styles[value] if not hasattr(value, "name_local") else value
            ).name_local.get()
    got = tb.style.name_local.get()
    if got != want:
        raise RuntimeError(f"Word answered without error, but the table's style is "
                           f"'{got}', not '{want}'.")


def _mac_apply_look(doc, tb, style, header_rows, borders, column_widths, autofit,
                    alignment) -> None:
    from appscript import k

    if style is not None:
        # Setting a style on Word for Mac clears the rows' heading format
        # (measured), so the heading rows are put back.
        if header_rows is None:
            header_rows = _mac_describe(doc, tb, 0, cells=False)["header_rows"] or None
        _mac_set_style(doc, tb, style)
    n_rows, n_cols = _mac_size(tb)
    if header_rows is not None:
        with _columns_reachable(ROWS_UNREACHABLE):
            for r in range(1, min(n_rows, max(header_rows, 10)) + 1):
                tb.rows[r].heading_format.set(r <= header_rows)
        tb.apply_style_heading_rows.set(header_rows > 0)
    if borders is not None:
        tb.border_options.enable_borders.set(bool(borders))
    if column_widths is not None and autofit is None:
        autofit = "fixed"  # a table fitted to the window stretches the widths given
    if autofit is not None:
        tb.auto_fit_behavior(behavior=getattr(k, "auto_fit_" + autofit))
    if column_widths is not None:
        if len(column_widths) != n_cols:
            raise ValueError(f"column_widths has {len(column_widths)} entries; the table "
                             f"has {n_cols} columns.")
        if not tb.uniform.get():
            raise ValueError("This table has merged cells, so its columns have no single "
                             "width to set. Set widths before merging.")
        tb.preferred_width_type.set(k.preferred_width_auto)
        for i, width in enumerate(column_widths, start=1):
            col = tb.columns[i]
            col.preferred_width_type.set(k.preferred_width_points)
            col.preferred_width.set(width)
            col.width.set(width)
    if alignment is not None:
        tb.row_options.alignment.set(getattr(k, "align_row_" + alignment))


def _mac_cell(tb, r, c):
    return tb.get_cell_from_table(row=r, column=c)


def _mac_with_active(app, doc, make):
    """Run `make` with `doc` the active document, then give the user's back.

    Word for Mac makes a table only in the active document (-1708 in any
    other, measured), so the target is made active for the moment.
    """
    try:
        before = app.active_document.name.get()
    except Exception:  # noqa: BLE001 - no document active
        before = None
    name = doc.name.get()
    if before == name:
        return make()
    doc.activate_object()
    try:
        return make()
    finally:
        if before:
            try:
                app.documents[before].activate_object()
            except Exception:  # noqa: BLE001 - the table landed either way
                pass


@mac_impl(_insert_table_impl)
def _mac_insert_table(rows, after, before, header_rows, style, borders, column_widths,
                      autofit, alignment, column_alignments):
    from appscript import k

    from . import mac as m
    from . import text
    from ._common import MAC_IMPLS

    word = backend.word
    app = word.get_app()
    doc = word.document()
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    if style is not None:
        _mac_table_style(doc, style)  # an unknown style refuses before anything changes
    count = m._count(doc)
    at = _place(count, after, before)
    def in_table(i):
        return _mac_true(doc.paragraphs[i].text_object.get_range_information(
            information_type=k.with_in_table))

    if at <= count and in_table(at):
        raise _in_table_error(at)
    blank = [{"text": "", "style": "Normal"}] * (2 if at > 1 and in_table(at - 1) else 1)
    index = MAC_IMPLS[text._insert_paragraphs_impl.__name__](
        blank, None, at if at <= count else None)["inserted"][-1]["index"]
    start, _ = m._bounds(doc, index)
    tables_before = int(doc.count(each=k.table))
    _mac_with_active(app, doc, lambda: doc.make(
        new=k.table, at=doc.create_range(start=start, end_=start),
        with_properties={k.number_of_rows: n_rows, k.number_of_columns: n_cols}))
    if int(doc.count(each=k.table)) != tables_before + 1:
        raise RuntimeError("Word answered without error, but no table was made.")
    tb, number = None, 0
    for n in range(1, tables_before + 2):
        if doc.tables[n].text_object.start_of_content.get() >= start:
            tb, number = doc.tables[n], n
            break
    if tb is None:
        raise RuntimeError("Word answered without error, but the table is not where it "
                           "was put.")
    _mac_drop_paragraph_after(doc, tb)
    for r, row in enumerate(rows, start=1):
        for c, value in enumerate(row, start=1):
            if value:
                _mac_cell(tb, r, c).text_object.content.set(_text(value))
    if borders is None and style is None:
        borders = True
    if autofit is None:
        autofit = "fixed" if column_widths else "window"
    _mac_apply_look(doc, tb, style, header_rows, borders, column_widths, autofit, alignment)
    if header_rows and style is None:
        for r in range(1, header_rows + 1):
            tb.rows[r].text_object.bold.set(True)
    for c, how in enumerate(column_alignments or [], start=1):
        if how:
            for r in range(1, n_rows + 1):
                _mac_cell(tb, r, c).text_object.paragraph_format.alignment.set(
                    getattr(k, m.ALIGN_K[how]))
    m._scroll(doc, tb.text_object.start_of_content.get(), tb.text_object.end_of_content.get())
    return {"success": True, **_mac_describe(doc, tb, number),
            "document_paragraphs": m._count(doc)}


def _mac_drop_paragraph_after(doc, tb) -> None:
    from appscript import k

    end = tb.text_object.end_of_content.get()
    if end + 1 >= doc.text_object.end_of_content.get():
        return
    if doc.create_range(start=end, end_=end + 1).content.get() != "\r":
        return
    if _mac_true(doc.create_range(start=end + 1, end_=end + 1).get_range_information(
            information_type=k.with_in_table)):
        return
    doc.create_range(start=end, end_=end + 1).delete()


@mac_impl(_list_tables_impl)
def _mac_list_tables():
    from appscript import k

    doc = backend.word.document()
    out = []
    for n in range(1, int(doc.count(each=k.table)) + 1):
        tb = doc.tables[n]
        info = _mac_describe(doc, tb, n, cells=False)
        first = _mac_grid(tb, 1, 1)
        info["first_row"] = first[0] if first else []
        out.append(info)
    return {"success": True, "tables": out}


@mac_impl(_get_table_impl)
def _mac_get_table(table, start_row):
    doc = backend.word.document()
    tb = _mac_table(doc, table)
    info = _mac_describe(doc, tb, table, cells=False)
    if start_row > info["rows"]:
        raise ValueError(f"The table has {info['rows']} rows.")
    info["start_row"] = start_row
    info["cells"] = _mac_grid(tb, start_row)
    last = start_row + len(info["cells"]) - 1
    if last < info["rows"]:
        info["more"] = f"Rows {last + 1} to {info['rows']} follow; pass start_row={last + 1}."
    return {"success": True, **info}


@mac_impl(_set_table_impl)
def _mac_set_table(table, style, header_rows, borders, column_widths, autofit, alignment):
    from . import mac as m

    doc = backend.word.document()
    tb = _mac_table(doc, table)
    _mac_apply_look(doc, tb, style, header_rows, borders, column_widths, autofit, alignment)
    m._scroll(doc, tb.text_object.start_of_content.get(), tb.text_object.end_of_content.get())
    return {"success": True, **_mac_describe(doc, tb, table, cells=False)}


@mac_impl(_set_cells_impl)
def _mac_set_cells(table, row, column, values):
    from . import mac as m

    doc = backend.word.document()
    tb = _mac_table(doc, table)
    _block(*_mac_size(tb), row, column, values)
    written, warnings = 0, []
    for r, line in enumerate(values, start=row):
        for c, value in enumerate(line, start=column):
            if value is None:
                continue
            try:
                rng = _mac_cell(tb, r, c).text_object
                rng.content.set(_text(value))
            except Exception:  # noqa: BLE001 - merged away
                warnings.append(f"There is no cell at row {r}, column {c} (merged).")
                continue
            got = _cell_text(m._clean(rng.content.get()))
            if got != _cell_text(_text(value)):
                warnings.append(f"Row {r}, column {c} reads back as {got!r}.")
            written += 1
    m._scroll(doc, tb.text_object.start_of_content.get(), tb.text_object.end_of_content.get())
    result = {"success": True, "table": table, "written": written,
              "cells": _mac_grid(tb, row, len(values))}
    if warnings:
        result["warnings"] = warnings
    return result


@mac_impl(_format_cells_impl)
def _mac_format_cells(table, from_row, to_row, from_column, to_column, shading, bold, italic,
                      color, size, alignment, vertical_alignment):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    tb = _mac_table(doc, table)
    n_rows, n_cols = _mac_size(tb)
    to_row, to_column = to_row or n_rows, to_column or n_cols
    _check_cells(n_rows, n_cols, to_row, to_column)
    done, missing = 0, 0
    for r in range(from_row, to_row + 1):
        for c in range(from_column, to_column + 1):
            cell = _mac_cell(tb, r, c)
            try:
                rng = cell.text_object
                rng.start_of_content.get()
            except Exception:  # noqa: BLE001 - merged away
                missing += 1
                continue
            if shading == "none":
                # A cell never shaded reads index `no highlight` (measured).
                cell.shading.background_pattern_color_index.set(k.no_highlight)
            elif shading is not None:
                v = shading.lstrip("#")
                cell.shading.background_pattern_color.set(
                    [int(v[i:i + 2], 16) * 257 for i in (0, 2, 4)])
            if bold is not None:
                rng.bold.set(bold)
            if italic is not None:
                rng.italic.set(italic)
            if color is not None:
                v = color.lstrip("#")
                rng.font_object.color.set([int(v[i:i + 2], 16) * 257 for i in (0, 2, 4)])
            if size is not None:
                rng.font_object.font_size.set(size)
            if alignment is not None:
                rng.paragraph_format.alignment.set(getattr(k, m.ALIGN_K[alignment]))
            if vertical_alignment is not None:
                cell.vertical_alignment.set(getattr(k, "cell_align_vertical_" +
                                                    vertical_alignment))
            done += 1
    first = _mac_cell(tb, from_row, from_column)
    raw = first.shading.background_pattern_color.get()
    if first.shading.background_pattern_color_index.get() == k.no_highlight:
        raw = None  # [0, 0, 0] then means no shading, not black
    bold_now = first.text_object.bold.get()
    sample = {"shading": (f"#{round(raw[0] / 257):02X}{round(raw[1] / 257):02X}"
                          f"{round(raw[2] / 257):02X}"
                          if isinstance(raw, (list, tuple)) and len(raw) == 3 else None),
              "bold": bold_now if isinstance(bold_now, bool) else None}
    m._scroll(doc, tb.text_object.start_of_content.get(), tb.text_object.end_of_content.get())
    result = {"success": True, "table": table, "cells": done, "first_cell_reads": sample}
    if missing:
        result["warnings"] = [f"{missing} positions have no cell of their own (merged)."]
    return result


@mac_impl(_insert_rows_impl)
def _mac_insert_rows(table, after_row, count, values):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    tb = _mac_table(doc, table)
    n = count or len(values)
    rows, n_cols = _mac_size(tb)
    after_row = rows if after_row is None else after_row
    if after_row > rows:
        raise ValueError(f"The table has {rows} rows.")
    if values and max(len(v) for v in values) > n_cols:
        raise ValueError(f"A new row has more cells than the table's {n_cols} columns.")
    # `make new row at end of rows` is refused (-1708); after a row, or before
    # the first, works (measured).
    first = after_row + 1
    with _columns_reachable(ROWS_UNREACHABLE):
        for i in range(n):
            if after_row == 0:
                doc.make(new=k.row, at=tb.rows[1].before)
            else:
                doc.make(new=k.row, at=tb.rows[after_row + i].after)
        _unhead(lambda r: tb.rows[r].heading_format.get() is True,
                lambda r: (tb.rows[r].heading_format.set(False),
                           tb.rows[r].text_object.bold.set(False)),
                after_row, first, n, _mac_size(tb)[0])
    for r, line in enumerate(values or [], start=first):
        for c, value in enumerate(line, start=1):
            _mac_cell(tb, r, c).text_object.content.set(_text(value))
    m._scroll(doc, tb.text_object.start_of_content.get(), tb.text_object.end_of_content.get())
    return {"success": True, "table": table, "inserted_rows": [first, first + n - 1],
            "rows": _mac_size(tb)[0], "cells": _mac_grid(tb, first, n),
            "paragraphs": _mac_span(doc, tb)}


@mac_impl(_delete_rows_impl)
def _mac_delete_rows(table, start, end):
    doc = backend.word.document()
    tb = _mac_table(doc, table)
    end = end or start
    rows = _mac_size(tb)[0]
    if end < start or end > rows:
        raise ValueError(f"Rows {start} to {end} are not all in the table: it has {rows}.")
    if start == 1 and end == rows:
        raise ValueError("That is every row. Use word_delete_table to delete the table.")
    with _columns_reachable(ROWS_UNREACHABLE):
        for r in range(end, start - 1, -1):
            tb.rows[r].delete()
    return {"success": True, "table": table, "rows": _mac_size(tb)[0],
            "paragraphs": _mac_span(doc, tb)}


@mac_impl(_insert_columns_impl)
def _mac_insert_columns(table, after_column, count, values):
    raise ValueError(
        "Word for Mac's scripting cannot add a column to a table (`make new column` is "
        "refused, measured). Make the table again with the columns it needs: read it "
        "with word_get_table, insert the new one with word_insert_table, then delete "
        "the old one with word_delete_table.")


@mac_impl(_delete_columns_impl)
def _mac_delete_columns(table, start, end):
    doc = backend.word.document()
    tb = _mac_table(doc, table)
    end = end or start
    cols = _mac_size(tb)[1]
    if end < start or end > cols:
        raise ValueError(f"Columns {start} to {end} are not all in the table: it has {cols}.")
    if start == 1 and end == cols:
        raise ValueError("That is every column. Use word_delete_table to delete the table.")
    if not tb.uniform.get():
        raise ValueError("This table has merged cells, so Word cannot reach its columns "
                         "one by one.")
    for c in range(end, start - 1, -1):
        tb.columns[c].delete()
    return {"success": True, "table": table, "columns": _mac_size(tb)[1],
            "cells": _mac_grid(tb), "paragraphs": _mac_span(doc, tb)}


@mac_impl(_merge_cells_impl)
def _mac_merge_cells(table, from_row, from_column, to_row, to_column, text):
    from . import mac as m

    doc = backend.word.document()
    tb = _mac_table(doc, table)
    _check_cells(*_mac_size(tb), to_row, to_column)
    # Merging makes the document the active one (measured), so the user's
    # active document is given back afterwards.
    _mac_with_active(backend.word.get_app(), doc, lambda: _mac_cell(
        tb, from_row, from_column).merge_cell(with_=_mac_cell(tb, to_row, to_column)))
    cell = _mac_cell(tb, from_row, from_column)
    if text is not None:
        cell.text_object.content.set(_text(text))
    return {"success": True, "table": table,
            "merged_cell": _cell_text(m._clean(cell.text_object.content.get())),
            "cells": _mac_grid(tb),
            "paragraphs": _mac_span(doc, tb)}


@mac_impl(_delete_table_impl)
def _mac_delete_table(table):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    tb = _mac_table(doc, table)
    span = _mac_span(doc, tb)
    tb.delete()
    return {"success": True, "deleted": table, "was_paragraphs": span,
            "tables": int(doc.count(each=k.table)), "paragraphs": m._count(doc)}


# -- tools -----------------------------------------------------------------------------

def register(mcp) -> None:
    @mcp.tool(name="word_insert_table", annotations=annotations("Insert a table", WRITES))
    async def word_insert_table(params: InsertTableInput) -> str:
        """Insert a table filled with the given cells, after or before a
        paragraph (default: at the end). It spans the text width with grid
        lines and a bold heading row unless a style, widths or autofit say
        otherwise. Answers its number, its cells and the paragraphs it covers."""
        p = params
        return await call(_insert_table_impl, p.rows, p.after, p.before, p.header_rows,
                          p.style, p.borders, p.column_widths, p.autofit, p.alignment,
                          p.column_alignments, undo="Insert table")

    @mcp.tool(name="word_list_tables", annotations=annotations("List tables", READ_ONLY))
    async def word_list_tables(params: Input) -> str:
        """The document's tables: size, style, heading rows, first row and the
        paragraphs each covers."""
        return await call(_list_tables_impl)

    @mcp.tool(name="word_get_table", annotations=annotations("Read a table", READ_ONLY))
    async def word_get_table(params: GetTableInput) -> str:
        """A table's cells as text, row by row (100 rows at a time), with its
        size and style. Rows with merged cells hold fewer cells, and a cell's
        column number is its place in that row."""
        return await call(_get_table_impl, params.table, params.start_row)

    @mcp.tool(name="word_set_table", annotations=annotations("Change a table", WRITES))
    async def word_set_table(params: SetTableInput) -> str:
        """Change a table as a whole: style, heading rows, grid lines, column
        widths, autofit, position between the margins."""
        p = params
        return await call(_set_table_impl, p.table, p.style, p.header_rows, p.borders,
                          p.column_widths, p.autofit, p.alignment, undo="Change table")

    @mcp.tool(name="word_set_table_cells", annotations=annotations("Write table cells", WRITES))
    async def word_set_table_cells(params: SetCellsInput) -> str:
        """Write the text of a block of cells, from `row`/`column` on, and read
        them back."""
        p = params
        return await call(_set_cells_impl, p.table, p.row, p.column, p.values,
                          undo="Write table cells")

    @mcp.tool(name="word_format_table_cells",
              annotations=annotations("Format table cells", WRITES))
    async def word_format_table_cells(params: FormatCellsInput) -> str:
        """Format a block of cells (default: all): background color, bold,
        italic, text color and size, horizontal and vertical alignment."""
        p = params
        return await call(_format_cells_impl, p.table, p.from_row, p.to_row, p.from_column,
                          p.to_column, p.shading, p.bold, p.italic, p.color, p.size,
                          p.alignment, p.vertical_alignment, undo="Format table cells")

    @mcp.tool(name="word_insert_table_rows",
              annotations=annotations("Insert table rows", WRITES))
    async def word_insert_table_rows(params: InsertRowsInput) -> str:
        """Add rows to a table, empty or filled, after a row (default: at the
        bottom)."""
        p = params
        return await call(_insert_rows_impl, p.table, p.after_row, p.count, p.values,
                          undo="Insert table rows")

    @mcp.tool(name="word_delete_table_rows",
              annotations=annotations("Delete table rows", DESTRUCTIVE))
    async def word_delete_table_rows(params: DeleteSpanInput) -> str:
        """Delete rows `start` to `end` of a table."""
        p = params
        return await call(_delete_rows_impl, p.table, p.start, p.end,
                          undo="Delete table rows")

    @mcp.tool(name="word_insert_table_columns",
              annotations=annotations("Insert table columns", WRITES))
    async def word_insert_table_columns(params: InsertColumnsInput) -> str:
        """Add columns to a table, empty or filled, after a column (default: at
        the right). Windows only: Word for Mac's scripting cannot add columns."""
        p = params
        return await call(_insert_columns_impl, p.table, p.after_column, p.count, p.values,
                          undo="Insert table columns")

    @mcp.tool(name="word_delete_table_columns",
              annotations=annotations("Delete table columns", DESTRUCTIVE))
    async def word_delete_table_columns(params: DeleteSpanInput) -> str:
        """Delete columns `start` to `end` of a table."""
        p = params
        return await call(_delete_columns_impl, p.table, p.start, p.end,
                          undo="Delete table columns")

    @mcp.tool(name="word_merge_table_cells",
              annotations=annotations("Merge table cells", WRITES))
    async def word_merge_table_cells(params: MergeCellsInput) -> str:
        """Merge the cells from one corner to the other into one cell, keeping
        their texts or putting `text` in it. Afterwards a row holds fewer
        cells, and a cell's column number is its place in the row as
        word_get_table shows it. Set column widths and add columns before
        merging; a cell merged down a column stops rows being added or
        deleted."""
        p = params
        return await call(_merge_cells_impl, p.table, p.from_row, p.from_column, p.to_row,
                          p.to_column, p.text, undo="Merge table cells")

    @mcp.tool(name="word_delete_table", annotations=annotations("Delete a table", DESTRUCTIVE))
    async def word_delete_table(params: TableInput) -> str:
        """Delete a table and its contents."""
        return await call(_delete_table_impl, params.table, undo="Delete table")
