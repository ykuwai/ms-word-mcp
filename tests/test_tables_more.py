"""Borders one side at a time, sorting and splitting cells (tools/tables.py)."""

from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _props as P
from ms_word_mcp.tools import tables

LINE = {"style": "single", "width": 1.5, "color": "#1F4E79"}
NONE = {"style": "none"}


# -- which sides a cell gets ---------------------------------------------------------

def test_outer_sides_go_on_the_edge_and_inside_ones_between():
    borders = {"top": LINE, "bottom": LINE, "inside_horizontal": NONE, "left": LINE}
    rows, columns = (1, 3), (1, 2)
    assert tables.cell_sides(borders, 1, 1, rows, columns) == {"top": LINE, "bottom": NONE,
                                                               "left": LINE}
    assert tables.cell_sides(borders, 2, 2, rows, columns) == {"top": NONE, "bottom": NONE}
    assert tables.cell_sides(borders, 3, 2, rows, columns) == {"top": NONE, "bottom": LINE}


def test_one_row_takes_both_its_outer_sides():
    assert tables.cell_sides({"top": LINE, "bottom": NONE}, 2, 1, (2, 2), (1, 3)) == {
        "top": LINE, "bottom": NONE}


def test_a_rule_under_a_heading_row_only():
    """bottom on row 1 alone: the next row's top is not touched."""
    assert tables.cell_sides({"bottom": LINE}, 1, 2, (1, 1), (1, 3)) == {"bottom": LINE}


def test_formatting_asks_for_something_borders_included():
    with pytest.raises(ValidationError, match="borders"):
        tables.FormatCellsInput(table=1)
    assert tables.FormatCellsInput(table=1, borders={"bottom": LINE}).borders.bottom.width == 1.5


def test_a_sort_and_a_split_are_checked():
    with pytest.raises(ValidationError):
        tables.SortTableInput(table=1, by=[])
    with pytest.raises(ValidationError, match="nothing to split"):
        tables.SplitCellInput(table=1, row=1, column=1, rows=1, columns=1)


# -- Windows, against a stand-in table ------------------------------------------------

class _Line:
    def __init__(self):
        self.LineStyle, self.LineWidth, self.Color = 0, 8, -16777216


class _Cell:
    def __init__(self, text):
        sides = {code: _Line() for code in P.BORDER_SIDES.values()}
        self.Borders = SimpleNamespace(Item=lambda code: sides[code])
        self.Range = SimpleNamespace(Text=text + "\r\x07", Font=SimpleNamespace(Bold=0))
        self.Shading = SimpleNamespace(BackgroundPatternColor=-16777216)
        self.split = None

    def Split(self, NumRows, NumColumns):
        self.split = (NumRows, NumColumns)


class _Table:
    def __init__(self, rows):
        self.cells = [[_Cell(t) for t in row] for row in rows]
        self.Rows = SimpleNamespace(Count=len(rows), Item=lambda i: SimpleNamespace(
            HeadingFormat=-1 if i == 1 else 0, Cells=SimpleNamespace(Count=len(rows[i - 1]))))
        self.Columns = SimpleNamespace(Count=len(rows[0]))
        self.sorted = None
        self.Range = "table range"

    def Cell(self, r, c):
        return self.cells[r - 1][c - 1]

    def Sort(self, **kwargs):
        self.sorted = kwargs


@pytest.fixture
def com_table(monkeypatch):
    tb = _Table([["Item", "Qty"], ["Pear", "10"], ["Apple", "3"]])
    monkeypatch.setattr(tables.backend, "word", SimpleNamespace(document=lambda: None))
    monkeypatch.setattr(tables, "_table", lambda d, n: tb)
    monkeypatch.setattr(tables, "_grid", lambda t: [[c.Range.Text[:-2] for c in row] for row in t.cells])
    monkeypatch.setattr(tables, "_span", lambda d, t: {"first": 1, "last": 9})
    monkeypatch.setattr(tables, "follow", lambda d, r: None)
    return tb


def test_on_windows_a_rule_goes_under_the_heading_row(com_table):
    out = tables._format_cells_impl(1, 1, 1, 1, None, None, None, None, None, None, None, None,
                                    {"bottom": {"style": "double", "width": 1.5}})
    for c in (1, 2):
        line = com_table.Cell(1, c).Borders.Item(P.BORDER_SIDES["bottom"])
        assert (line.LineStyle, line.LineWidth) == (P.BORDER_STYLES["double"], P.BORDER_WIDTHS[1.5])
    assert com_table.Cell(2, 1).Borders.Item(P.BORDER_SIDES["top"]).LineStyle == 0
    assert out["first_cell_reads"]["borders"]["bottom"]["style"] == "double"
    assert "warnings" not in out


def test_on_windows_a_sort_keeps_the_heading_row_and_takes_three_keys(com_table):
    out = tables._sort_table_impl(1, [{"column": 2, "order": "descending", "kind": "number"},
                                      {"column": 1, "order": "ascending", "kind": "text"}], None)
    assert com_table.sorted == {"ExcludeHeader": True, "FieldNumber": 2, "SortFieldType": 1,
                                "SortOrder": 1, "FieldNumber2": 1, "SortFieldType2": 0,
                                "SortOrder2": 0}
    assert out["header_kept"] is True


def test_on_windows_a_column_past_the_last_is_refused(com_table):
    with pytest.raises(ValueError, match="There is no column 3"):
        tables._sort_table_impl(1, [{"column": 3, "order": "ascending", "kind": "text"}], False)
    assert com_table.sorted is None


def test_on_windows_a_cell_is_split(com_table):
    tables._split_cell_impl(1, 2, 2, 1, 3)
    assert com_table.Cell(2, 2).split == (1, 3)


def test_a_block_whose_first_position_was_merged_away_reads_its_first_real_cell(com_table):
    real = com_table.Cell

    def cell(r, c):
        if (r, c) == (1, 1):
            raise RuntimeError("merged away")
        return real(r, c)
    com_table.Cell = cell
    out = tables._format_cells_impl(1, 1, 1, 1, 2, None, True, None, None, None, None, None)
    assert out["cells"] == 1 and "no cell of their own" in out["warnings"][0]


def test_a_block_with_no_cell_of_its_own_is_refused(com_table):
    com_table.Cell = lambda r, c: (_ for _ in ()).throw(RuntimeError("merged away"))
    with pytest.raises(ValueError, match="No cell of that block"):
        tables._format_cells_impl(1, 1, 1, 1, 1, None, True, None, None, None, None, None)


def test_a_cell_past_the_last_row_is_refused_before_word_picks_one(com_table):
    """Word answers a cell of the last row for a row past it, on both platforms."""
    with pytest.raises(ValueError, match="There is no cell at row 9.*Nothing was changed"):
        tables._split_cell_impl(1, 9, 1, 1, 2)
    with pytest.raises(ValueError, match="There is no cell at row 2, column 3"):
        tables._split_cell_impl(1, 2, 3, 1, 2)
    assert all(c.split is None for row in com_table.cells for c in row)
