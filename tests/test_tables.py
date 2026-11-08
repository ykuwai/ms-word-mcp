"""Tables: the parts that do not need Word."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import tables


def test_a_tables_text_is_cut_into_rows_of_cells():
    # Each cell ends in \r\x07, and so does each row after its cells.
    text = "Item\r\x07Qty\r\x07\r\x07Two\rlines\r\x07\r\x07\r\x07"
    assert tables._rows_from_tokens(text, [2, 2]) == [["Item", "Qty"], ["Two\nlines", ""]]


def test_rows_of_merged_cells_hold_fewer_cells():
    text = "Total\r\x07\r\x07a\r\x07b\r\x07\r\x07"
    assert tables._rows_from_tokens(text, [1, 2]) == [["Total"], ["a", "b"]]


def _unheaded(heading_rows, after_row, n, total):
    done = []
    tables._unhead(lambda r: r in heading_rows, done.append, after_row, after_row + 1, n,
                   total)
    return done


def test_rows_added_under_the_headings_are_data():
    assert _unheaded({1}, 1, 2, 3) == [2, 3]


def test_rows_added_inside_or_above_the_headings_stay_headings():
    assert _unheaded({1, 2, 3}, 1, 1, 4) == []
    assert _unheaded({1, 2}, 0, 1, 3) == []


def test_rows_added_among_data_are_left_alone():
    assert _unheaded({1}, 3, 1, 5) == []


def test_insert_checks_its_shape():
    ok = tables.InsertTableInput(rows=[["a", "b"], ["c"]], column_widths=[100, 50])
    assert ok.header_rows == 1
    with pytest.raises(ValidationError, match="2 columns"):
        tables.InsertTableInput(rows=[["a", "b"]], column_widths=[100])
    with pytest.raises(ValidationError, match="not both"):
        tables.InsertTableInput(rows=[["a"]], after=1, before=2)
    with pytest.raises(ValidationError, match="header_rows"):
        tables.InsertTableInput(rows=[["a"]], header_rows=2)


def test_format_cells_needs_something_to_do():
    with pytest.raises(ValidationError, match="at least one"):
        tables.FormatCellsInput(table=1)
    assert tables.FormatCellsInput(table=1, shading="none").shading == "none"


def test_new_lines_in_a_cell_become_line_breaks():
    assert tables._text("a\nb\r\nc") == "a\vb\vc"
