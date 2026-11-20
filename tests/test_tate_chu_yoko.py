"""Tate-chu-yoko: a few characters set across in vertical text, by word_format_text."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import formatting


def test_it_is_read_back_by_name_and_warns_when_word_did_not_take_it():
    assert formatting.across_answer(None, None) == ({}, [])
    assert formatting.across_answer("fit_in_line", 1) == ({"horizontal_in_vertical": "fit_in_line"}, [])
    got, warnings = formatting.across_answer("fit_in_line", 0)
    assert got == {"horizontal_in_vertical": "none"} and "not 'fit_in_line'" in warnings[0]


def test_it_alone_is_enough_to_format_text():
    made = formatting.FormatTextInput(start=1, match="12", horizontal_in_vertical="resize_line")
    assert made.horizontal_in_vertical == "resize_line"
    with pytest.raises(ValidationError):
        formatting.FormatTextInput(start=1, horizontal_in_vertical="sideways")
    with pytest.raises(ValidationError, match="at least one"):
        formatting.FormatTextInput(start=1)


def test_the_mac_names_cover_every_choice():
    assert set(formatting.HORIZONTAL_IN_VERTICAL_MAC) == set(formatting.HORIZONTAL_IN_VERTICAL)
