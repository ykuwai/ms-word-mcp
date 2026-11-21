"""Inline equations: among a paragraph's words, found again by their place in it."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import equations


def test_which_equation_in_the_paragraph_the_new_one_is():
    assert equations.nth_of([10, 40], 10) == 1
    assert equations.nth_of([10, 40], 25) == 2
    assert equations.nth_of([40, 10], 41) == 2  # past the last: the last


def test_an_equation_word_made_display_says_so():
    inline = equations.inline_answer(1, 2, False, "x", "𝑥")
    assert inline["inline"] is True and "warnings" not in inline
    display = equations.inline_answer(1, 1, True, "x", "𝑥")
    assert display["inline"] is False and "line of its own" in display["warnings"][0]


def test_an_inline_equation_is_one_line_and_its_place_is_checked():
    assert equations.InsertInlineEquationInput(index=1, equation="x^2", match="X").match == "X"
    with pytest.raises(ValidationError, match="one line"):
        equations.InsertInlineEquationInput(index=1, equation="a\nb")
    with pytest.raises(ValidationError):
        equations.GetEquationInput(index=1, nth=0)
    assert equations.SetEquationInput(index=1, equation="y").nth == 1
