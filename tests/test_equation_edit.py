"""Changing an equation in place: the new text goes between the array's
opening and the number, which stays."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import equations

# A numbered equation, linearized, as Word holds it: the array opening, the
# equation, "#(", the SEQ field (begin mark, code, result, end mark), ")".
LINEAR = "█(E=mc^2#(\x13SEQ Equation\x141\x15) )"


def read(a, b):
    return LINEAR[a:b]


def test_the_number_is_found_before_its_field():
    code_start = LINEAR.index("SEQ")
    stop = equations.kept_number_at(code_start, read)
    assert LINEAR[:stop] == "█(E=mc^2"


def test_a_number_that_is_not_where_it_was_put_is_not_guessed():
    assert equations.kept_number_at(3, read) is None


def test_the_new_text_goes_after_the_arrays_opening():
    assert LINEAR[equations.own_text_start(0, read):].startswith("E=mc^2")
    assert equations.own_text_start(0, lambda a, b: "x+y"[a:b]) == 0


def test_the_input_is_one_line_without_a_number():
    with pytest.raises(ValidationError, match="without its number"):
        equations.SetEquationInput(index=2, equation="x#(3)")
    with pytest.raises(ValidationError, match="one line"):
        equations.SetEquationInput(index=2, equation="a\nb")
    assert equations.SetEquationInput(index=2, equation="\\frac{a}{b}").format == "latex"
