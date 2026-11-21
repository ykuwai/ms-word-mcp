"""Captions and equations numbered with their chapter: Figure 2.1, (2.1)."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import equations, references


def test_the_fields_of_a_number_with_its_chapter():
    assert references.chapter_codes("Figure", 1) == ("STYLEREF 1 \\s", "SEQ Figure \\* ARABIC \\s 1")
    assert references.chapter_codes("Equation", 2)[1].endswith("\\s 2")


def test_a_number_reads_with_its_chapter_without_word_for_macs_marks():
    assert references.joined_number("‎2", ".", "1") == "2.1"
    assert references.joined_number(None, "", "3") == "3"
    assert references.no_marks("‏2‎") == "2"


def test_a_chapter_number_that_is_not_one_is_warned_of():
    assert references.chapter_warning("‎2", 1) == []
    assert references.chapter_warning("A", 1) == []  # a lettered heading
    said = references.chapter_warning("Error! No text of specified style in document.", 2)
    assert "numbered Heading 2" in said[0]


def test_an_equation_is_written_with_a_place_for_its_chapter():
    assert equations.number_mark(None) == "#(1)"
    assert equations.number_mark({"level": 1, "separator": "-"}) == "#(1-1)"


def test_a_chapter_number_is_for_a_numbered_equation_and_a_known_separator():
    made = equations.InsertEquationInput(equation="x", numbered=True, chapter={"level": 2})
    assert made.chapter.separator == "."
    with pytest.raises(ValidationError, match="give `numbered` too"):
        equations.InsertEquationInput(equation="x", chapter={"level": 1})
    with pytest.raises(ValidationError):
        references.InsertCaptionInput(after=1, chapter={"level": 1, "separator": "/"})
