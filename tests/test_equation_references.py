"""References to numbered equations: text, not math, and "(n)" as the equation shows it."""

from ms_word_mcp.tools import references


def test_a_reference_to_an_equation_takes_the_look_of_the_text_around_it():
    assert references._reference_field("_Ref1", "number", True) == "REF _Ref1 \\h \\* CHARFORMAT"
    assert references._reference_field("_Ref1", "label_and_number", True).endswith("CHARFORMAT")
    assert references._reference_field("_Ref1", "number") == "REF _Ref1 \\h"
    assert references._reference_field("_Ref1", "page", True) == "PAGEREF _Ref1 \\h"


def test_an_equation_has_no_label_so_its_label_and_number_is_its_number_in_parentheses():
    assert references.equation_reference({"result": "2"}, "label_and_number")["result"] == "(2)"
    assert references.equation_reference({"result": "2"}, "number")["result"] == "2"
