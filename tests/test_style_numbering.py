"""A style's numbering: reading it, joining another style's, taking it off."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _vba, formatting, mac
from ms_word_mcp.tools.lists import NUMBER_STYLES

from test_mac_gaps import _Model, word  # noqa: F401 - the fixture


LEVELS = [(1, NUMBER_STYLES["arabic"], "%1.", "Heading 1"), (2, NUMBER_STYLES["arabic"], "%1.%2", "Heading 2"),
          (3, NUMBER_STYLES["lower_roman"], "%1.%2.%3", None)]


def test_a_style_numbering_says_its_level_and_the_lists_styles():
    assert formatting.style_list(2, LEVELS) == {
        "level": 2, "number_style": "arabic", "number_format": "%1.%2",
        "linked_styles": {"1": "Heading 1", "2": "Heading 2"}}


@pytest.mark.parametrize("asked, got, says", [
    ("none", None, None),
    ("none", {"level": 1}, "still has numbering"),
    ({"like": "Heading 1", "level": 3}, {"level": 3}, None),
    ({"like": "Heading 1", "level": 3}, None, "at list level None, not 3"),
])
def test_numbering_read_back(asked, got, says):
    out = formatting.numbering_warnings(asked, got, "Heading 3")
    assert (out == []) if says is None else (says in out[0])


def test_numbering_is_for_paragraph_styles_and_counts_as_something_to_set():
    assert formatting.SetStyleInput(style="Heading 3", numbering="none").numbering == "none"
    joined = formatting.SetStyleInput(style="Heading 3", numbering={"like": "Heading 1", "level": 3})
    assert joined.numbering.level == 3
    with pytest.raises(ValidationError, match="apply to paragraph styles"):
        formatting.SetStyleInput(style="X", create="character",
                                 numbering={"like": "Heading 1", "level": 2})
    with pytest.raises(ValidationError):
        formatting.SetStyleInput(style="X", numbering={"like": "Heading 1", "level": 10})


def test_on_the_mac_a_style_numbering_is_read_through_the_bridge(word):  # noqa: F811
    levels = [{"NumberStyle": n, "NumberFormat": f, "LinkedStyle": s or ""} for _, n, f, s in LEVELS]
    template = {"ListLevels": {"Count": 3, "Item": lambda i: levels[i - 1]}}
    model = word({"Styles": {"Item": lambda name: {"ListTemplate": template if name == "Heading 2" else None,
                                                   "ListLevelNumber": 2 if name == "Heading 2" else 0}}})
    assert mac._mac_style_list(None, "Heading 2") == formatting.style_list(2, LEVELS)
    assert mac._mac_style_list(None, "Body Text") is None
    assert len(model.requests) == 4  # three for the numbered style (its levels counted), one for the other


def test_an_object_a_call_answers_is_reached_by_its_arguments(word):  # noqa: F811
    word({"Levels": {"Item": lambda i: {"Format": f"level {i}"}}})
    level = _vba.Remote("/d/Report.docx", "Levels").call("Item", 3)
    assert level._path == "Levels/Item(n3)" and level.Format == "level 3"
