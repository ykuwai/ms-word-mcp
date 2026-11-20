"""Picture bullets: a list level whose bullet is a picture."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import lists


def test_a_picture_level_is_a_bullet_first():
    spec = lists._level_spec_defaults({"level": 1, "picture": "/p/dot.png"})
    assert (spec["number_style"], spec["number_format"], spec["picture"]) == ("bullet", "\u2022", "/p/dot.png")


def test_a_picture_is_the_bullet_and_a_full_path():
    assert lists.LevelSpec(picture="/p/dot.png").picture == "/p/dot.png"
    assert lists.LevelSpec(picture="C:\\p\\dot.png").picture == "C:\\p\\dot.png"
    with pytest.raises(ValidationError, match="full path"):
        lists.LevelSpec(picture="dot.png")
    with pytest.raises(ValidationError, match="without `bullet`"):
        lists.LevelSpec(picture="/p/dot.png", bullet="●")
    with pytest.raises(ValidationError, match="without `bullet`"):
        lists.LevelSpec(picture="/p/dot.png", number_style="arabic")


def test_a_picture_word_did_not_take_is_warned_of():
    specs = [{"level": 1, "picture": "/p/dot.png"}]
    assert lists.picture_warnings(specs, ["picture_bullet", "picture_bullet"]) == []
    assert lists.picture_warnings(specs, ["bullet"]) == []  # a level of an outline list
    assert "not picture bullets" in lists.picture_warnings(specs, ["simple_numbering"])[0]
    assert lists.picture_warnings([{"level": 1, "bullet": "●"}], ["simple_numbering"]) == []
