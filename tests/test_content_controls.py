"""Content controls: the bridge's wire format, how a control is answered, and the inputs."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import content_controls as cc


def test_plain_ascii_travels_as_it_is_and_the_rest_encoded():
    assert cc.wire("MENDELEY_CITATION_v3_eyJj+/==") == "r:MENDELEY_CITATION_v3_eyJj+/=="
    assert cc.wire("תג").startswith("e:")
    for text in ("", "plain", "תג-עברי", "a|b", "line\nbreak", "en–dash"):
        assert cc.unwire(cc.wire(text)) == text


def test_a_long_tag_is_cut_unless_asked_for_whole():
    control = {"nth": 1, "index": 3, "type": "rich_text", "text": "[1]", "title": "", "tag": "T" * 500,
               "lock_control": False, "lock_contents": True}
    short = cc.view(control, False)
    assert short["tag"].endswith("…") and len(short["tag"]) == 201 and short["tag_length"] == 500
    assert short["lock_contents"] is True and "lock_control" not in short and "title" not in short
    assert cc.view(control, True)["tag"] == "T" * 500


def test_the_inputs_check_what_can_be_done():
    with pytest.raises(ValidationError, match="not both"):
        cc.InsertContentControlInput(index=1, after="a", before="b")
    with pytest.raises(ValidationError, match="at least one"):
        cc.SetContentControlInput(control=1)
    with pytest.raises(ValidationError):
        cc.InsertContentControlInput(index=1, kind="date")
    assert cc.DeleteContentControlInput(control=2).keep_text is True


def test_around_takes_text_already_there_alone():
    from ms_word_mcp.tools import content_controls as cc

    assert cc.InsertContentControlInput(index=1, around="bold part").around == "bold part"
    for extra in ({"text": "x"}, {"after": "y"}, {"before": "y"}):
        with pytest.raises(ValidationError, match="around"):
            cc.InsertContentControlInput(index=1, around="bold part", **extra)
