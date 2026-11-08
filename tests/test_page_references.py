"""Pictures, captions, references, headers and ruby: the parts that do not need Word."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import page, references
from ms_word_mcp.tools._word import ruby_of, ruby_to_drop, with_ruby


def test_page_placeholders_become_one_character_and_a_field():
    plain, fields = page.parse_placeholders("Page {page} of {pages}")
    assert plain == "Page # of #"
    assert fields == [(5, "PAGE"), (10, "NUMPAGES")]
    assert page.parse_placeholders("none") == ("none", [])
    assert page.parse_placeholders("a\nb")[0] == "a\vb"


def _reader(word):
    return lambda a, b: word[a:b]


def test_spans_are_found_in_order_and_read_back():
    raw = "漢字と文章"
    assert page.locate_spans(raw, ["漢字", "文章"], _reader(raw), 0, 0) == [(0, 2), (3, 5)]


def test_spans_are_found_behind_hidden_characters():
    # Word holds 10 characters of a ruby field where the text shows one mark.
    raw = "A\x15B漢字"
    word = "A" + "#" * 10 + "B漢字"
    assert page.locate_spans(raw, ["漢字"], _reader(word), 0, 9) == [(12, 14)]


def test_a_span_out_of_order_changes_nothing():
    raw = "abc"
    with pytest.raises(ValueError, match="Nothing was changed"):
        page.locate_spans(raw, ["c", "a"], _reader(raw), 0, 0)


def test_ruby_is_read_out_of_its_field():
    code = 'EQ \\* jc2 \\* "Font:游明朝" \\* hps11 \\o\\ad(\\s\\up 10(かんじ),漢字)'
    assert ruby_of(code) == ("漢字", "かんじ")
    assert ruby_of(" PAGE ") is None
    text, rubies = with_ruby("\x15と\x15", [code, " PAGE ", code.replace("漢字", "文章")])
    assert text == "漢字と文章"
    assert [r["ruby"] for r in rubies] == ["かんじ", "かんじ"]


def test_a_label_is_one_word_for_its_field():
    assert references.seq_name("図") == "図"
    assert references.seq_name("Fig ure") == "Fig_ure"


def test_a_missing_side_follows_in_proportion():
    assert references.proportional(100, None, 200, 50) == (100, 25.0)
    assert references.proportional(None, 25, 200, 50) == (100.0, 25)
    assert references.proportional(10, 20, 200, 50) == (10, 20)


def test_a_caption_has_one_place():
    with pytest.raises(ValidationError, match="one of"):
        references.InsertCaptionInput(text="x")
    with pytest.raises(ValidationError, match="one of"):
        references.InsertCaptionInput(text="x", for_picture=1, after=2)
    assert references.InsertCaptionInput(text="x", for_table=1).label == "Figure"


def test_page_setup_needs_something_to_do():
    with pytest.raises(ValidationError, match="at least one"):
        page.SetPageSetupInput(section=1)
    with pytest.raises(ValidationError, match="not both"):
        page.SetPageSetupInput(paper="a4", width=100)


_CODE = r"EQ \* jc2 \* \"Font:Yu Mincho\" \* hps5 \o\ad(\s\up 9({}),{})"


def _ruby(base, reading):
    return " " + _CODE.replace("\\\"", '"').format(reading, base) + " "


def test_ruby_to_be_replaced_is_found():
    codes = [_ruby("漢字", "かんじ"), " PAGE ", _ruby("文章", "ぶんしょう")]
    assert ruby_of(codes[0]) == ("漢字", "かんじ")
    assert ruby_to_drop("\x15と\x15", codes, ["文章"]) == [(2, "文章")]


def test_a_missing_span_is_refused_before_any_ruby_comes_off():
    codes = [_ruby("漢字", "かんじ")]
    with pytest.raises(ValueError, match="Nothing was changed"):
        ruby_to_drop("\x15です", codes, ["漢字", "ない"])


def test_spans_out_of_order_are_refused_before_any_ruby_comes_off():
    codes = [_ruby("漢字", "かんじ"), _ruby("文章", "ぶんしょう")]
    with pytest.raises(ValueError, match="in the order"):
        ruby_to_drop("\x15と\x15", codes, ["文章", "漢字"])


def test_a_span_cannot_cross_ruby_that_stays():
    codes = [_ruby("漢字", "かんじ")]
    with pytest.raises(ValueError, match="Nothing was changed"):
        ruby_to_drop("\x15です", codes, ["字です"])


def test_spaces_of_a_caption_and_of_ruby_are_kept():
    assert references.InsertCaptionInput(after=1, separator=": ").separator == ": "
    assert references.InsertCaptionInput(after=1, separator=" ").separator == " "
    assert references.InsertCaptionInput(after=1, label=" 図 ").label == "図"
    assert page.RubySpan(text="a b ", ruby=" x").text == "a b "
