import json

import pytest

from ms_word_mcp.tools import _common, formatting


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def fmt(start, end=None, match=None, **kw):
    args = dict(latin_font=None, east_asian_font=None, complex_script_font=None, size=None,
                size_complex_script=None, bold=None, italic=None, underline=None,
                color=None)
    args.update(kw)
    return answer(formatting._format_text_impl, start, end, match, *args.values())


def test_three_scripts_each_get_their_own_font(make_doc):
    doc = make_doc("日本語と English", "two")
    out = fmt(1, latin_font="Segoe UI", east_asian_font="BIZ UDPゴシック",
              complex_script_font="Arial")
    assert "warnings" not in out
    font = doc.paragraph_objects()[0].Range.Font
    assert (font.NameAscii, font.NameOther, font.NameFarEast, font.NameBi) == (
        "Segoe UI", "Segoe UI", "BIZ UDPゴシック", "Arial")
    # The second paragraph is untouched.
    assert doc.paragraph_objects()[1].Range.Font.NameAscii == "Calibri"


def test_an_east_asian_font_word_ignores_comes_back_as_a_warning(make_doc):
    make_doc("日本語")
    out = fmt(1, east_asian_font="Segoe UI")
    assert out["success"]
    assert any("East Asian" in w for w in out["warnings"])


def test_match_formats_only_the_occurrences(make_doc):
    doc = make_doc("a key word and a key")
    out = fmt(1, match="key", bold=True)
    assert out["ranges"] == 2
    bold = [i for i, a in enumerate(doc.attrs) if a["Bold"]]
    assert bold == [2, 3, 4, 17, 18, 19]


def test_match_that_is_not_there_changes_nothing(make_doc):
    make_doc("abc")
    out = fmt(1, match="zzz", bold=True)
    assert "error" in out and "Nothing was changed" in out["error"]


def test_color_is_written_as_word_bgr():
    assert formatting._bgr("#FF0000") == 0x0000FF
    assert formatting._hex(0x0000FF) == "#FF0000"
    assert formatting._hex(-16777216) is None


def test_paragraph_format_line_spacing_in_lines_and_points(make_doc):
    make_doc("a", "b")
    out = answer(formatting._set_paragraph_format_impl, 1, 2, "justify", "multiple", 1.5,
                 6, 0, None, None, None, False, None, None)
    first = out["paragraphs"][0]
    assert first["alignment"] == "justify"
    assert first["line_spacing_rule"] == "multiple"
    assert first["line_spacing_lines"] == 1.5
    assert first["space_before"] == 6
    out = answer(formatting._set_paragraph_format_impl, 1, None, None, "exactly", 18,
                 None, None, None, None, None, None, None, None)
    assert out["paragraphs"][0]["line_spacing_pt"] == 18


def test_paragraph_format_model_checks_rule_and_value():
    with pytest.raises(ValueError):
        formatting.ParagraphFormatInput(start=1, line_spacing=12)
    with pytest.raises(ValueError):
        formatting.ParagraphFormatInput(start=1, line_spacing_rule="exactly")
    formatting.ParagraphFormatInput(start=1, line_spacing_rule="double")


def test_apply_style_by_english_builtin_name(make_doc):
    doc = make_doc("a", "b")
    out = answer(formatting._apply_style_impl, 1, 2, "heading 2", False)
    assert doc.styles() == ["Heading 2", "Heading 2"]
    assert out["style"] == "Heading 2" and "warnings" not in out


def test_apply_unknown_style_says_what_exists(make_doc):
    make_doc("a")
    out = answer(formatting._apply_style_impl, 1, None, "Fancy", False)
    assert "No style called 'Fancy'" in out["error"]


def test_format_text_needs_something_to_set():
    with pytest.raises(ValueError):
        formatting.FormatTextInput(start=1)
