import json

import pytest

from ms_word_mcp.tools import _common, _props, formatting


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def fmt(start, end=None, match=None, **font):
    return answer(formatting._format_text_impl, start, end, match, font)


def para(start, end=None, **paragraph):
    return answer(formatting._set_paragraph_format_impl, start, end, paragraph)


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


def test_an_east_asian_font_word_refuses_comes_back_as_a_warning(make_doc):
    make_doc("日本語")
    out = fmt(1, east_asian_font="Segoe UI", bold=True)
    assert out["success"]
    assert any("east_asian_font" in w and "refused" in w for w in out["warnings"])
    # The rest of the call still applied.
    assert out["read_back"][0]["bold"] is True


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


def test_every_character_property_is_set_and_read_back(make_doc):
    make_doc("text")
    out = fmt(1, underline="double", underline_color="#00B050", color="#C00000",
              strikethrough=True, small_caps=True, character_spacing=1.5, scaling=90,
              kerning_from=10.5, position=2, emphasis_mark="dot_above",
              ignore_character_grid=True)
    assert "warnings" not in out, out.get("warnings")
    got = out["read_back"][0]
    assert got["underline"] == "double" and got["color"] == "#C00000"
    assert got["emphasis_mark"] == "dot_above" and got["scaling"] == 90


def test_color_is_written_as_word_bgr():
    assert _props.bgr("#FF0000") == 0x0000FF
    assert _props.hex_of(0x0000FF) == "#FF0000"
    assert _props.hex_of(-16777216) is None


def test_paragraph_format_line_spacing_in_lines_and_points(make_doc):
    make_doc("a", "b")
    out = para(1, 2, alignment="justify", line_spacing_rule="multiple", line_spacing=1.2,
               space_before=6, space_after=0, snap_to_grid=False)
    first = out["paragraphs"][0]
    assert first["alignment"] == "justify"
    assert first["line_spacing_rule"] == "multiple"
    assert first["line_spacing_lines"] == 1.2
    assert first["space_before"] == 6 and first["snap_to_grid"] is False
    assert "warnings" not in out
    out = para(1, line_spacing_rule="exactly", line_spacing=18)
    assert out["paragraphs"][0]["line_spacing_pt"] == 18


def test_one_and_a_half_lines_reads_back_as_lines(make_doc):
    make_doc("a")
    out = para(1, line_spacing_rule="multiple", line_spacing=1.5)
    assert out["paragraphs"][0]["line_spacing_rule"] == "one_and_half"
    assert out["paragraphs"][0]["line_spacing_lines"] == 1.5
    assert "warnings" not in out


def test_japanese_paragraph_settings(make_doc):
    doc = make_doc("本文")
    out = para(1, space_before_lines=0.5, first_line_indent_chars=1,
               east_asian_line_breaking=True, hanging_punctuation=False,
               space_between_east_asian_and_latin=False)
    assert "warnings" not in out, out.get("warnings")
    mark = doc.attrs[-1]
    assert mark["LineUnitBefore"] == 0.5 and mark["CharacterUnitFirstLineIndent"] == 1
    assert mark["HangingPunctuation"] is False


def test_snap_to_grid_writes_the_inverted_member(make_doc):
    doc = make_doc("a")
    para(1, snap_to_grid=False)
    assert doc.attrs[-1]["DisableLineHeightGrid"] is True


def test_paragraph_spec_checks_rule_value_and_units():
    with pytest.raises(ValueError):
        formatting.ParagraphFormatInput(start=1, line_spacing=12)
    with pytest.raises(ValueError):
        formatting.ParagraphFormatInput(start=1, line_spacing_rule="exactly")
    with pytest.raises(ValueError, match="not both"):
        formatting.ParagraphFormatInput(start=1, space_before=6, space_before_lines=0.5)
    with pytest.raises(ValueError, match="at least one"):
        formatting.ParagraphFormatInput(start=1)
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
    with pytest.raises(ValueError):
        formatting.FormatTextInput(start=1, superscript=True, subscript=True)


def set_style(style, create=None, base_style=None, next_style=None, font=None,
              paragraph=None, **flags):
    return answer(formatting._set_style_impl, style, create, base_style, next_style, font,
                  paragraph, flags.get("quick_style"), flags.get("priority"),
                  flags.get("automatically_update"),
                  flags.get("no_space_between_same_style"))


def test_create_a_paragraph_style_with_every_part(make_doc):
    doc = make_doc("a")
    out = set_style("本文インデント", create="paragraph", base_style="Normal",
                    next_style="Normal",
                    font={"east_asian_font": "游明朝", "size": 10.5},
                    paragraph={"first_line_indent_chars": 1, "space_after_lines": 0.5,
                               "line_spacing_rule": "exactly", "line_spacing": 18},
                    quick_style=True, priority=1, no_space_between_same_style=True)
    assert out["created"] and "warnings" not in out, out.get("warnings")
    style = out["style"]
    assert style["font"]["east_asian_font"] == "游明朝"
    assert style["paragraph"]["first_line_indent_chars"] == 1
    assert style["paragraph"]["line_spacing_pt"] == 18
    assert style["quick_style"] is True and style["no_space_between_same_style"] is True
    assert doc.Styles.Item("本文インデント").Font.Size == 10.5


def test_create_refuses_an_existing_name_and_change_refuses_a_missing_one(make_doc):
    make_doc("a")
    assert "already exists" in set_style("Normal", create="paragraph")["error"]
    assert "No style called" in set_style("Nope", font={"size": 12})["error"]


def test_change_a_builtin_style_by_its_english_name(make_doc):
    doc = make_doc("a")
    out = set_style("Heading 1", font={"east_asian_font": "BIZ UDPゴシック", "bold": True},
                    paragraph={"space_before": 12, "keep_with_next": True})
    assert "warnings" not in out, out.get("warnings")
    assert doc.Styles.Item(-2).ParagraphFormat.SpaceBefore == 12


def test_a_character_style_has_no_paragraph_formatting(make_doc):
    make_doc("a")
    with pytest.raises(ValueError):
        formatting.SetStyleInput(style="x", create="character", paragraph={"space_before": 1})
    out = set_style("強調", create="character", font={"color": "#C00000"})
    assert out["style"]["type"] == "character" and "paragraph" not in out["style"]


def test_get_style_answers_in_set_style_names(make_doc):
    make_doc("a")
    out = answer(formatting._get_style_impl, "Normal")
    style = out["style"]
    assert style["name"] == "Normal" and style["base_style"] is None
    assert set(style["font"]) <= set(_props.FONT_NAMES)
    assert "line_spacing_rule" in style["paragraph"]


def test_a_new_style_does_not_take_the_formatting_where_the_cursor_is(make_doc):
    doc = make_doc(("Big bold title", "Title"), "body")
    doc.attrs[0]["Bold"] = -1
    doc.attrs[0]["Size"] = 28
    out = set_style("Plain", create="paragraph", font={"color": "#333333"})
    style = out["style"]
    assert style["base_style"] == "Normal"
    assert style["font"]["bold"] is False and style["font"]["size"] == 10.5


def test_points_take_over_from_characters():
    from ms_word_mcp.tools import _props

    before, after = _props.paired_steps({"first_line_indent": 0.0})
    assert before == {"first_line_indent_chars": 0.0} and after == {}
    before, after = _props.paired_steps({"first_line_indent_chars": 0})
    assert before == {} and after == {"first_line_indent": 0.0}
    assert _props.paired_steps({"first_line_indent_chars": 1}) == ({}, {})


def test_a_style_set_to_zero_characters_also_gets_zero_points(make_doc):
    doc = make_doc("a")
    style = doc.Styles.Item(-2)
    style.ParagraphFormat.CharacterUnitFirstLineIndent = 1
    style.ParagraphFormat.FirstLineIndent = 10.5
    set_style("Heading 1", paragraph={"first_line_indent_chars": 0})
    assert style.ParagraphFormat.FirstLineIndent == 0


def test_a_paragraph_can_run_right_to_left(make_doc):
    doc = make_doc("مرحبا")
    out = para(1, direction="right_to_left", alignment="right")
    assert "warnings" not in out, out.get("warnings")
    assert doc.attrs[-1]["ReadingOrder"] == 0


def test_bold_and_size_reach_right_to_left_text(make_doc):
    doc = make_doc("abc")
    fmt(1, bold=True, size=14)
    assert doc.attrs[0]["BoldBi"] == -1 and doc.attrs[0]["SizeBi"] == 14
    fmt(1, size=12, size_complex_script=16)
    assert doc.attrs[0]["SizeBi"] == 16
