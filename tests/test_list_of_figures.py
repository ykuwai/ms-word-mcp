"""A list of figures, or of tables: a table of contents of one caption label."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import fields


def test_the_field_lists_one_labels_captions():
    assert fields.figures_code("Figure", True, True) == 'TOC \\h \\z \\c "Figure"'
    assert fields.figures_code("図", False, False) == 'TOC \\z \\c "図" \\n'


def test_a_label_of_two_words_is_named_as_its_captions_are():
    """word_insert_caption names the SEQ field with references.seq_name."""
    from ms_word_mcp.tools.references import seq_name

    assert f'\\c "{seq_name("Data Table")}"' in fields.figures_code("Data Table", True, True)


def test_one_place_only():
    with pytest.raises(ValidationError, match="not both"):
        fields.InsertFiguresInput(after=1, before=3)
    assert fields.InsertFiguresInput().label == "Figure"


def test_an_empty_list_says_where_captions_come_from():
    assert "word_insert_caption" in fields.no_captions("Table")


def test_a_list_of_figures_is_told_from_a_table_of_contents():
    assert fields.is_list_of_figures('TOC \\h \\z \\c "Figure"')
    assert fields.is_list_of_figures(' toc \\C "Table" ')
    assert not fields.is_list_of_figures('TOC \\o "1-3" \\h \\z \\u')
    assert not fields.is_list_of_figures('SEQ Figure \\* ARABIC')


def test_with_no_caption_of_the_label_words_own_sentence_is_not_an_entry():
    """Word fills an empty list with "No table of figures entries found." in its
    own language; the captions are counted instead."""
    summary = {"code": 'TOC \\c "Chart"', "entries": 1,
               "first_entries": ["No table of figures entries found."]}
    out = fields.figures_answer(3, "Chart", summary, [{"label": "Figure"}], 9)
    assert (out["entries"], out["first_entries"]) == (0, []) and "Chart" in out["warnings"][0]
    figures = {"code": 'TOC \\c "Figure"', "entries": 2, "first_entries": ["Figure 1: a 1", "Figure 2: b 2"]}
    out = fields.figures_answer(3, "Figure", figures, [{"label": "Figure"}] * 2, 9)
    assert out["entries"] == 2 and "warnings" not in out
