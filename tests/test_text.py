import json
from types import SimpleNamespace

import pytest

from ms_word_mcp.tools import _common, mac, tables, text


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def test_insert_into_a_new_document_fills_its_empty_paragraph(make_doc):
    doc = make_doc()
    out = answer(text._insert_paragraphs_impl,
                 [{"text": "Title", "style": "Title"}, {"text": "Body", "style": None}],
                 None, None)
    assert out["success"]
    assert doc.texts() == ["Title", "Body"]
    assert doc.styles()[0] == "Title"
    assert [p["index"] for p in out["inserted"]] == [1, 2]


def test_insert_at_the_end_keeps_what_was_there(make_doc):
    doc = make_doc("one", "two")
    out = answer(text._insert_paragraphs_impl,
                 [{"text": "three", "style": "Heading 1"}], None, None)
    assert doc.texts() == ["one", "two", "three"]
    assert out["inserted"] == [{"index": 3, "text": "three", "style": "Heading 1",
                                "outline_level": 1}]


def test_insert_after_and_before(make_doc):
    doc = make_doc("one", "two", "three")
    answer(text._insert_paragraphs_impl, [{"text": "1.5", "style": None}], 1, None)
    assert doc.texts() == ["one", "1.5", "two", "three"]
    answer(text._insert_paragraphs_impl, [{"text": "zero", "style": None}], None, 1)
    assert doc.texts() == ["zero", "one", "1.5", "two", "three"]
    answer(text._insert_paragraphs_impl, [{"text": "start", "style": None}], 0, None)
    assert doc.texts()[0] == "start"


def test_insert_with_a_style_clears_manual_formatting(make_doc):
    doc = make_doc("one")
    doc.attrs[0]["Bold"] = -1
    doc.attrs[3]["Bold"] = -1  # the paragraph mark carries it on
    answer(text._insert_paragraphs_impl, [{"text": "new", "style": "Heading 1"}], None, 1)
    first = doc.paragraph_objects()[0]
    assert first.Range.Font.Bold == 0


def test_insert_scrolls_to_the_new_text_without_touching_a_selection(word, make_doc):
    doc = make_doc("one")
    answer(text._insert_paragraphs_impl, [{"text": "two", "style": None}], None, None)
    assert doc.window.scrolled_to, "the window should follow the edit"


def test_insert_refuses_a_place_past_the_end(make_doc):
    make_doc("one")
    out = answer(text._insert_paragraphs_impl, [{"text": "x", "style": None}], 5, None)
    assert "error" in out and "paragraph 5" in out["error"]


def test_set_paragraph_text_keeps_style(make_doc):
    doc = make_doc(("Old heading", "Heading 1"), "body")
    out = answer(text._set_paragraph_text_impl, 1, "New heading")
    assert doc.texts() == ["New heading", "body"]
    assert doc.styles()[0] == "Heading 1"
    assert out["was"] == "Old heading"
    assert out["paragraph"]["text"] == "New heading"


def test_delete_paragraphs_in_the_middle_and_at_the_end(make_doc):
    doc = make_doc("a", "b", "c", "d")
    out = answer(text._delete_paragraphs_impl, 2, 3)
    assert doc.texts() == ["a", "d"]
    assert out["texts"] == ["b", "c"] and "warnings" not in out
    answer(text._delete_paragraphs_impl, 2, None)
    assert doc.texts() == ["a"]


def test_delete_refuses_a_span_outside_the_document(make_doc):
    make_doc("a")
    out = answer(text._delete_paragraphs_impl, 1, 3)
    assert "error" in out


def test_find_replace_counts_and_replaces(make_doc):
    doc = make_doc("cat and cat", "dog", "cat")
    found = answer(text._find_replace_impl, "cat", None, False, False, False)
    assert found["matches"] == 3 and found["paragraphs"] == [1, 3]
    assert doc.texts()[0] == "cat and cat"
    out = answer(text._find_replace_impl, "cat", "lion", False, False, False)
    assert out["replaced"] == 3
    assert doc.texts() == ["lion and lion", "dog", "lion"]


def test_get_paragraphs_and_outline(make_doc):
    make_doc(("Intro", "Heading 1"), "text", ("Part", "Heading 2"), "more")
    outline = answer(text._get_outline_impl)
    assert [(h["index"], h["level"]) for h in outline["headings"]] == [(1, 1), (3, 2)]
    out = answer(text._get_paragraphs_impl, 2, 3, True)
    assert [p["index"] for p in out["paragraphs"]] == [2, 3]
    assert out["paragraphs"][0]["fonts"]["east_asian"] == "游明朝"
    assert out["paragraphs"][0]["line_spacing_rule"] == "single"


def test_paragraph_models_refuse_paragraph_marks():
    with pytest.raises(ValueError):
        text.NewParagraph(text="a\rb")
    with pytest.raises(ValueError):
        text.InsertParagraphsInput(paragraphs=[{"text": "a"}], after=1, before=2)


def test_styles_go_on_runs_of_paragraphs():
    items = [{"text": "a", "style": "Heading 1"}, {"text": "b", "style": None},
             {"text": "c", "style": "Normal"}, {"text": "d", "style": "Normal"},
             {"text": "e", "style": "Heading 1"}]
    assert text.style_runs(items) == [("Heading 1", 0, 0), ("Normal", 2, 3),
                                      ("Heading 1", 4, 4)]


def test_more_than_500_matches_are_all_replaced(make_doc):
    doc = make_doc("cat " * 501)
    out = answer(text._find_replace_impl, "cat", "dog", False, False, False)
    assert (out["matches"], out["replaced"]) == (501, 501)
    assert "cat" not in doc.texts()[0]


def test_a_search_that_stops_moving_changes_nothing(make_doc, monkeypatch):
    doc = make_doc("cat")

    class Stuck:
        Start, End = 0, 3

        def ClearFormatting(self):
            pass

        def Execute(self, **kwargs):
            return True

        @property
        def Find(self):
            return self

        Text = "cat"

        def Collapse(self, direction):
            pass

    monkeypatch.setattr(type(doc), "Content", property(lambda self: Stuck()))
    with pytest.raises(ValueError, match="Nothing was changed"):
        text._find_all(doc, "cat", False, False, False, float("inf"))


def test_spaces_in_text_arguments_are_kept():
    replace = text.FindReplaceInput(find="cat ", replace=" ")
    assert (replace.find, replace.replace) == ("cat ", " ")
    assert text.NewParagraph(text="　indented ").text == "　indented "
    assert text.SetParagraphTextInput(index=1, text=" a ").text == " a "


def test_spaces_around_a_name_are_not():
    assert text.NewParagraph(text="a", style=" Heading 1 ").style == "Heading 1"


def test_a_search_out_of_time_says_there_are_more(make_doc, monkeypatch):
    make_doc("cat cat cat")
    monkeypatch.setattr(text, "SEARCH_SECONDS", -1.0)
    out = answer(text._find_replace_impl, "cat", None, False, False, False)
    assert (out["matches"], out["complete"]) == (1, False)
    assert "still searching" in out["warnings"][0] and "at least 1" in out["warnings"][0]


def test_a_replacement_out_of_time_says_how_far_it_got(make_doc, monkeypatch):
    doc = make_doc("cat cat cat")
    monkeypatch.setattr(text, "WORK_SECONDS", -1.0)
    out = answer(text._find_replace_impl, "cat", "dog", False, False, False)
    assert (out["matches"], out["replaced"], out["complete"]) == (3, 0, False)
    assert "Call again" in out["warnings"][0]
    assert doc.texts() == ["cat cat cat"]


def test_calling_again_is_not_offered_when_it_would_replace_twice(make_doc, monkeypatch):
    make_doc("cat cat")
    monkeypatch.setattr(text, "WORK_SECONDS", -1.0)
    out = answer(text._find_replace_impl, "cat", "catfish", False, False, False)
    assert "narrow the search" in out["warnings"][0]


def test_calling_again_is_not_offered_for_a_wildcard_pattern(make_doc, monkeypatch):
    make_doc("cat cat")
    monkeypatch.setattr(text, "WORK_SECONDS", -1.0)
    out = answer(text._find_replace_impl, "cat", "dog", False, False, True)
    assert out["complete"] is False
    assert "wildcard pattern can match" in out["warnings"][0]
    assert "same arguments" not in out["warnings"][0]


def test_a_read_out_of_time_answers_what_it_has(make_doc, monkeypatch):
    make_doc("one", "two", "three")
    monkeypatch.setattr(text, "READ_SECONDS", -1.0)
    out = answer(text._get_paragraphs_impl, 1, 3, False)
    assert [p["text"] for p in out["paragraphs"]] == ["one"]
    assert out["complete"] is False and "start=2" in out["warnings"][0]
    assert "pass start=2" in out["more"]


def test_a_read_in_time_is_not_marked(make_doc):
    make_doc("one", "two", "three")
    out = answer(text._get_paragraphs_impl, 1, 2, False)
    assert "complete" not in out and "warnings" not in out
    assert "pass start=3" in out["more"]


def test_a_mac_read_out_of_time_answers_what_it_has(monkeypatch):
    monkeypatch.setattr(mac.backend, "word", SimpleNamespace(document=lambda: None))
    monkeypatch.setattr(mac, "_count", lambda doc: 3)
    monkeypatch.setattr(mac, "_info", lambda doc, i, detail=False: {"index": i})
    monkeypatch.setattr(tables, "mac_mark_tables", lambda doc, paragraphs: None)
    monkeypatch.setattr(text, "READ_SECONDS", -1.0)
    out = mac.get_paragraphs(1, 3, False)
    assert [p["index"] for p in out["paragraphs"]] == [1]
    assert out["complete"] is False and "pass start=2" in out["more"]


def _mac_text(content):
    value = lambda v: SimpleNamespace(get=lambda: v)  # noqa: E731
    return SimpleNamespace(text_object=SimpleNamespace(
        content=value(content), end_of_content=value(len(content.encode("utf-16-le")) // 2)))


def test_a_mac_search_out_of_time_stops_on_a_plain_document():
    found, finished = mac._matches_until(_mac_text("cat cat cat"), "cat", False, False,
                                         float("-inf"))
    assert (found, finished) == ([(0, 3)], False)


def test_a_mac_search_counts_positions_past_an_emoji():
    found, finished = mac._matches_until(_mac_text("😀cat 😀 cat"), "cat", False, False, None)
    assert (found, finished) == ([(2, 5), (9, 12)], True)


def test_only_the_first_matches_are_given_paragraph_numbers(make_doc):
    make_doc(*["cat"] * 101)
    out = answer(text._find_replace_impl, "cat", None, False, False, False)
    assert out["matches"] == 101 and out["paragraphs"] == list(range(1, 101))
    assert out["paragraphs_cover"] == "the first 100 matches"
