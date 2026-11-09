import json

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _common, notes


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def insert(index, text, kind="footnote", after=None):
    return answer(notes._insert_note_impl, index, text, kind, after)


def listed(kind="footnote", start=1, limit=100):
    return answer(notes._list_notes_impl, kind, start, limit)


def test_a_note_goes_at_the_end_of_the_paragraphs_text(make_doc):
    doc = make_doc("One two.", "Three.")
    out = insert(1, "A note")
    assert out == {"success": True, "kind": "footnote", "number": 1, "index": 1, "text": "A note"}
    assert doc.texts() == ["One two.\x02", "Three."]


def test_a_note_goes_right_after_the_text_asked_for(make_doc):
    doc = make_doc("One two three.")
    insert(1, "A note", after="two")
    assert doc.texts() == ["One two\x02 three."]


def test_after_is_found_past_an_earlier_mark(make_doc):
    doc = make_doc("One two three.")
    insert(1, "first", after="One")
    insert(1, "second", after="two")
    assert doc.texts() == ["One\x02 two\x02 three."]


def test_text_that_is_not_there_is_refused_and_nothing_changes(make_doc):
    doc = make_doc("One two.")
    out = insert(1, "A note", after="nine")
    assert "'nine' is not in paragraph 1" in out["error"]
    assert doc.texts() == ["One two."] and doc.Footnotes.Count == 0


def test_notes_are_numbered_in_the_order_of_their_marks(make_doc):
    make_doc("One two.", "Three.")
    assert insert(2, "later")["number"] == 1
    assert insert(1, "earlier")["number"] == 1
    out = listed()
    assert [(n["number"], n["index"], n["text"]) for n in out["notes"]] == [
        (1, 1, "earlier"), (2, 2, "later")]


def test_footnotes_and_endnotes_are_counted_apart(make_doc):
    make_doc("One.", "Two.")
    insert(1, "a foot")
    insert(2, "an end", "endnote")
    out = listed("endnote")
    assert (out["footnotes"], out["endnotes"]) == (1, 1)
    assert [n["text"] for n in out["notes"]] == ["an end"]


def test_a_long_list_says_where_to_go_on(make_doc):
    make_doc("One.")
    for t in "abc":
        insert(1, t)
    out = listed(limit=2)
    assert [n["text"] for n in out["notes"]] == ["a", "b"]
    assert out["more"] == "Notes 3 to 3 follow; pass start=3."
    rest = listed(start=3, limit=2)
    assert [n["text"] for n in rest["notes"]] == ["c"] and "more" not in rest


def test_a_notes_text_changes_and_its_mark_stays(make_doc):
    doc = make_doc("One.")
    insert(1, "old")
    out = answer(notes._set_note_impl, 1, "new", "footnote")
    assert out["text"] == "new" and doc.texts() == ["One.\x02"]


def test_deleting_a_note_takes_its_mark_and_renumbers(make_doc):
    doc = make_doc("One.", "Two.")
    insert(1, "a")
    insert(2, "b")
    out = answer(notes._delete_note_impl, 1, "footnote")
    assert out == {"success": True, "kind": "footnote", "deleted": 1, "left": 1}
    assert doc.texts() == ["One.", "Two.\x02"]
    assert [(n["number"], n["text"]) for n in listed()["notes"]] == [(1, "b")]


@pytest.mark.parametrize("impl, args", [
    (notes._set_note_impl, (3, "x", "footnote")),
    (notes._delete_note_impl, (3, "footnote"))])
def test_a_note_that_is_not_there_is_refused(make_doc, impl, args):
    make_doc("One.")
    insert(1, "only")
    out = answer(impl, *args)
    assert "There is no footnote 3: the document has 1" in out["error"]
    assert listed()["footnotes"] == 1


@pytest.mark.parametrize("text", ["two\nlines", "two\rlines"])
def test_a_note_is_one_paragraph(text):
    with pytest.raises(ValidationError):
        notes.InsertNoteInput(index=1, text=text)
    with pytest.raises(ValidationError):
        notes.SetNoteInput(number=1, text=text)
