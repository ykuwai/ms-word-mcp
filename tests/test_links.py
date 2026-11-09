import json

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _common, links


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def link(index, address=None, bookmark=None, match=None, text=None, tip=None):
    return answer(links._insert_hyperlink_impl, index, address, bookmark, match, text, tip)


def mark(name, index, match=None):
    return answer(links._add_bookmark_impl, name, index, match)


def test_a_link_over_some_text(make_doc):
    doc = make_doc("One two three.")
    out = link(1, address="https://example.com", match="two", tip="A tip")
    assert out == {"success": True, "number": 1, "index": 1, "text": "two",
                   "address": "https://example.com", "bookmark": None, "screen_tip": "A tip"}
    assert doc.texts() == ["One two three."]


def test_a_link_over_a_whole_paragraph(make_doc):
    make_doc("One two.")
    assert link(1, address="https://example.com")["text"] == "One two."


def test_a_link_as_new_text_at_the_end(make_doc):
    doc = make_doc("See ")
    assert link(1, address="https://example.com", text="here")["text"] == "here"
    assert doc.texts() == ["See here"]


def test_an_empty_paragraph_has_nothing_to_link(make_doc):
    make_doc("")
    assert "has no text to make a link of" in link(1, address="https://example.com")["error"]


def test_text_that_is_not_there_is_refused(make_doc):
    make_doc("One.")
    assert "'nine' is not in paragraph 1" in link(1, address="x", match="nine")["error"]


def test_a_link_to_a_bookmark_needs_the_bookmark(make_doc):
    doc = make_doc("One.", "Two.")
    assert "no bookmark named 'Nope'" in link(1, bookmark="Nope", match="One")["error"]
    assert doc.Hyperlinks.Count == 0
    mark("Target", 2)
    out = link(1, bookmark="Target", match="One")
    assert out["bookmark"] == "Target" and out["address"] is None


def test_links_are_numbered_by_place_and_a_new_one_says_its_number(make_doc):
    make_doc("One two.", "Three.")
    assert link(2, address="b")["number"] == 1
    assert link(1, address="a", match="One")["number"] == 1
    out = answer(links._list_hyperlinks_impl, 1, 100)
    assert [(h["number"], h["index"], h["address"]) for h in out["hyperlinks"]] == [
        (1, 1, "a"), (2, 2, "b")]


def test_a_long_list_says_where_to_go_on(make_doc):
    make_doc("a b c")
    for t in "abc":
        link(1, address=t, match=t)
    out = answer(links._list_hyperlinks_impl, 1, 2)
    assert len(out["hyperlinks"]) == 2 and out["more"] == "Links 3 to 3 follow; pass start=3."


def test_a_links_target_text_and_tip_change(make_doc):
    make_doc("One.", "Two.")
    link(1, address="old", match="One", tip="tip")
    mark("Target", 2)
    out = answer(links._set_hyperlink_impl, 1, None, "Target", "Eins", "")
    assert (out["address"], out["bookmark"], out["text"], out["screen_tip"]) == (
        None, "Target", "Eins", None)
    out = answer(links._set_hyperlink_impl, 1, "new", None, None, None)
    assert (out["address"], out["bookmark"]) == ("new", None)


def test_an_address_link_is_remade_to_go_to_a_bookmark(make_doc):
    doc = make_doc("One two.", "Three.")
    link(1, address="old", match="two", tip="tip")
    mark("Target", 2)
    out = answer(links._set_hyperlink_impl, 1, None, "Target", None, None)
    assert (out["number"], out["text"], out["address"], out["bookmark"], out["screen_tip"]) == (
        1, "two", None, "Target", "tip")
    assert doc.Hyperlinks.Count == 1 and doc.texts() == ["One two.", "Three."]


def test_removing_a_link_keeps_its_text(make_doc):
    doc = make_doc("One two.")
    link(1, address="a", match="two")
    out = answer(links._remove_hyperlink_impl, 1)
    assert out == {"success": True, "removed": 1, "text_kept": "two", "left": 0}
    assert doc.texts() == ["One two."]


@pytest.mark.parametrize("impl, args", [(links._remove_hyperlink_impl, (2,)),
                                        (links._set_hyperlink_impl, (2, "a", None, None, None))])
def test_a_link_that_is_not_there_is_refused(make_doc, impl, args):
    make_doc("One.")
    assert "There is no hyperlink 2: the document has 0" in answer(impl, *args)["error"]


def test_a_bookmark_over_some_text_and_over_a_paragraph(make_doc):
    make_doc("One two.", "Three.")
    assert mark("Part", 1, "two") == {"success": True, "name": "Part", "index": 1, "text": "two",
                                      "empty": False}
    assert mark("Whole", 2)["text"] == "Three."


def test_a_name_already_used_moves_the_bookmark_and_says_so(make_doc):
    make_doc("One.", "Two.")
    assert "note" not in mark("Spot", 1)
    out = mark("Spot", 2)
    assert out["index"] == 2 and "moved" in out["note"]
    assert len(answer(links._list_bookmarks_impl)["bookmarks"]) == 1


def test_bookmarks_are_listed_by_place_not_by_name(make_doc):
    make_doc("One.", "Two.")
    mark("Zed", 1)
    mark("Abc", 2)
    assert [b["name"] for b in answer(links._list_bookmarks_impl)["bookmarks"]] == ["Zed", "Abc"]


def test_words_hidden_bookmarks_are_left_out_even_when_word_shows_them(make_doc):
    doc = make_doc("One.")
    mark("Spot", 1)
    mark("Other", 1)
    doc.Bookmarks.Add("_Ref123", doc.Range(0, 3))
    assert {b["name"] for b in answer(links._list_bookmarks_impl)["bookmarks"]} == {"Other", "Spot"}
    assert answer(links._delete_bookmark_impl, "Spot")["left"] == 1


def test_deleting_a_bookmark(make_doc):
    make_doc("One.")
    mark("Spot", 1)
    assert answer(links._delete_bookmark_impl, "Spot") == {
        "success": True, "deleted": "Spot", "left": 0}
    assert "no bookmark named 'Spot'" in answer(links._delete_bookmark_impl, "Spot")["error"]


@pytest.mark.parametrize("name", ["見出し_1", "Part2", "a" * 40])
def test_names_word_takes(name):
    assert links.AddBookmarkInput(name=name, index=1).name == name


@pytest.mark.parametrize("name", ["two words", "1st", "_hidden", "a" * 41, "a-b", ""])
def test_names_word_refuses(name):
    with pytest.raises(ValidationError):
        links.AddBookmarkInput(name=name, index=1)


def test_a_link_goes_to_one_place():
    with pytest.raises(ValidationError):
        links.InsertHyperlinkInput(index=1)
    with pytest.raises(ValidationError):
        links.InsertHyperlinkInput(index=1, address="a", bookmark="b")
    with pytest.raises(ValidationError):
        links.InsertHyperlinkInput(index=1, address="a", match="m", text="t")


def test_a_change_changes_something():
    with pytest.raises(ValidationError):
        links.SetHyperlinkInput(number=1)
    assert links.SetHyperlinkInput(number=1, screen_tip="").screen_tip == ""
