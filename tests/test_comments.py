import json
from datetime import datetime
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _common, comments
from ms_word_mcp.tools import mac


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def add(index, text, on=None):
    return answer(comments._add_comment_impl, index, text, on)


def listed(start=1, limit=100):
    return answer(comments._list_comments_impl, start, limit)


# -- Windows, against the fake Word ----------------------------------------------------

def test_a_comment_is_about_the_paragraphs_text_without_its_mark(make_doc):
    doc = make_doc("One two.", "Three.")
    out = add(1, "Check this")
    assert out == {"success": True, "number": 1, "index": 1, "on": "One two.", "text": "Check this",
                   "author": "Office User", "initials": "OU", "date": "2026-11-20T09:30:00",
                   "done": False, "count": 1}
    assert (doc.Comments.Item(1).start, doc.Comments.Item(1).end) == (0, 8)
    assert doc.texts() == ["One two.", "Three."]  # the text is left as it was


def test_a_comment_can_be_about_some_text_in_the_paragraph(make_doc):
    make_doc("One two three.")
    out = add(1, "Which two?", on="two")
    assert (out["on"], out["index"]) == ("two", 1)


def test_text_that_is_not_there_is_refused_and_nothing_is_added(make_doc):
    doc = make_doc("One two.")
    out = add(1, "A note", on="nine")
    assert "'nine' is not in paragraph 1" in out["error"]
    assert doc.Comments.Count == 0


def test_a_paragraph_that_is_not_there_is_refused(make_doc):
    make_doc("One.")
    assert "There is no paragraph 4" in add(4, "x")["error"]


def test_comments_are_numbered_in_the_order_of_their_text(make_doc):
    make_doc("One.", "Two.")
    assert add(2, "later")["number"] == 1
    out = add(1, "earlier")
    assert (out["number"], out["count"]) == (1, 2)
    assert [(c["number"], c["index"], c["text"]) for c in listed()["comments"]] == [
        (1, 1, "earlier"), (2, 2, "later")]


def test_a_long_list_says_where_to_go_on(make_doc):
    make_doc("One.", "Two.", "Three.")
    for i in (1, 2, 3):
        add(i, f"c{i}")
    out = listed(limit=2)
    assert [c["text"] for c in out["comments"]] == ["c1", "c2"]
    assert out["more"] == "Comments 3 to 3 follow; pass start=3."
    rest = listed(start=3)
    assert [c["text"] for c in rest["comments"]] == ["c3"] and "more" not in rest


def test_a_reply_is_listed_after_its_comment_and_says_which_it_answers(make_doc):
    make_doc("One.", "Two.")
    add(1, "first")
    add(2, "second")
    out = answer(comments._reply_to_comment_impl, 1, "agreed")
    assert (out["number"], out["reply_to"], out["text"], out["count"]) == (2, 1, "agreed", 3)
    assert [(c["text"], c.get("reply_to")) for c in listed()["comments"]] == [
        ("first", None), ("agreed", 1), ("second", None)]


def test_a_reply_to_a_reply_goes_to_its_comment(make_doc):
    make_doc("One.")
    add(1, "first")
    answer(comments._reply_to_comment_impl, 1, "agreed")
    out = answer(comments._reply_to_comment_impl, 2, "me too")
    assert out["reply_to"] == 1


def test_a_comment_is_marked_done_and_opened_again(make_doc):
    doc = make_doc("One.")
    add(1, "first")
    assert answer(comments._resolve_comment_impl, 1, True)["done"] is True
    assert doc.Comments.Item(1).Done is True
    assert answer(comments._resolve_comment_impl, 1, False)["done"] is False


def test_a_number_that_is_not_there_is_refused(make_doc):
    make_doc("One.")
    add(1, "first")
    for impl, args in ((comments._reply_to_comment_impl, (3, "x")),
                       (comments._resolve_comment_impl, (3, True)),
                       (comments._delete_comments_impl, ([3], False))):
        assert "There is no comment 3: the document has 1" in answer(impl, *args)["error"]


def test_comments_are_deleted_from_the_last_so_each_number_holds(make_doc):
    doc = make_doc("One.", "Two.", "Three.")
    for i in (1, 2, 3):
        add(i, f"c{i}")
    out = answer(comments._delete_comments_impl, [1, 3], False)
    assert (out["deleted"], out["left"]) == (2, 1)
    assert [c.text for c in doc.Comments.ordered()] == ["c2"]


def test_deleting_a_comment_takes_its_replies_and_says_so(make_doc):
    doc = make_doc("One.", "Two.")
    add(1, "first")
    add(2, "second")
    answer(comments._reply_to_comment_impl, 1, "agreed")
    out = answer(comments._delete_comments_impl, [1], False)
    assert (out["deleted"], out["left"]) == (2, 1)
    assert "deletes its replies too" in out["note"]
    assert [c.text for c in doc.Comments.ordered()] == ["second"]


def test_every_comment_goes(make_doc):
    doc = make_doc("One.", "Two.")
    add(1, "first")
    add(2, "second")
    answer(comments._reply_to_comment_impl, 1, "agreed")
    out = answer(comments._delete_comments_impl, None, True)
    assert (out["deleted"], out["left"]) == (3, 0) and "note" not in out
    assert doc.Comments.Count == 0


@pytest.mark.parametrize("model, args", [
    (comments.AddCommentInput, {"index": 1, "text": "two\nlines"}),
    (comments.ReplyInput, {"number": 1, "text": "two\rlines"}),
    (comments.DeleteCommentsInput, {}),
    (comments.DeleteCommentsInput, {"numbers": [1], "every": True}),
    (comments.DeleteCommentsInput, {"numbers": [0]}),
])
def test_arguments_that_cannot_be_carried_out_are_refused(model, args):
    with pytest.raises(ValidationError):
        model(**args)


def test_a_long_text_a_comment_is_about_is_cut_short():
    assert comments.shown_on("x" * 300) == {"on": "x" * comments.ON_SHOWN, "on_cut": True}


# -- Mac, against stand-ins for Word for Mac's scripting ------------------------------

def _keys():
    return SimpleNamespace(Word_comment="Word comment", paragraph="paragraph",
                           scope="scope", comment_text="comment text")


class _Value:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value


class _MacComment:
    def __init__(self, doc, start, end, text):
        self.doc, self.start, self.end, self.text = doc, start, end, text

    @property
    def scope(self):
        return SimpleNamespace(start_of_content=_Value(self.start),
                               content=_Value(self.doc.text[self.start:self.end]))

    @property
    def comment_text(self):
        return SimpleNamespace(content=_Value(self.text))

    @property
    def comment_index(self):
        return _Value(self.doc.ordered().index(self) + 1)

    author = _Value("Office User")
    initials = _Value("OU")
    date_value = _Value(datetime(2026, 11, 20, 9, 30))

    def delete(self):
        self.doc.deleted.append(self.doc.ordered().index(self) + 1)
        self.doc.items.remove(self)


class _MacDoc:
    """Paragraphs as text, and comments made through `make` (which is noted)."""

    def __init__(self, *paragraphs):
        self.text = "".join(p + "\r" for p in paragraphs)
        self.items, self.made, self.deleted = [], [], []

    def ordered(self):
        return sorted(self.items, key=lambda c: c.start)

    def _starts(self):
        starts, at = [], 0
        for p in self.text.split("\r")[:-1]:
            starts.append((at, at + len(p) + 1))
            at += len(p) + 1
        return starts

    def count(self, each):
        return len(self._starts()) if each == "paragraph" else len(self.items)

    @property
    def paragraphs(self):
        doc = self

        class Paragraphs:
            def __getitem__(self, i):
                a, b = doc._starts()[i - 1]
                return SimpleNamespace(text_object=SimpleNamespace(
                    start_of_content=_Value(a), end_of_content=_Value(b), content=_Value(doc.text[a:b])))
        return Paragraphs()

    def create_range(self, start, end_):
        doc = self
        return SimpleNamespace(start=start, end=end_, content=_Value(self.text[start:end_]),
                               count=lambda each: doc.text[:end_].count("\r") + (0 if doc.text[end_ - 1:end_] == "\r" else 1))

    @property
    def Word_comments(self):
        doc = self

        class Comments:
            def __getitem__(self, i):
                return doc.ordered()[i - 1]
        return Comments()

    def make(self, new, at, with_properties):
        self.made.append((new, at is self, dict(with_properties)))
        scope = with_properties["scope"]
        comment = _MacComment(self, scope.start, scope.end, with_properties.get("comment text"))
        self.items.append(comment)
        return comment


@pytest.fixture
def mac_doc(monkeypatch):
    monkeypatch.setattr(mac, "_k", _keys)

    def make(*paragraphs):
        doc = _MacDoc(*paragraphs)
        monkeypatch.setattr(comments.backend, "word", SimpleNamespace(document=lambda: doc))
        return doc
    return make


def test_on_the_mac_a_comment_is_made_with_its_text_as_well_as_its_scope(mac_doc):
    """Made with the scope alone, Word for Mac ended its process."""
    doc = mac_doc("One two three.", "Four.")
    out = comments._mac_add_comment(1, "Which two?", "two")
    [(new, at_document, props)] = doc.made
    assert (new, at_document) == ("Word comment", True)
    assert props["comment text"] == "Which two?"
    assert (props["scope"].start, props["scope"].end) == (4, 7)
    assert out == {"success": True, "number": 1, "index": 1, "on": "two", "text": "Which two?",
                   "author": "Office User", "initials": "OU", "date": "2026-11-20T09:30:00", "count": 1}


def test_on_the_mac_a_whole_paragraph_leaves_out_its_mark(mac_doc):
    doc = mac_doc("One.", "Four five.")
    out = comments._mac_add_comment(2, "x", None)
    scope = doc.made[0][2]["scope"]
    assert (scope.start, scope.end, out["on"], out["index"]) == (5, 15, "Four five.", 2)


def test_on_the_mac_text_that_is_not_there_is_refused_before_anything_is_made(mac_doc):
    doc = mac_doc("One two.")
    with pytest.raises(ValueError, match="'nine' is not in paragraph 1"):
        comments._mac_add_comment(1, "x", "nine")
    assert doc.made == []


def test_on_the_mac_comments_are_listed_in_order(mac_doc):
    mac_doc("One.", "Two.")
    comments._mac_add_comment(2, "later", None)
    comments._mac_add_comment(1, "earlier", None)
    out = comments._mac_list_comments(1, 100)
    assert [(c["number"], c["index"], c["text"]) for c in out["comments"]] == [
        (1, 1, "earlier"), (2, 2, "later")]
    assert out["count"] == 2


def test_on_the_mac_comments_are_deleted_one_at_a_time_from_the_last(mac_doc):
    doc = mac_doc("One.", "Two.", "Three.")
    for i in (1, 2, 3):
        comments._mac_add_comment(i, f"c{i}", None)
    out = comments._mac_delete_comments([1, 3], False)
    assert doc.deleted == [3, 1] and (out["deleted"], out["left"]) == (2, 1)
    out = comments._mac_delete_comments(None, True)
    assert (out["deleted"], out["left"]) == (1, 0)


@pytest.mark.parametrize("impl, args, says", [
    (comments._mac_reply_to_comment, (1, "x"), "no replies"),
    (comments._mac_resolve_comment, (1, True), "no done state"),
])
def test_on_the_mac_without_the_bridge_replies_and_done_say_why_they_cannot(mac_doc, impl, args, says):
    mac_doc("One.")
    comments._mac_add_comment(1, "c", None)
    with pytest.raises(ValueError, match=f"{says}.*cannot be used with this Word \\(this server has no bridge\\)"):
        impl(*args)


class _Vba:
    """The bridge, answering for a document whose comment 2 replies to comment 1."""

    def __init__(self):
        self.asked = []

    def call(self, app, what, *fields):
        self.asked.append((what, *fields[1:]))
        if what == "CommentState":
            return ["0,-1,0", "1,0,5", "0,-1,9"]
        if what == "Reply":
            return ["2", "3"]
        if what == "SetDone":
            return ["1"]
        if what == "DeleteComments":
            return ["1", "1"]  # comment 1 and its reply went, one number was asked
        raise AssertionError(what)


@pytest.fixture
def mac_doc_with_bridge(mac_doc, monkeypatch):
    def make(*paragraphs):
        doc = mac_doc(*paragraphs)
        vba = _Vba()
        monkeypatch.setattr(comments.backend, "word", SimpleNamespace(
            document=lambda: doc, vba=vba, get_app=lambda: None))
        monkeypatch.setattr(mac, "_full", lambda d: "/d/Report.docx")
        for i in range(1, len(paragraphs) + 1):
            comments._mac_add_comment(i, f"c{i}", None)
        return doc, vba
    return make


def test_on_the_mac_the_bridge_gives_the_done_state_and_what_a_reply_answers(mac_doc_with_bridge):
    _, vba = mac_doc_with_bridge("One.", "Two.", "Three.")
    out = comments._mac_list_comments(1, 10)["comments"]
    assert [(c["number"], c["done"], c.get("reply_to")) for c in out] == [
        (1, False, None), (2, True, 1), (3, False, None)]
    assert vba.asked[-1] == ("CommentState",)


def test_on_the_mac_a_reply_goes_through_the_bridge(mac_doc_with_bridge):
    _, vba = mac_doc_with_bridge("One.", "Two.", "Three.")
    out = comments._mac_reply_to_comment(1, "Agreed")
    assert ("Reply", "1", "65,103,114,101,101,100") in vba.asked
    assert (out["number"], out["reply_to"], out["count"]) == (2, 1, 3)


def test_on_the_mac_done_goes_through_the_bridge(mac_doc_with_bridge):
    _, vba = mac_doc_with_bridge("One.", "Two.", "Three.")
    out = comments._mac_resolve_comment(2, True)
    assert ("SetDone", "2", "1") in vba.asked and out["done"] is True


def test_on_the_mac_a_comment_that_is_not_there_is_refused_before_the_bridge(mac_doc_with_bridge):
    _, vba = mac_doc_with_bridge("One.")
    with pytest.raises(ValueError, match="There is no comment 4"):
        comments._mac_resolve_comment(4, True)
    assert not any(a[0] == "SetDone" for a in vba.asked)


def test_on_the_mac_a_comment_and_its_replies_are_deleted_through_the_bridge(mac_doc_with_bridge):
    """Word for Mac's own `delete` leaves a comment that has replies."""
    doc, vba = mac_doc_with_bridge("One.", "Two.", "Three.")
    out = comments._mac_delete_comments([1], False)
    assert ("DeleteComments", "1") in vba.asked and doc.deleted == []
    assert (out["deleted"], out["left"]) == (2, 1) and "replies" in out["note"]
