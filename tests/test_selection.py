"""word_get_selection: what the user has selected, on both platforms."""

import sys
from types import SimpleNamespace

import pytest

from ms_word_mcp.tools import mac, selection


# -- shared ----------------------------------------------------------------------------

def test_a_selection_in_the_text_says_which_paragraphs():
    out = selection.selection_answer("Report.docx", "/d/Report.docx", "main", 10, 40,
                                     "Second paragraph.\r", (2, 2), "/d/Report.docx")
    assert out == {"success": True, "document": "Report.docx", "full_name": "/d/Report.docx",
                   "story": "main", "empty": False, "text": "Second paragraph.", "start": 2, "end": 2}


def test_the_cursor_alone_says_where_it_is():
    # Word for Mac answers the character after the cursor as its text.
    out = selection.selection_answer("R", "/d/R", "main", 12, 12, "第", (3, 3), "/d/R")
    assert out["empty"] is True and out["text"] == ""
    assert out["note"] == "Nothing is selected: the cursor is in paragraph 3."


def test_a_selection_in_a_header_has_no_paragraph_numbers():
    out = selection.selection_answer("R", "/d/R", "header", 0, 6, "Report", None, "/d/R")
    assert "start" not in out and "in the header" in out["note"]


def test_a_selection_in_another_document_says_how_to_reach_it():
    out = selection.selection_answer("Other.docx", "/d/Other.docx", "main", 0, 3, "abc", (1, 1),
                                     "/d/Report.docx")
    assert "pass document='/d/Other.docx'" in out["note"]


def test_a_long_selection_is_cut_short():
    out = selection.selection_answer("R", "/d/R", "main", 0, 5000, "x" * 5000, (1, 9), None)
    assert len(out["text"]) == selection.TEXT_SHOWN and out["text_cut"] is True


def test_a_selection_ending_where_a_paragraph_begins_does_not_reach_into_it():
    # Paragraphs start at 0, 10 and 20; counting up to a position.
    count_to = lambda at: 1 + (at > 10) + (at > 20)  # noqa: E731
    assert selection._paragraphs(count_to, 0, 10) == (1, 1)  # all of paragraph 1, its mark too
    assert selection._paragraphs(count_to, 5, 15) == (1, 2)
    assert selection._paragraphs(count_to, 12, 12) == (2, 2)


@pytest.mark.parametrize("name, story", [
    ("main_text_story", "main"), ("primary_header_story", "header"), ("even_pages_footer_story", "footer"),
    ("footnotes_story", "footnotes"), ("comments_story", "comments"), ("text_frame_story", "text_box"),
    ("footnote_separator_story", "footnotes"),
])
def test_mac_story_names(name, story):
    assert selection.mac_story(name) == story


# -- Windows, against a stand-in app ---------------------------------------------------

def _com(monkeypatch, start, end, story=1, documents=1, target="C:\\d\\Report.docx"):
    text = "First.\rSecond paragraph.\rThird.\r"
    starts = [0, 7, 25]

    def rng(a, b):
        return SimpleNamespace(Paragraphs=SimpleNamespace(Count=sum(1 for s in starts if s < b) or 1))
    doc = SimpleNamespace(Name="Report.docx", FullName="C:\\d\\Report.docx", Range=rng)
    sel = SimpleNamespace(Document=doc, StoryType=story,
                          Range=SimpleNamespace(Start=start, End=end, Text=text[start:end]))
    app = SimpleNamespace(Documents=SimpleNamespace(Count=documents), Selection=sel)
    monkeypatch.setattr(selection.backend, "word", SimpleNamespace(
        get_app=lambda: app, document=lambda: SimpleNamespace(FullName=target)))


def test_on_windows_the_selection_is_read(monkeypatch):
    _com(monkeypatch, 7, 25)
    out = selection._get_selection_impl()
    assert (out["start"], out["end"], out["text"], out["story"]) == (2, 2, "Second paragraph.", "main")
    assert "note" not in out


def test_on_windows_with_no_document_it_says_so(monkeypatch):
    _com(monkeypatch, 0, 0, documents=0)
    with pytest.raises(ValueError, match="No document is open"):
        selection._get_selection_impl()


def test_on_windows_a_header_is_named(monkeypatch):
    _com(monkeypatch, 0, 6, story=7)
    assert selection._get_selection_impl()["story"] == "header"


# -- Mac, against stand-ins ------------------------------------------------------------

class _V:
    def __init__(self, v):
        self.v = v

    def get(self):
        return self.v


def test_on_the_mac_the_selection_is_read(monkeypatch):
    starts = [0, 7, 25]
    doc = SimpleNamespace(create_range=lambda start, end_: SimpleNamespace(
        count=lambda each: sum(1 for s in starts if s < end_) or 1))
    sel = SimpleNamespace(document=SimpleNamespace(name=_V("Report.docx")), selection_start=_V(7),
                          selection_end=_V(25), story_type=_V("main_text_story"),
                          content=_V("Second paragraph.\r"))
    app = SimpleNamespace(count=lambda each: 1, selection=sel, documents={"Report.docx": doc})
    monkeypatch.setattr(mac, "_k", lambda: SimpleNamespace(document="document", paragraph="paragraph"))
    monkeypatch.setattr(mac, "_name", lambda v: v)
    monkeypatch.setattr(mac, "_summary", lambda d: {"name": "Report.docx", "full_name": "/d/Report.docx"})
    monkeypatch.setattr(selection.backend, "word", SimpleNamespace(get_app=lambda: app, document=lambda: doc))
    out = selection._mac_get_selection()
    assert (out["start"], out["end"], out["text"], out["story"]) == (2, 2, "Second paragraph.", "main")
    assert "note" not in out
