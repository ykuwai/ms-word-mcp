"""Section breaks and headers of several paragraphs."""

from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import mac, page


# -- shared ----------------------------------------------------------------------------

def test_a_header_is_given_as_text_or_as_paragraphs():
    with pytest.raises(ValidationError, match="one of the two"):
        page.SetHeaderFooterInput(part="header")
    with pytest.raises(ValidationError, match="one of the two"):
        page.SetHeaderFooterInput(part="header", text="x", paragraphs=[{"text": "y"}])
    with pytest.raises(ValidationError, match="its own alignment"):
        page.SetHeaderFooterInput(part="header", paragraphs=[{"text": "y"}], alignment="left")
    with pytest.raises(ValidationError, match="take the line breaks out"):
        page.SetHeaderFooterInput(part="header", paragraphs=[{"text": "a\nb"}])
    assert page.SetHeaderFooterInput(part="footer", text="").text == ""


def test_paragraphs_are_written_apart_with_their_alignments():
    whole, aligns = page.header_text(None, [{"text": "Example Ltd", "alignment": "left"},
                                            {"text": "page {page}", "alignment": None}])
    assert (whole, aligns) == ("Example Ltd\rpage {page}", ["left", None])
    plain, fields = page.parse_placeholders(whole)
    assert plain == "Example Ltd\rpage #" and fields == [(17, "PAGE")]
    assert page.header_text("one", None) == ("one", [])


def test_a_section_break_says_where_the_paragraphs_went():
    out = page.section_break_answer("next_page", 1, 2, 2, 4, at_end=False)
    assert (out["section"], out["starts_at"], out["paragraphs"]) == (2, 3, 4)
    assert "is now 3" in out["note"]
    out = page.section_break_answer("continuous", 4, 3, 3, 5, at_end=True)
    assert out["starts_at"] == 5 and "empty paragraph, 5" in out["note"]


# -- Windows, against a stand-in document ----------------------------------------------

class _Range:
    def __init__(self, doc, start, end):
        self.doc, self.Start, self.End = doc, start, end

    def InsertBreak(self, kind):
        self.doc.breaks.append((self.Start, kind))
        self.doc.text = self.doc.text[:self.Start] + "\x0c" + self.doc.text[self.Start:]

    @property
    def Sections(self):
        before = self.doc.text[:self.Start].count("\x0c")
        return SimpleNamespace(Item=lambda i: SimpleNamespace(Index=before + 1))


class _Doc:
    """Paragraphs end with \\r; a section break is a \\x0c of its own."""

    def __init__(self, *paragraphs):
        self.text = "".join(p + "\r" for p in paragraphs)
        self.breaks = []
        self.Windows = SimpleNamespace(Count=0)

    def _paragraphs(self):
        out, start = [], 0
        for i, ch in enumerate(self.text):
            if ch in "\r\x0c":
                out.append((start, i + 1))
                start = i + 1
        return out

    @property
    def Paragraphs(self):
        spans = self._paragraphs()
        return SimpleNamespace(Count=len(spans),
                               Item=lambda i: SimpleNamespace(Range=_Range(self, *spans[i - 1])))

    @property
    def Sections(self):
        return SimpleNamespace(Count=self.text.count("\x0c") + 1)

    def Range(self, start, end):
        return _Range(self, start, end)


@pytest.fixture
def com_doc(monkeypatch):
    def make(*paragraphs):
        doc = _Doc(*paragraphs)
        monkeypatch.setattr(page.backend, "word", SimpleNamespace(document=lambda: doc))
        return doc
    return make


def test_a_break_goes_at_the_start_of_the_next_paragraph(com_doc):
    doc = com_doc("One", "Two", "Three")
    out = page._insert_section_break_impl(1, "next_page")
    assert doc.breaks == [(4, 2)]  # wdSectionBreakNextPage, where "Two" began
    assert doc.text == "One\r\x0cTwo\rThree\r"
    assert (out["section"], out["sections"], out["starts_at"], out["paragraphs"]) == (2, 2, 3, 4)


def test_after_the_last_paragraph_the_break_ends_its_text(com_doc):
    doc = com_doc("One", "Two")
    out = page._insert_section_break_impl(2, "continuous")
    assert doc.breaks == [(7, 3)]  # before the last paragraph's mark
    assert doc.text == "One\rTwo\x0c\r"
    assert (out["section"], out["starts_at"], out["paragraphs"]) == (2, 3, 3)
    assert "empty paragraph" in out["note"]


def test_a_paragraph_that_is_not_there_is_refused(com_doc):
    com_doc("One")
    with pytest.raises(ValueError, match="There is no paragraph 5"):
        page._insert_section_break_impl(5, "next_page")


# -- Mac, against stand-ins for Word for Mac's scripting ------------------------------

def _keys():
    return SimpleNamespace(paragraph="paragraph", section="section",
                           section_break_next_page="next page", section_break_continuous="continuous")


class _MacDoc(_Doc):
    def count(self, each):
        return len(self._paragraphs()) if each == "paragraph" else self.Sections.Count

    @property
    def paragraphs(self):
        spans = self._paragraphs()

        class Paragraphs:
            def __getitem__(self, i):
                a, b = spans[i - 1]
                return SimpleNamespace(text_object=SimpleNamespace(
                    start_of_content=SimpleNamespace(get=lambda: a),
                    end_of_content=SimpleNamespace(get=lambda: b)))
        return Paragraphs()

    @property
    def sections(self):
        doc = self

        class Sections:
            def __getitem__(self, n):
                starts = [0] + [i + 1 for i, ch in enumerate(doc.text) if ch == "\x0c"]
                return SimpleNamespace(text_object=SimpleNamespace(
                    start_of_content=SimpleNamespace(get=lambda: starts[n - 1])))
        return Sections()

    def create_range(self, start, end_):
        return SimpleNamespace(start=start, end=end_)


def test_on_the_mac_a_break_is_made_in_the_document_named(monkeypatch):
    doc = _MacDoc("One", "Two", "Three")
    asked = []

    def insert_break(at, break_type):
        asked.append((at.start, break_type))
        doc.text = doc.text[:at.start] + "\x0c" + doc.text[at.start:]

    monkeypatch.setattr(mac, "_k", _keys)
    monkeypatch.setattr(page.backend, "word", SimpleNamespace(
        document=lambda: doc, get_app=lambda: SimpleNamespace(insert_break=insert_break)))
    out = page._mac_insert_section_break(1, "next_page")
    assert asked == [(4, "next page")]
    assert (out["section"], out["sections"], out["starts_at"], out["paragraphs"]) == (2, 2, 3, 4)
    out = page._mac_insert_section_break(4, "continuous")
    assert asked[-1] == (14, "continuous")  # before the mark of "Three"
    assert (out["section"], out["sections"], out["starts_at"]) == (3, 3, 5)


def test_on_the_mac_a_header_aligned_apart_reads_as_no_one_alignment(monkeypatch):
    """As Windows answers wdUndefined, which reads back as None."""
    monkeypatch.setattr(mac, "_from_k", lambda table, value: value)
    hf = SimpleNamespace(
        text_object=SimpleNamespace(content=SimpleNamespace(get=lambda: "a\rb\r"),
                                    paragraph_format=SimpleNamespace(
                                        alignment=SimpleNamespace(get=lambda: "missing_value"))),
        link_to_previous=SimpleNamespace(get=lambda: False))
    assert page._mac_hf_info(hf) == {"text": "a\rb", "alignment": None, "linked_to_previous": False}
    assert page.ALIGNMENT_NAMES.get(9999999) is None  # Windows, for the same header
