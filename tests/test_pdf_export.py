"""word_save_document_as to a PDF: some pages only, and with or without markup."""

import os
import sys
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ms_word_mcp.backend import com
from ms_word_mcp.tools import documents, mac, screen


# -- shared ----------------------------------------------------------------------------

@pytest.mark.parametrize("args, says", [
    ({"path": "/d/a.docx", "from_page": 1}, "for a .pdf only"),
    ({"path": "/d/a.docx", "markup": True}, "for a .pdf only"),
    ({"path": "/d/a.pdf", "to_page": 2}, "Give `from_page` with `to_page`"),
    ({"path": "/d/a.pdf", "from_page": 3, "to_page": 2}, "comes before"),
])
def test_pdf_options_that_cannot_be_carried_out_are_refused(args, says):
    with pytest.raises(ValidationError, match=says):
        documents.SaveDocumentAsInput(**args)


def test_the_pages_are_checked_against_the_document():
    assert documents.check_pages(None, None, 5) is None
    assert documents.check_pages(2, None, 5) == (2, 2)
    assert documents.check_pages(2, 4, 5) == (2, 4)
    with pytest.raises(ValueError, match="There is no page 6: the document has 5 pages"):
        documents.check_pages(2, 6, 5)


# -- Windows, against a stand-in document ----------------------------------------------

def _window(caption):
    return SimpleNamespace(Count=1, Item=lambda i: SimpleNamespace(Caption=caption))


@pytest.fixture
def com_doc(monkeypatch, tmp_path):
    exported = []

    def export(**kwargs):
        exported.append(kwargs)
        open(kwargs["OutputFileName"], "wb").write(b"%PDF")

    doc = SimpleNamespace(Name="Report.docx", FullName="C:\\d\\Report.docx", ExportAsFixedFormat=export,
                          ComputeStatistics=lambda kind: 5, Windows=_window("Document1"))
    monkeypatch.setattr(documents.backend, "word", SimpleNamespace(document=lambda: doc))
    return exported, str(tmp_path / "out.pdf")


def test_on_windows_a_pdf_is_the_document_as_it_will_print(com_doc):
    exported, path = com_doc
    out = documents._save_document_as_impl(path, False)
    assert exported[0]["Item"] == documents.WD_EXPORT_CONTENT and "Range" not in exported[0]
    assert out["markup"] is False and "pages" not in out


def test_on_windows_some_pages_with_their_markup(com_doc):
    exported, path = com_doc
    out = documents._save_document_as_impl(path, False, 2, 3, True)
    assert {k: exported[0][k] for k in ("Item", "Range", "From", "To")} == {
        "Item": documents.WD_EXPORT_WITH_MARKUP, "Range": documents.WD_EXPORT_FROM_TO, "From": 2, "To": 3}
    assert out["pages"] == [2, 3] and out["markup"] is True


def test_on_windows_a_page_that_is_not_there_is_refused_before_exporting(com_doc):
    exported, path = com_doc
    with pytest.raises(ValueError, match="There is no page 9"):
        documents._save_document_as_impl(path, False, 9, None, False)
    assert exported == [] and not os.path.exists(path)


def test_on_windows_the_answer_names_the_document_word_has_just_lost_the_name_of(com_doc, monkeypatch):
    exported, path = com_doc
    doc = documents.backend.word.document()
    real = doc.ExportAsFixedFormat

    def export(**kwargs):
        real(**kwargs)
        doc.Name = doc.FullName = ""  # what Word does to a document never saved
    doc.Name, doc.FullName = "Document1", "Document1"
    doc.ExportAsFixedFormat = export
    assert documents._save_document_as_impl(path, True)["document"] == "Document1"
    doc.Name = ""  # Word has lost both now; the window's caption still says it
    assert documents._save_document_as_impl(path, True)["document"] == "Document1"


def test_a_document_word_answers_no_name_for_goes_by_its_windows_caption():
    lost = SimpleNamespace(Name="", FullName="", Windows=_window("Document1"))
    assert com.names_of(lost) == ("Document1", "Document1")
    saved = SimpleNamespace(Name="a.docx", FullName="C:\\d\\a.docx", Windows=_window("a.docx"))
    assert com.names_of(saved) == ("a.docx", "C:\\d\\a.docx")
    windowless = SimpleNamespace(Name="", FullName="", Windows=SimpleNamespace(Count=0))
    assert com.names_of(windowless) == ("", "")


# -- Mac, against stand-ins ------------------------------------------------------------

def _mac(monkeypatch, comments=0, revisions=0, pages=5):
    log = []

    def saver(label):
        def save_as(file_name, file_format):
            log.append((label, os.path.basename(file_name)))
            open(file_name, "wb").write(b"%PDF")
        return save_as

    doc = SimpleNamespace(count=lambda each: {"Word comment": comments, "revision": revisions}[each],
                          save_as=saver("original"))
    copy = SimpleNamespace(save_as=saver("copy"))
    monkeypatch.setattr(mac, "_k", lambda: SimpleNamespace(Word_comment="Word comment",
                                                           revision="revision", format_PDF="pdf"))
    monkeypatch.setattr(screen, "_mac_active_name", lambda app: "Mine.docx")
    monkeypatch.setattr(screen, "_mac_without_markup", lambda app, d: log.append(("copy made",)) or copy)
    monkeypatch.setattr(screen, "_mac_close_copy", lambda app, c, active: log.append(("copy closed", active)))
    monkeypatch.setattr(mac, "pdf_page_count", lambda p: pages)
    monkeypatch.setattr(mac, "keep_pdf_pages", lambda s, t, kept: log.append(("kept", kept))
                        or open(t, "wb").write(b"%PDF"))
    return doc, log


def test_on_the_mac_a_document_without_markup_is_exported_as_it_is(monkeypatch, tmp_path):
    doc, log = _mac(monkeypatch)
    assert mac._mac_save_pdf(None, doc, str(tmp_path / "a.pdf"), None, None, False) is None
    assert log == [("original", "a.pdf")]


def test_on_the_mac_comments_are_left_out_through_a_copy(monkeypatch, tmp_path):
    doc, log = _mac(monkeypatch, comments=2)
    mac._mac_save_pdf(None, doc, str(tmp_path / "a.pdf"), None, None, False)
    assert log == [("copy made",), ("copy", "a.pdf"), ("copy closed", "Mine.docx")]


def test_on_the_mac_markup_asked_for_is_drawn_by_word(monkeypatch, tmp_path):
    doc, log = _mac(monkeypatch, revisions=1)
    mac._mac_save_pdf(None, doc, str(tmp_path / "a.pdf"), None, None, True)
    assert log == [("original", "a.pdf")]


def test_on_the_mac_a_range_is_taken_out_of_the_whole_pdf(monkeypatch, tmp_path):
    doc, log = _mac(monkeypatch)
    target = tmp_path / "a.pdf"
    assert mac._mac_save_pdf(None, doc, str(target), 2, 3, False) == (2, 3)
    assert log == [("original", "a.pdf.all.pdf"), ("kept", (2, 3))]
    assert target.exists() and not (tmp_path / "a.pdf.all.pdf").exists()


def test_on_the_mac_a_page_past_the_end_is_refused_and_nothing_left(monkeypatch, tmp_path):
    doc, _ = _mac(monkeypatch, pages=2)
    with pytest.raises(ValueError, match="There is no page 4"):
        mac._mac_save_pdf(None, doc, str(tmp_path / "a.pdf"), 3, 4, False)
    assert list(tmp_path.iterdir()) == []


@pytest.mark.skipif(sys.platform != "darwin", reason="Quartz")
def test_on_the_mac_pages_are_copied_out_of_a_pdf(tmp_path):
    import Quartz
    from Foundation import NSURL

    whole = tmp_path / "whole.pdf"
    context = Quartz.CGPDFContextCreateWithURL(NSURL.fileURLWithPath_(str(whole)), None, None)
    for _ in range(4):
        Quartz.CGContextBeginPage(context, Quartz.CGRectMake(0, 0, 612, 792))
        Quartz.CGContextEndPage(context)
    Quartz.CGPDFContextClose(context)
    part = tmp_path / "part.pdf"
    mac.keep_pdf_pages(str(whole), str(part), (2, 3))
    assert mac.pdf_page_count(str(part)) == 2


def test_a_document_that_lost_its_name_is_still_named_in_answers():
    lost = SimpleNamespace(Name="", FullName="", Saved=False, ReadOnly=False, Windows=_window("Document1"))
    summary = documents._doc_summary(lost)
    assert (summary["name"], summary["full_name"]) == ("Document1", "Document1")
