import json

import pytest

from ms_word_mcp.backend.documents import OpenDocument, full_name_key, pick_document
from ms_word_mcp.tools import _common, documents

DOCS = [
    OpenDocument(1, "report.docx", "C:\\work\\report.docx"),
    OpenDocument(2, "report.docx", "C:\\old\\report.docx"),
    OpenDocument(3, "notes.docx", "https://contoso.sharepoint.com/Shared%20Docs/notes.docx"),
]


def test_pick_by_full_path_name_and_stem():
    assert pick_document(DOCS, "c:/work/REPORT.docx").index == 1
    assert pick_document(DOCS, "notes.docx").index == 3
    assert pick_document(DOCS, "notes").index == 3
    assert pick_document(
        DOCS, "https://contoso.sharepoint.com/Shared Docs/notes.docx").index == 3


def test_pick_never_falls_back():
    with pytest.raises(ValueError, match="More than one"):
        pick_document(DOCS, "report.docx")
    with pytest.raises(ValueError, match="No open document"):
        pick_document(DOCS, "missing.docx")


def test_full_name_key_ignores_case_and_slashes():
    assert full_name_key("C:\\A\\b.docx") == full_name_key("c:/a/B.DOCX")


def test_paths_must_be_absolute():
    with pytest.raises(ValueError, match="full path"):
        documents.OpenDocumentInput(path="relative.docx")
    with pytest.raises(ValueError, match="not a format"):
        documents.SaveDocumentAsInput(path="/tmp/x.xyz")


def test_close_refuses_unsaved_changes_without_a_decision(word):
    word.doc.Saved = False
    out = json.loads(_common.run(documents._close_document_impl, None))
    assert "unsaved changes" in out["error"]


def test_save_refuses_a_document_without_a_file(word):
    word.doc.Path = ""
    out = json.loads(_common.run(documents._save_document_impl))
    assert "word_save_document_as" in out["error"]
