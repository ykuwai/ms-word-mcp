import json
import os

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



class _App:
    """Records the macro setting at the moment a file is opened."""

    def __init__(self, doc):
        self.AutomationSecurity = 1
        self.seen = []
        self.Documents = self
        self.doc = doc

    def Open(self, **kwargs):
        self.seen.append(self.AutomationSecurity)
        return self.doc

    def Add(self, **kwargs):
        self.seen.append(self.AutomationSecurity)
        return self.doc


def _with_app(word, doc):
    from types import SimpleNamespace

    doc.Paragraphs  # the fake document stands in for the opened one
    app = _App(doc)
    word.get_app = lambda allow_launch=False: app
    word.open_documents = lambda: []
    doc.__dict__.setdefault("Windows_", SimpleNamespace())
    return app


def test_a_document_opens_with_macros_off_and_the_setting_comes_back(make_doc, word, tmp_path):
    path = tmp_path / "a.docx"
    path.write_bytes(b"x")
    app = _with_app(word, make_doc("One."))
    out = json.loads(_common.run(documents._open_document_impl, str(path), False, True))
    assert "error" not in out
    assert app.seen == [3] and app.AutomationSecurity == 1


def test_a_template_opens_with_macros_off(make_doc, word, tmp_path):
    path = tmp_path / "a.dotm"
    path.write_bytes(b"x")
    app = _with_app(word, make_doc("One."))
    out = json.loads(_common.run(documents._create_document_impl, str(path), None))
    assert "error" not in out
    assert app.seen == [3] and app.AutomationSecurity == 1


def test_the_macro_setting_comes_back_when_opening_fails(make_doc, word, tmp_path):
    path = tmp_path / "a.docx"
    path.write_bytes(b"x")
    app = _with_app(word, make_doc("One."))

    def refuses(**kwargs):
        raise RuntimeError("The file is damaged.")

    app.Open = refuses
    assert "damaged" in json.loads(_common.run(
        documents._open_document_impl, str(path), False, True))["error"]
    assert app.AutomationSecurity == 1


@pytest.mark.parametrize("path", ["\\\\\\\\evil\\\\share\\\\a.docx", "//evil/share/a.docx"])
def test_a_network_path_is_refused(path):
    with pytest.raises(ValueError, match="network path"):
        _common.local_path(path)


def test_the_setting_allows_a_network_path(monkeypatch):
    monkeypatch.setattr(_common, "ALLOW_NETWORK_PATHS", True)
    assert _common.local_path("//server/a.docx") == "//server/a.docx"


def test_a_local_path_is_left_alone():
    assert _common.local_path("C:\\\\a.docx") == "C:\\\\a.docx"
    assert _common.local_path("/Users/a/b.docx") == "/Users/a/b.docx"


def test_a_new_document_saves_only_as_a_format_word_writes():
    with pytest.raises(ValueError, match="must end in"):
        documents.CreateDocumentInput(save_as=os.path.abspath("a.exe"))
    assert documents.CreateDocumentInput(save_as=os.path.abspath("a.docx")).save_as
