"""Word itself and its documents: connect, list, choose, open, create, save, close."""

from __future__ import annotations

import os

from pydantic import Field, field_validator

from .. import backend
from ._common import READ_ONLY, WRITES, DESTRUCTIVE, Input, annotations, call

# WdSaveFormat values for the formats save-as takes, by file extension.
SAVE_FORMATS = {".docx": 16, ".docm": 13, ".dotx": 14, ".doc": 0, ".rtf": 6,
                ".txt": 2, ".pdf": 17, ".odt": 23, ".htm": 8, ".html": 8}


def _doc_summary(doc) -> dict:
    return {"name": doc.Name, "full_name": doc.FullName, "saved": bool(doc.Saved),
            "read_only": bool(doc.ReadOnly)}


def _require_absolute(path: str) -> str:
    path = os.path.expanduser(path)
    if not os.path.isabs(path):
        raise ValueError(f"'{path}' is not a full path. Pass the whole path to the file.")
    return path


# -- app ------------------------------------------------------------------

def _get_app_info_impl():
    word = backend.word
    app = word.get_app()
    docs = word.open_documents()
    active = None
    if docs:
        try:
            active = app.ActiveDocument.FullName
        except Exception:  # noqa: BLE001 - Word has no active document
            active = None
    return {
        "success": True,
        "version": app.Version,
        "build": app.Build,
        "platform": backend.PLATFORM_NAME,
        "documents": [{"index": d.index, "name": d.name, "full_name": d.full_name}
                      for d in docs],
        "active_document": active,
        "target_document": word.target_full_name,
        "started_by_server": word.launched_by_us,
    }


def _list_documents_impl():
    word = backend.word
    app = word.get_app()
    out = []
    for d in word.open_documents():
        doc = app.Documents.Item(d.index)
        out.append({"index": d.index, **_doc_summary(doc)})
    return {"success": True, "documents": out, "target_document": word.target_full_name}


class ActivateDocumentInput(Input):
    document: str = Field(..., min_length=1, description=(
        "The document to work on from now on: its full path, file name, or file "
        "name without extension, as word_list_documents gives it."))


def _activate_document_impl(document: str):
    from ..backend.documents import pick_document

    word = backend.word
    found = pick_document(word.open_documents(), document)
    word.target_full_name = found.full_name
    return {"success": True, "name": found.name, "full_name": found.full_name,
            "note": ("Calls without `document` now work on this one. Its window was "
                     "not brought forward, so the user's own window stays as it was.")}


# -- open, create ------------------------------------------------------------

class OpenDocumentInput(Input):
    path: str = Field(..., min_length=1, description="Full path of the file to open.")
    read_only: bool = Field(False, description="Open it read-only.")
    visible: bool = Field(True, description=(
        "Open it in a window. False opens it without one, for reading only."))

    @field_validator("path")
    @classmethod
    def _absolute(cls, value: str) -> str:
        return _require_absolute(value)


def _open_document_impl(path: str, read_only: bool, visible: bool):
    if not os.path.exists(path):
        raise ValueError(f"There is no file at '{path}'.")
    word = backend.word
    app = word.get_app(allow_launch=True)
    from ..backend.documents import full_name_key

    for d in word.open_documents():
        if full_name_key(d.full_name) == full_name_key(path):
            doc = app.Documents.Item(d.index)
            word.target_full_name = doc.FullName
            return {"success": True, **_doc_summary(doc), "already_open": True}
    doc = app.Documents.Open(FileName=path, ConfirmConversions=False, ReadOnly=read_only,
                             AddToRecentFiles=False, Visible=visible)
    word.target_full_name = doc.FullName
    return {"success": True, **_doc_summary(doc), "paragraphs": doc.Paragraphs.Count,
            "note": "This document is now the target of calls without `document`."}


class CreateDocumentInput(Input):
    template: str | None = Field(None, description=(
        "Full path of a template (.dotx, .dotm) or document to base it on. "
        "Leave out for Word's Normal template."))
    save_as: str | None = Field(None, description=(
        "Full path to save it to straight away (.docx). Saving early is safer."))

    @field_validator("template", "save_as")
    @classmethod
    def _absolute(cls, value: str | None) -> str | None:
        return None if value is None else _require_absolute(value)


def _create_document_impl(template: str | None, save_as: str | None):
    if template and not os.path.exists(template):
        raise ValueError(f"There is no template at '{template}'.")
    if save_as and os.path.exists(save_as):
        raise ValueError(
            f"'{save_as}' already exists. Choose another name; nothing was created.")
    word = backend.word
    app = word.get_app(allow_launch=True)
    doc = app.Documents.Add(Template=template) if template else app.Documents.Add()
    if save_as:
        doc.SaveAs2(FileName=save_as, FileFormat=SAVE_FORMATS.get(
            os.path.splitext(save_as)[1].lower(), 16))
    word.target_full_name = doc.FullName
    return {"success": True, **_doc_summary(doc),
            "note": "This document is now the target of calls without `document`."}


# -- save, close -------------------------------------------------------------

def _save_document_impl():
    doc = backend.word.document()
    if not doc.Path:
        raise ValueError(
            f"'{doc.Name}' has never been saved, so it has no file to save to. "
            "Use word_save_document_as with a full path.")
    doc.Save()
    return {"success": True, **_doc_summary(doc)}


class SaveDocumentAsInput(Input):
    path: str = Field(..., min_length=1, description=(
        "Full path to save to. The extension picks the format: "
        + ", ".join(sorted(SAVE_FORMATS)) + ". A .pdf is an export: the document "
        "keeps its own file."))
    overwrite: bool = Field(False, description="Replace a file already at `path`.")

    @field_validator("path")
    @classmethod
    def _known_format(cls, value: str) -> str:
        value = _require_absolute(value)
        ext = os.path.splitext(value)[1].lower()
        if ext not in SAVE_FORMATS:
            raise ValueError(f"'{ext or value}' is not a format this saves to. Use one of "
                             + ", ".join(sorted(SAVE_FORMATS)) + ".")
        return value


def _save_document_as_impl(path: str, overwrite: bool):
    if os.path.exists(path) and not overwrite:
        raise ValueError(f"'{path}' already exists. Pass overwrite=true to replace it.")
    folder = os.path.dirname(path)
    if folder and not os.path.isdir(folder):
        raise ValueError(f"The folder '{folder}' does not exist.")
    doc = backend.word.document()
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        doc.ExportAsFixedFormat(OutputFileName=path, ExportFormat=17,
                                OpenAfterExport=False)
    else:
        doc.SaveAs2(FileName=path, FileFormat=SAVE_FORMATS[ext])
        backend.word.target_full_name = doc.FullName
    if not os.path.exists(path):
        raise RuntimeError(f"Word answered without error, but there is no file at '{path}'.")
    return {"success": True, "path": path, "bytes": os.path.getsize(path),
            "document": doc.FullName}


class CloseDocumentInput(Input):
    save: bool | None = Field(None, description=(
        "true saves first, false discards unsaved changes. Leave out to refuse "
        "when there are unsaved changes."))


def _close_document_impl(save: bool | None):
    word = backend.word
    doc = word.document()
    name, full_name = doc.Name, doc.FullName
    if save is None and not doc.Saved:
        raise ValueError(
            f"'{name}' has unsaved changes. Pass save=true to save them or save=false "
            "to discard them.")
    if save and not doc.Path:
        raise ValueError(f"'{name}' has never been saved. Save it with "
                         "word_save_document_as first, or pass save=false.")
    # WdSaveOptions: wdSaveChanges -1, wdDoNotSaveChanges 0.
    doc.Close(SaveChanges=-1 if save else 0)
    if word.target_full_name == full_name:
        word.target_full_name = None
    return {"success": True, "closed": full_name}


def register(mcp) -> None:
    @mcp.tool(name="word_get_app_info", annotations=annotations("Word info", READ_ONLY))
    async def word_get_app_info() -> str:
        """Word's version, its open documents, and which one calls work on."""
        return await call(_get_app_info_impl)

    @mcp.tool(name="word_list_documents", annotations=annotations("List documents", READ_ONLY))
    async def word_list_documents() -> str:
        """The documents open in Word, with their full paths and whether each is saved."""
        return await call(_list_documents_impl)

    @mcp.tool(name="word_activate_document",
              annotations=annotations("Choose the document", READ_ONLY))
    async def word_activate_document(params: ActivateDocumentInput) -> str:
        """Choose the document that calls without `document` work on.

        Its window is not brought forward: the user keeps working where they
        are. The choice is shared by every conversation this server serves, so
        pass `document` on each call as well.
        """
        return await call(_activate_document_impl, params.document)

    @mcp.tool(name="word_open_document", annotations=annotations("Open a document", WRITES))
    async def word_open_document(params: OpenDocumentInput) -> str:
        """Open a file in Word (starting Word if needed) and make it the target."""
        return await call(_open_document_impl, params.path, params.read_only,
                                   params.visible)

    @mcp.tool(name="word_create_document", annotations=annotations("New document", WRITES))
    async def word_create_document(params: CreateDocumentInput) -> str:
        """Create a new document, optionally from a template, and make it the target.

        Pass `save_as` to save it at once: an unsaved document has no file to
        come back to.
        """
        return await call(_create_document_impl, params.template,
                                   params.save_as)

    @mcp.tool(name="word_save_document", annotations=annotations("Save", WRITES))
    async def word_save_document() -> str:
        """Save the document to its own file."""
        return await call(_save_document_impl)

    @mcp.tool(name="word_save_document_as", annotations=annotations("Save as", WRITES))
    async def word_save_document_as(params: SaveDocumentAsInput) -> str:
        """Save the document under a new path, or export it as PDF.

        The extension picks the format. Checks that the file exists afterwards.
        """
        return await call(_save_document_as_impl, params.path,
                                   params.overwrite)

    @mcp.tool(name="word_close_document", annotations=annotations("Close", DESTRUCTIVE))
    async def word_close_document(params: CloseDocumentInput) -> str:
        """Close the document. Refuses when there are unsaved changes unless
        `save` says what to do with them."""
        return await call(_close_document_impl, params.save)
