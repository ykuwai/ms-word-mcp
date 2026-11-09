"""Footnotes and endnotes.

Word numbers a document's notes itself, in the order of their marks, so adding
or deleting one renumbers the ones after it. A note's mark is one character of
the paragraph, U+0002 in the text Word answers.
https://learn.microsoft.com/en-us/office/vba/api/word.footnotes.add
https://learn.microsoft.com/en-us/office/vba/api/word.endnotes.add

On Word for Mac a note is made with `make new footnote at <document>` and its
place as the `text object` property; made `at` a range or a paragraph it went
to the start of the document instead. A new note holds the text "No Reference"
until it is written (both measured: scripts/probes/mac_notes_links.py).
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, field_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, annotations, call, follow, mac_impl
from ._word import clean_text, locate_spans, paragraph

Kind = Literal["footnote", "endnote"]
KIND_FIELD = Field("footnote", description="A footnote, or an endnote.")


def _one_paragraph(text: str) -> str:
    if "\r" in text or "\n" in text:
        raise ValueError("A note is written as one paragraph here: take the line breaks out.")
    return text


class InsertNoteInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the note's mark goes in.")
    text: str = Field(..., min_length=1, max_length=5000, description="The note's text.")
    kind: Kind = KIND_FIELD
    after: str | None = Field(None, min_length=1, max_length=255, description=(
        "Put the mark right after the first occurrence of this text in the paragraph. "
        "Left out: at the end of the paragraph's text."))

    _text = field_validator("text")(_one_paragraph)


class ListNotesInput(Input):
    kind: Kind = KIND_FIELD
    start: int = Field(1, ge=1, description="The note to start from (numbered from 1).")
    limit: int = Field(100, ge=1, le=200, description="How many to read at most.")


class SetNoteInput(Input):
    number: int = Field(..., ge=1, description="Which note, as word_list_notes numbers them.")
    text: str = Field(..., min_length=1, max_length=5000, description="The note's new text.")
    kind: Kind = KIND_FIELD

    _text = field_validator("text")(_one_paragraph)


class DeleteNoteInput(Input):
    number: int = Field(..., ge=1, description="Which note, as word_list_notes numbers them.")
    kind: Kind = KIND_FIELD


def check_number(kind: str, number: int, count: int) -> None:
    if not 1 <= number <= count:
        raise ValueError(f"There is no {kind} {number}: the document has {count} "
                         "(numbered from 1; word_list_notes lists them). Nothing was changed.")


def not_in_paragraph(after: str, index: int) -> ValueError:
    return ValueError(f"'{after}' is not in paragraph {index}. Nothing was inserted.")


def listing(kind: str, out: list[dict], footnotes: int, endnotes: int, start: int) -> dict:
    """The notes read, the counts of both kinds, and what is left to read."""
    count = footnotes if kind == "footnote" else endnotes
    result = {"success": True, "kind": kind, "footnotes": footnotes, "endnotes": endnotes,
              "notes": out}
    last = start + len(out) - 1
    if out and last < count:
        result["more"] = f"Notes {last + 1} to {count} follow; pass start={last + 1}."
    return result


def _notes(doc, kind):
    return doc.Footnotes if kind == "footnote" else doc.Endnotes


def _note(doc, kind, number):
    notes = _notes(doc, kind)
    check_number(kind, number, notes.Count)
    return notes.Item(number)


def _describe(doc, note) -> dict:
    return {"number": note.Index, "index": doc.Range(0, note.Reference.Start + 1).Paragraphs.Count,
            "text": clean_text(note.Range.Text)}


def _insert_note_impl(index, text, kind, after):
    doc = backend.word.document()
    rng = paragraph(doc, index).Range
    at = rng.End - 1
    if after is not None:
        raw = rng.Text
        if raw.find(after) < 0:
            raise not_in_paragraph(after, index)
        hidden = max(0, rng.End - rng.Start - len(raw.encode("utf-16-le")) // 2)
        at = locate_spans(raw, [after], lambda a, b: doc.Range(a, b).Text, rng.Start, hidden)[0][1]
    note = _notes(doc, kind).Add(doc.Range(at, at))
    note.Range.Text = text
    follow(doc, note.Reference)
    return {"success": True, "kind": kind, **_describe(doc, note)}


def _list_notes_impl(kind, start, limit):
    doc = backend.word.document()
    notes = _notes(doc, kind)
    out = [_describe(doc, notes.Item(i))
           for i in range(start, min(notes.Count, start + limit - 1) + 1)]
    return listing(kind, out, doc.Footnotes.Count, doc.Endnotes.Count, start)


def _set_note_impl(number, text, kind):
    doc = backend.word.document()
    note = _note(doc, kind, number)
    note.Range.Text = text
    follow(doc, note.Reference)
    return {"success": True, "kind": kind, **_describe(doc, note)}


def _delete_note_impl(number, kind):
    doc = backend.word.document()
    _note(doc, kind, number).Delete()
    return {"success": True, "kind": kind, "deleted": number,
            "left": _notes(doc, kind).Count}


# -- Mac -------------------------------------------------------------------------------

def _mac_class(kind):
    from appscript import k

    return k.footnote if kind == "footnote" else k.endnote


def _mac_notes(doc, kind):
    return doc.footnotes if kind == "footnote" else doc.endnotes


def _mac_count(doc, kind) -> int:
    return int(doc.count(each=_mac_class(kind)))


def _mac_note(doc, kind, number):
    check_number(kind, number, _mac_count(doc, kind))
    return _mac_notes(doc, kind)[number]


def _mac_describe(doc, note) -> dict:
    from appscript import k

    from . import mac as m

    at = note.note_reference.start_of_content.get()
    return {"number": int(note.entry_index.get()),
            "index": int(doc.create_range(start=0, end_=at + 1).count(each=k.paragraph)),
            "text": m._clean(note.text_object.content.get())}


@mac_impl(_insert_note_impl)
def _mac_insert_note(index, text, kind, after):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    m._para(doc, index)
    start, end = m._bounds(doc, index)
    at = end - 1
    if after is not None:
        raw = doc.paragraphs[index].text_object.content.get()
        if raw.find(after) < 0:
            raise not_in_paragraph(after, index)

        def read(a, b):
            try:
                return doc.create_range(start=a, end_=b).content.get()
            except Exception:  # noqa: BLE001 - half a character: not here
                return None

        hidden = max(0, end - start - m._u16(raw, len(raw)))
        at = locate_spans(raw, [after], read, start, hidden)[0][1]
    note = doc.make(new=_mac_class(kind), at=doc,
                    with_properties={k.text_object: doc.create_range(start=at, end_=at)})
    note.text_object.content.set(text)
    return {"success": True, "kind": kind, **_mac_describe(doc, note)}


@mac_impl(_list_notes_impl)
def _mac_list_notes(kind, start, limit):
    doc = backend.word.document()
    count = _mac_count(doc, kind)
    notes = _mac_notes(doc, kind)
    out = [_mac_describe(doc, notes[i]) for i in range(start, min(count, start + limit - 1) + 1)]
    return listing(kind, out, _mac_count(doc, "footnote"), _mac_count(doc, "endnote"), start)


@mac_impl(_set_note_impl)
def _mac_set_note(number, text, kind):
    doc = backend.word.document()
    note = _mac_note(doc, kind, number)
    note.text_object.content.set(text)
    return {"success": True, "kind": kind, **_mac_describe(doc, note)}


@mac_impl(_delete_note_impl)
def _mac_delete_note(number, kind):
    doc = backend.word.document()
    _mac_note(doc, kind, number).delete()
    return {"success": True, "kind": kind, "deleted": number, "left": _mac_count(doc, kind)}


def register(mcp) -> None:
    @mcp.tool(name="word_insert_note", annotations=annotations("Insert a note", WRITES))
    async def word_insert_note(params: InsertNoteInput) -> str:
        """Add a footnote or an endnote: its mark at the end of a paragraph's
        text or right after some text in it. Word numbers the notes in the order
        of their marks, so the ones after it renumber; `number` is where this
        one landed. A paragraph's text shows a note's mark as the character
        U+0002."""
        p = params
        return await call(_insert_note_impl, p.index, p.text, p.kind, p.after,
                          undo=f"Insert {p.kind}")

    @mcp.tool(name="word_list_notes", annotations=annotations("List notes", READ_ONLY))
    async def word_list_notes(params: ListNotesInput) -> str:
        """The footnotes, or the endnotes, in order: each one's number, the
        paragraph its mark is in (`index`) and its text. The answer also counts
        both kinds."""
        p = params
        return await call(_list_notes_impl, p.kind, p.start, p.limit)

    @mcp.tool(name="word_set_note", annotations=annotations("Change a note", WRITES))
    async def word_set_note(params: SetNoteInput) -> str:
        """Replace the text of a footnote or an endnote. Its mark stays where it
        is."""
        p = params
        return await call(_set_note_impl, p.number, p.text, p.kind, undo=f"Change {p.kind}")

    @mcp.tool(name="word_delete_note", annotations=annotations("Delete a note", DESTRUCTIVE))
    async def word_delete_note(params: DeleteNoteInput) -> str:
        """Delete a footnote or an endnote, its mark and its text. The notes
        after it renumber."""
        p = params
        return await call(_delete_note_impl, p.number, p.kind, undo=f"Delete {p.kind}")
