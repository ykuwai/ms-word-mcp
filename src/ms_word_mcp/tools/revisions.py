"""Tracked changes.

With tracking on, what is deleted stays in the document, marked, until the
change is accepted, and Range.Text holds it while the markup is shown; taking
the deletions' spans out gives the text Word shows (measured on both platforms:
scripts/probes/windows_comments_revisions.py and mac_comments_revisions.py).
Accepting or rejecting a change renumbers the ones after it.
https://learn.microsoft.com/en-us/office/vba/api/word.document.trackrevisions
https://learn.microsoft.com/en-us/office/vba/api/word.revision
https://learn.microsoft.com/en-us/office/vba/api/word.wdrevisiontype
"""

from __future__ import annotations

import time
from datetime import datetime

from pydantic import Field, model_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, annotations, call, mac_impl
from ._word import clean_text

#: How much of a change's text an answer holds.
TEXT_SHOWN = 500

#: Seconds a listing may take before it answers what it has. The Mac worker
#: gives up waiting at 84 s, and on the Mac each question about a change takes
#: longer the further along it is (0.46 s at the 181st, docs/PROBES.md).
LIST_SECONDS = 40.0

#: WdRevisionType, by value. On the Mac the same types go by name (MAC_KINDS).
KINDS = {
    1: "insert", 2: "delete", 3: "format", 4: "paragraph_number", 5: "field_display",
    6: "reconcile", 7: "conflict", 8: "style", 9: "replace", 10: "paragraph_format",
    11: "table_format", 12: "section_format", 13: "style_definition", 14: "moved_from",
    15: "moved_to", 16: "cell_insert", 17: "cell_delete", 18: "cell_merge", 19: "cell_split",
    20: "conflict_insert", 21: "conflict_delete",
}

MAC_KINDS = {
    "revision_insert": "insert", "revision_delete": "delete", "revision_property": "format",
    "revision_paragraph_number": "paragraph_number", "revision_display_field": "field_display",
    "revision_reconcile": "reconcile", "revision_conflict": "conflict", "revision_style": "style",
    "revision_replace": "replace", "revision_paragraph_property": "paragraph_format",
    "revision_table_property": "table_format", "revision_section_property": "section_format",
    "revision_style_definition": "style_definition", "revision_move_from": "moved_from",
    "revision_move_to": "moved_to", "revision_cell_insertion": "cell_insert",
    "revision_cell_deletion": "cell_delete", "revision_cell_merge": "cell_merge",
    "revision_cell_split": "cell_split", "revision_conflict_insert": "conflict_insert",
    "revision_conflict_delete": "conflict_delete",
}


class SetTrackChangesInput(Input):
    on: bool = Field(..., description="True to track changes from now on, false to stop.")


class ListRevisionsInput(Input):
    start: int = Field(1, ge=1, description="The change to start from (numbered from 1).")
    limit: int = Field(100, ge=1, le=200, description="How many to read at most.")


class ResolveRevisionsInput(Input):
    numbers: list[int] | None = Field(None, min_length=1, max_length=200, description=(
        "Which changes, as word_list_revisions numbers them."))
    every: bool = Field(False, description="Every change in the document's body.")

    @model_validator(mode="after")
    def _one_way(self):
        if (self.numbers is None) == (not self.every):
            raise ValueError("Give `numbers`, or `every` true: one of the two.")
        if self.numbers is not None and min(self.numbers) < 1:
            raise ValueError("Changes are numbered from 1.")
        return self


def when(value) -> str | None:
    """A change's time as Word recorded it: local time, with no zone."""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT%H:%M:%S")
    return None if value is None else str(value)


def shown(text) -> dict:
    """A change's text, cut short when it is long."""
    text = clean_text(text)
    if len(text) <= TEXT_SHOWN:
        return {"text": text}
    return {"text": text[:TEXT_SHOWN], "text_cut": True}


def check_numbers(numbers: list[int], count: int) -> list[int]:
    """The numbers from the highest down, so that each holds until its turn."""
    bad = sorted({n for n in numbers if not 1 <= n <= count})
    if bad:
        raise ValueError(f"There is no change {bad[0]}: the document has {count} "
                         "(numbered from 1; word_list_revisions lists them). Nothing was changed.")
    return sorted(set(numbers), reverse=True)


def listing(describe, count: int, tracking: bool, start: int, limit: int) -> dict:
    """Changes `start` on, `describe(number)` each, until `limit` or the time runs out."""
    began = time.monotonic()
    last = min(count, start + limit - 1)
    out = []
    for number in range(start, last + 1):
        out.append(describe(number))
        if time.monotonic() > began + LIST_SECONDS:
            break
    result = {"success": True, "tracking": tracking, "count": count, "revisions": out}
    reached = start + len(out) - 1
    if reached < last:
        result["complete"] = False
        result["warnings"] = [
            f"The time for one call ran out after change {reached}, before change {last}. "
            f"Call again with start={reached + 1} for the rest."]
    if out and reached < count:
        result["more"] = f"Changes {reached + 1} to {count} follow; pass start={reached + 1}."
    return result


def resolved(verb: str, done: int, before: int, left: int, paragraphs: int) -> dict:
    result = {"success": True, verb: done, "left": left, "paragraphs": paragraphs}
    if done and left != before - done:
        result["note"] = (f"Word went from {before} changes to {left}: one change can take "
                          "others in its text with it. The numbers have moved; list them again.")
    elif left:
        result["note"] = "The changes after these have new numbers; list them again."
    return result


# -- Windows ---------------------------------------------------------------------------

def _describe(doc, number: int, revision) -> dict:
    rng = revision.Range
    return {"number": number, "kind": KINDS.get(revision.Type, str(revision.Type)),
            "author": revision.Author, "date": when(revision.Date),
            "index": doc.Range(0, rng.Start + 1).Paragraphs.Count, **shown(rng.Text),
            **({"description": revision.FormatDescription} if revision.FormatDescription else {})}


def _set_track_changes_impl(on):
    doc = backend.word.document()
    doc.TrackRevisions = on
    return {"success": True, "tracking": bool(doc.TrackRevisions),
            "count": doc.Revisions.Count}


def _list_revisions_impl(start, limit):
    doc = backend.word.document()
    revisions = doc.Revisions
    return listing(lambda i: _describe(doc, i, revisions.Item(i)), revisions.Count,
                   bool(doc.TrackRevisions), start, limit)


def _resolve(verb, method, whole, numbers, every):
    doc = backend.word.document()
    revisions = doc.Revisions
    before = revisions.Count
    if every:
        getattr(revisions, whole)()
        done = before
    else:
        order = check_numbers(numbers, before)
        for number in order:
            getattr(revisions.Item(number), method)()
        done = len(order)
    return resolved(verb, done, before, doc.Revisions.Count, doc.Paragraphs.Count)


def _accept_revisions_impl(numbers, every):
    return _resolve("accepted", "Accept", "AcceptAll", numbers, every)


def _reject_revisions_impl(numbers, every):
    return _resolve("rejected", "Reject", "RejectAll", numbers, every)


# -- Mac -------------------------------------------------------------------------------

def _mac_count(doc) -> int:
    from appscript import k

    return int(doc.count(each=k.revision))


def _mac_kind(value) -> str:
    name = getattr(value, "name", str(value))
    return MAC_KINDS.get(name, name)


def _mac_describe(doc, number: int, revision) -> dict:
    """Each question about a document's change takes longer the further along it
    is, so its properties come in one, and its text is read by position."""
    from appscript import k

    props = revision.properties.get()
    rng = revision.text_object
    start, end = rng.start_of_content.get(), rng.end_of_content.get()
    described = {"number": number, "kind": _mac_kind(props[k.revision_type]),
                 "author": props[k.author], "date": when(props[k.date_value]),
                 "index": int(doc.create_range(start=0, end_=start + 1).count(each=k.paragraph)),
                 **shown(doc.create_range(start=start, end_=end).content.get())}
    description = props.get(k.format_description)
    if isinstance(description, str) and description:
        described["description"] = description
    return described


@mac_impl(_set_track_changes_impl)
def _mac_set_track_changes(on):
    doc = backend.word.document()
    doc.track_revisions.set(on)
    return {"success": True, "tracking": bool(doc.track_revisions.get()),
            "count": _mac_count(doc)}


@mac_impl(_list_revisions_impl)
def _mac_list_revisions(start, limit):
    doc = backend.word.document()
    return listing(lambda i: _mac_describe(doc, i, doc.revisions[i]), _mac_count(doc),
                   bool(doc.track_revisions.get()), start, limit)


def _mac_resolve(verb, method, whole, numbers, every):
    from appscript import k

    doc = backend.word.document()
    before = _mac_count(doc)
    if every:
        getattr(doc, whole)()
        done = before
    else:
        order = check_numbers(numbers, before)
        for number in order:
            getattr(doc.revisions[number], method)()
        done = len(order)
    return resolved(verb, done, before, _mac_count(doc), int(doc.count(each=k.paragraph)))


@mac_impl(_accept_revisions_impl)
def _mac_accept_revisions(numbers, every):
    return _mac_resolve("accepted", "accept", "accept_all_revisions", numbers, every)


@mac_impl(_reject_revisions_impl)
def _mac_reject_revisions(numbers, every):
    return _mac_resolve("rejected", "reject", "reject_all_revisions", numbers, every)


def register(mcp) -> None:
    @mcp.tool(name="word_set_track_changes",
              annotations=annotations("Track changes", WRITES))
    async def word_set_track_changes(params: SetTrackChangesInput) -> str:
        """Turn change tracking on or off for the document. While it is on,
        every edit, by the user or by these tools, is marked for review and
        can be accepted or rejected later."""
        return await call(_set_track_changes_impl, params.on, undo="Track changes")

    @mcp.tool(name="word_list_revisions",
              annotations=annotations("List tracked changes", READ_ONLY))
    async def word_list_revisions(params: ListRevisionsInput) -> str:
        """The tracked changes in the document's body, in order: each one's
        number, kind (insert, delete, format...), author, date, the paragraph
        it starts in (`index`) and its text. The answer says whether tracking
        is on (`tracking`) and how many changes there are (`count`)."""
        return await call(_list_revisions_impl, params.start, params.limit)

    @mcp.tool(name="word_accept_revisions",
              annotations=annotations("Accept tracked changes", DESTRUCTIVE))
    async def word_accept_revisions(params: ResolveRevisionsInput) -> str:
        """Accept tracked changes, by number or every one: inserted text
        stays, deleted text goes for good. The changes after them renumber,
        so list them again before the next call."""
        return await call(_accept_revisions_impl, params.numbers, params.every,
                          undo="Accept changes")

    @mcp.tool(name="word_reject_revisions",
              annotations=annotations("Reject tracked changes", DESTRUCTIVE))
    async def word_reject_revisions(params: ResolveRevisionsInput) -> str:
        """Reject tracked changes, by number or every one: inserted text
        goes, deleted text comes back. The changes after them renumber, so
        list them again before the next call."""
        return await call(_reject_revisions_impl, params.numbers, params.every,
                          undo="Reject changes")
