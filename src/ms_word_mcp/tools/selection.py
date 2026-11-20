"""What the user has selected in Word.

The tools write through ranges and never move the selection, so the user can
go on working. Reading it lets the model work on what the user points at:
"rewrite this", "check the paragraph I'm in". The selection is the active
window's, so it can be in another document than the one calls go to; the
answer says so.
https://learn.microsoft.com/en-us/office/vba/api/word.selection
https://learn.microsoft.com/en-us/office/vba/api/word.wdstorytype
"""

from __future__ import annotations

from .. import backend
from ..backend.com import names_of
from ._common import READ_ONLY, Input, annotations, call, mac_impl
from ._word import clean_text

#: How much of the selected text an answer holds.
TEXT_SHOWN = 2000

#: WdStoryType, by value. On the Mac the same stories go by name.
STORIES = {1: "main", 2: "footnotes", 3: "endnotes", 4: "comments", 5: "text_box",
           6: "header", 7: "header", 8: "footer", 9: "footer", 10: "header", 11: "footer"}


class NoInput(Input):
    pass


def mac_story(name: str) -> str:
    """A Word for Mac story type's name as STORIES names it."""
    for word, story in (("main text", "main"), ("footnote", "footnotes"), ("endnote", "endnotes"),
                        ("comments", "comments"), ("text frame", "text_box"), ("header", "header"),
                        ("footer", "footer")):
        if word.replace(" ", "_") in name:
            return story
    return "other"


def selection_answer(name, full_name, story, start, end, text, paragraphs, target) -> dict:
    """What word_get_selection answers, from either platform's reads.

    `paragraphs` is (first, last) in the main text, else None; `target` is the
    full name of the document calls without `document` go to, or None."""
    # With only the cursor there, Word for Mac answers the character after it
    # (measured); Windows answers nothing.
    text = "" if start == end else clean_text(text)
    result = {"success": True, "document": name, "full_name": full_name, "story": story,
              "empty": start == end, "text": text[:TEXT_SHOWN]}
    if len(text) > TEXT_SHOWN:
        result["text_cut"] = True
    if paragraphs is not None:
        result["start"], result["end"] = paragraphs
    notes = []
    if start == end:
        notes.append("Nothing is selected: the cursor is "
                     + (f"in paragraph {paragraphs[0]}." if paragraphs else f"in the {story}."))
    if story != "main":
        notes.append(f"The selection is in the {story}, not in the document's main text, "
                     "so it has no paragraph numbers.")
    if target is not None and target != full_name:
        notes.append(f"It is in {name}, and calls without `document` go to another document; "
                     f"pass document={full_name!r} to work on it.")
    if notes:
        result["note"] = " ".join(notes)
    return result


def _paragraphs(count_to, start: int, end: int) -> tuple[int, int]:
    """The first and last paragraph of [start, end): `count_to(position)` counts
    the paragraphs from the start of the text up to a position. A selection
    that ends where a paragraph begins does not reach into it."""
    first = count_to(start + 1)
    return first, (count_to(end) if end > start else first)


def _target_full_name(full) -> str | None:
    try:
        return full(backend.word.document())
    except Exception:  # noqa: BLE001 - no document chosen yet: nothing to compare
        return None


def _get_selection_impl():
    app = backend.word.get_app()
    if not app.Documents.Count:
        raise ValueError("No document is open in Word.")
    sel = app.Selection
    doc = sel.Document
    start, end = sel.Range.Start, sel.Range.End
    story = STORIES.get(sel.StoryType, "other")
    paragraphs = (_paragraphs(lambda at: doc.Range(0, at).Paragraphs.Count, start, end)
                  if story == "main" else None)
    return selection_answer(*names_of(doc), story, start, end, sel.Range.Text,
                            paragraphs, _target_full_name(lambda d: names_of(d)[1]))


@mac_impl(_get_selection_impl)
def _mac_get_selection():
    from . import mac as m

    k = m._k()
    app = backend.word.get_app()
    if not int(app.count(each=k.document)):
        raise ValueError("No document is open in Word.")
    sel = app.selection
    doc = app.documents[sel.document.name.get()]
    start, end = sel.selection_start.get(), sel.selection_end.get()
    story = mac_story(m._name(sel.story_type.get()))
    paragraphs = (_paragraphs(lambda at: int(doc.create_range(start=0, end_=at).count(each=k.paragraph)),
                              start, end) if story == "main" else None)
    summary = m._summary(doc)
    return selection_answer(summary["name"], summary["full_name"], story, start, end,
                            sel.content.get(), paragraphs,
                            _target_full_name(lambda d: m._summary(d)["full_name"]))


def register(mcp) -> None:
    @mcp.tool(name="word_get_selection", annotations=annotations("What is selected", READ_ONLY))
    async def word_get_selection(params: NoInput) -> str:
        """What the user has selected in Word, or where the cursor is: the
        document, the paragraphs it spans (`start`, `end`), its text, and
        whether it is in the main text or in a header, a footnote, a comment.
        For "rewrite this" or "check what I've selected". It may be in another
        document than the one calls go to; the answer says so."""
        return await call(_get_selection_impl)
