"""Reading and writing a document's text, a paragraph at a time.

Everything goes through Ranges, never through the Selection, so the user's
cursor stays where they left it while a tool writes elsewhere in the same
document.
"""

from __future__ import annotations

import time

from pydantic import Field, model_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, Trimmed, annotations, call, follow
from .shapes import leading_anchors
from ._word import (
    WD_COLLAPSE_END, WD_FIND_STOP, WD_OUTLINE_BODY,
    WD_STATISTIC_CHARACTERS, WD_STATISTIC_PAGES, WD_STATISTIC_WORDS,
    check_span, clean_text, deletion_spans, index_of, inside_spans, paragraph, paragraph_info,
    paragraphs_in, visible_text, BUILTIN_STYLES, resolve_style, style_name,
)

MAX_PARAGRAPHS_PER_READ = 200
WD_WITHIN_TABLE = 12  # WdInformation


# -- reading -------------------------------------------------------------------

def _get_document_info_impl():
    doc = backend.word.document()
    sections = doc.Sections
    setup = doc.PageSetup
    return {
        "success": True,
        "name": doc.Name,
        "full_name": doc.FullName,
        "saved": bool(doc.Saved),
        "paragraphs": doc.Paragraphs.Count,
        "pages": doc.ComputeStatistics(WD_STATISTIC_PAGES),
        "words": doc.ComputeStatistics(WD_STATISTIC_WORDS),
        "characters": doc.ComputeStatistics(WD_STATISTIC_CHARACTERS),
        "sections": sections.Count,
        "tables": doc.Tables.Count,
        "page": {"width_pt": setup.PageWidth, "height_pt": setup.PageHeight,
                 "margins_pt": {"top": setup.TopMargin, "bottom": setup.BottomMargin,
                                "left": setup.LeftMargin, "right": setup.RightMargin}},
        "track_revisions": bool(doc.TrackRevisions),
    }


WD_REF_TYPE_HEADING = 1  # WdReferenceType


class OutlineInput(Input):
    every_paragraph: bool = Field(False, description=(
        "Also find paragraphs given an outline level without a heading style, by "
        "reading every paragraph: seconds on a long document."))


def _fast_outline(doc, count: int) -> list[dict] | None:
    """The headings from Word's heading list, each read back; None if one is not found."""
    from ._outline import locate_headings, split_texts

    items = list(doc.GetCrossReferenceItems(WD_REF_TYPE_HEADING) or ())
    if not items:
        return []
    paragraphs, texts, levels, styles = {}, {}, {}, {}
    heading_styles = {doc.Styles.Item(BUILTIN_STYLES[f"heading {n}"]).NameLocal
                      for n in range(1, 10)}

    def para(i):
        if i not in paragraphs:
            paragraphs[i] = doc.Paragraphs.Item(i)
        return paragraphs[i]

    def is_heading(i):
        if i not in styles:
            styles[i] = style_name(para(i).Style)
        return styles[i] in heading_styles

    def read_text(i):
        if i not in texts:
            texts[i] = para(i).Range.Text if 1 <= i <= count else None
        return texts[i]

    def read_level(i):
        if i not in levels:
            level = para(i).OutlineLevel
            levels[i] = None if level == WD_OUTLINE_BODY else int(level)
        return levels[i]

    found = locate_headings(items, split_texts(doc.Content.Text), read_text, read_level,
                            is_heading)
    if found is None:
        return None
    # The text, level and style were all read to place each heading.
    return [{"index": i, "level": levels[i], "text": clean_text(texts[i]),
             "style": styles[i]} for i in found]


def _get_outline_impl(every_paragraph=False):
    doc = backend.word.document()
    count = doc.Paragraphs.Count
    if not every_paragraph:
        headings = _fast_outline(doc, count)
        if headings is not None:
            return {"success": True, "paragraphs": count, "headings": headings}
    headings = []
    for index, p in paragraphs_in(doc, 1, count):
        level = p.OutlineLevel
        if level != WD_OUTLINE_BODY:
            headings.append({"index": index, "level": level,
                             "text": clean_text(p.Range.Text),
                             "style": style_name(p.Style)})
    return {"success": True, "paragraphs": count, "headings": headings}


class GetParagraphsInput(Input):
    start: int = Field(1, ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description=(
        f"Last paragraph. Leave out for up to {MAX_PARAGRAPHS_PER_READ} from `start`."))
    detail: bool = Field(False, description=(
        "Also give each paragraph's alignment, fonts (one per script), line "
        "spacing and spacing before and after."))


#: How long one word_get_paragraphs reads. A call has 45 s on Windows and 69 s
#: on the Mac; 200 paragraphs with `detail` took 55 s on the Mac (measured,
#: docs/PROBES.md), and paragraphs with more formatting than that take longer.
READ_SECONDS = 40.0


def finish_reading(result: dict, last: int, count: int) -> None:
    """Say what follows the paragraphs read, and if time ran out before `last`."""
    reached = result["paragraphs"][-1]["index"]
    if reached < last:
        result["complete"] = False
        result.setdefault("warnings", []).append(
            f"The time for one call ran out after paragraph {reached}, before paragraph "
            f"{last}. Call again with start={reached + 1} for the rest, or leave out "
            "`detail` to read faster.")
    if reached < count:
        result["more"] = (f"Paragraphs {reached + 1} to {count} follow; "
                          f"pass start={reached + 1}.")


def _get_paragraphs_impl(start: int, end: int | None, detail: bool):
    doc = backend.word.document()
    began = time.monotonic()
    count = doc.Paragraphs.Count
    if start > count:
        raise ValueError(f"The document has {count} paragraphs; start is {start}.")
    last = min(count, end if end is not None else start + MAX_PARAGRAPHS_PER_READ - 1)
    if last - start + 1 > MAX_PARAGRAPHS_PER_READ:
        last = start + MAX_PARAGRAPHS_PER_READ - 1
    from .tables import mark_tables

    out = []
    for i, p in paragraphs_in(doc, start, last):
        out.append(paragraph_info(i, p, detail))
        if time.monotonic() > began + READ_SECONDS:
            break
    mark_tables(doc, out)
    result = {"success": True, "paragraphs": out, "count": count}
    finish_reading(result, last, count)
    return result


# -- writing -------------------------------------------------------------------

class NewParagraph(Input):
    text: str = Field(..., description=(
        "The paragraph's text, without a paragraph mark. A \\n inside becomes a "
        "line break within the paragraph."))
    style: Trimmed | None = Field(None, description=(
        "Paragraph style: a built-in English name such as \"Heading 1\" or "
        "\"Normal\", which works in every language of Word, or the document's "
        "own name for a style. When given, manual character formatting is "
        "cleared so the text looks as the style says."))

    @model_validator(mode="after")
    def _no_paragraph_marks(self):
        if "\r" in self.text:
            raise ValueError("text must not contain \\r; give each paragraph its own item.")
        return self


class InsertParagraphsInput(Input):
    paragraphs: list[NewParagraph] = Field(..., min_length=1, max_length=200)
    after: int | None = Field(None, ge=0, description=(
        "Insert after this paragraph (0 = at the very start)."))
    before: int | None = Field(None, ge=1, description="Insert before this paragraph.")

    @model_validator(mode="after")
    def _one_place(self):
        if self.after is not None and self.before is not None:
            raise ValueError("Give `after` or `before`, not both. Leave both out to "
                             "add at the end.")
        return self


def _starts_table(doc, at: int) -> bool:
    """Whether paragraph `at` is the first of a table (the one before is not in it)."""
    return bool(paragraph(doc, at).Range.Information(WD_WITHIN_TABLE)
                and not paragraph(doc, at - 1).Range.Information(WD_WITHIN_TABLE))


def style_runs(paragraphs: list[dict]) -> list[tuple[str, int, int]]:
    """(style, first offset, last offset) of each run of paragraphs given one style."""
    runs: list[list] = []
    for offset, item in enumerate(paragraphs):
        style = item.get("style")
        if not style:
            continue
        if runs and runs[-1][0] == style and runs[-1][2] == offset - 1:
            runs[-1][2] = offset
        else:
            runs.append([style, offset, offset])
    return [tuple(r) for r in runs]


def _insert_paragraphs_impl(paragraphs: list[dict], after: int | None,
                            before: int | None):
    doc = backend.word.document()
    count = doc.Paragraphs.Count
    texts = [p["text"].replace("\n", "\v") for p in paragraphs]
    if before is not None:
        paragraph(doc, before)
        at = before
    elif after is not None:
        if after > count:
            raise ValueError(f"There is no paragraph {after}: the document has {count}.")
        at = after + 1
    else:
        at = count + 1

    only = paragraph(doc, 1) if count == 1 else None
    if at > count:
        last = paragraph(doc, count)
        if only is not None and clean_text(only.Range.Text) == "":
            # A new document is one empty paragraph: fill it rather than
            # leaving it empty above the new text.
            first_index = 1
            rng = doc.Range(last.Range.Start, last.Range.Start)
            rng.InsertAfter("\r".join(texts))
        else:
            first_index = count + 1
            end = last.Range.End - 1
            rng = doc.Range(end, end)
            rng.InsertAfter("\r" + "\r".join(texts))
    else:
        first_index = at
        target = paragraph(doc, at)
        if at > 1 and (target.Range.OMaths.Count or _starts_table(doc, at)):
            # Text put at the start of a paragraph that holds an equation goes
            # into the equation, which refuses a paragraph mark (measured on
            # Word for Windows; the Mac refuses it too), and at the start of a
            # table's first cell it goes into the cell. The new paragraphs go
            # in at the end of the paragraph before instead, and take its
            # formatting, so each gets a style below.
            end = paragraph(doc, at - 1).Range.End - 1
            doc.Range(end, end).InsertAfter("\r" + "\r".join(texts))
            paragraphs = [{**item, "style": item.get("style") or "Normal"}
                          for item in paragraphs]
        else:
            start = target.Range.Start
            rng = doc.Range(start, start)
            rng.InsertAfter("\r".join(texts) + "\r")

    # A style goes on each run of paragraphs that share it, and the text is
    # read back in one call, rather than several calls a paragraph.
    last_index = first_index + len(texts) - 1
    for style, a, b in style_runs(paragraphs):
        rng = doc.Range(paragraph(doc, first_index + a).Range.Start,
                        paragraph(doc, first_index + b).Range.End)
        rng.Style = resolve_style(doc, style)
        rng.Font.Reset()
    span = doc.Range(paragraph(doc, first_index).Range.Start,
                     paragraph(doc, last_index).Range.End)
    got = (span.Text or "").split("\r")
    looks: dict = {}
    written, warnings = [], []
    for offset, item in enumerate(paragraphs):
        index = first_index + offset
        key = item.get("style")
        if key not in looks:  # the style and level, read once for each style given
            info = paragraph_info(index, paragraph(doc, index))
            looks[key] = {k: info[k] for k in ("style", "outline_level") if k in info}
        now = clean_text(got[offset]) if offset < len(got) else ""
        if now != texts[offset]:
            warnings.append(f"Paragraph {index} reads back as {now!r}, not "
                            f"{texts[offset]!r}.")
        written.append({"index": index, "text": now, **looks[key]})
    if written:
        follow(doc, paragraph(doc, last_index).Range)
    result = {"success": True, "inserted": written, "paragraphs": doc.Paragraphs.Count}
    if warnings:
        result["warnings"] = warnings
    return result


class SetParagraphTextInput(Input):
    index: int = Field(..., ge=1, description="The paragraph to rewrite.")
    text: str = Field(..., description=(
        "Its new text, without a paragraph mark. The paragraph keeps its style "
        "and paragraph formatting; the text takes the formatting of its first "
        "character."))

    @model_validator(mode="after")
    def _no_paragraph_marks(self):
        if "\r" in self.text:
            raise ValueError("text must not contain \\r. Use word_insert_paragraphs to "
                             "add paragraphs.")
        return self


def lost_shapes(count: int) -> list[str]:
    """A warning when floating shapes went with the text they were anchored in."""
    if count <= 0:
        return []
    return [f"{count} floating shape(s) anchored in this text went with it: word_undo brings "
            "them back on Windows; on the Mac insert them again."]


def _set_paragraph_text_impl(index: int, text: str):
    doc = backend.word.document()
    p = paragraph(doc, index)
    rng = p.Range
    old = clean_text(visible_text(rng)[0])
    shapes = doc.Shapes.Count
    hidden = max(0, rng.End - rng.Start - len(rng.Text.encode("utf-16-le")) // 2)
    # A shape anchored at the paragraph's start stays; its anchor is not text.
    kept = leading_anchors(doc, rng.Start, rng.End, hidden)
    target = doc.Range(rng.Start + kept, rng.End - 1)  # a field hides its code from the text
    target.Text = text.replace("\n", "\v")
    p = paragraph(doc, index)
    info = paragraph_info(index, p)
    follow(doc, p.Range)
    result = {"success": True, "paragraph": info, "was": old}
    warnings = lost_shapes(shapes - doc.Shapes.Count)
    if info["text"] != text.replace("\n", "\v"):
        warnings.append(f"It reads back as {info['text']!r}.")
    if warnings:
        result["warnings"] = warnings
    return result


class DeleteParagraphsInput(Input):
    start: int = Field(..., ge=1, description="First paragraph to delete.")
    end: int | None = Field(None, ge=1, description="Last one. Leave out for one paragraph.")


TRACKED_DELETE = ("Change tracking is on, so these are marked as deleted: they stay in the "
                  "document, and the paragraphs keep their numbers, until the change is "
                  "accepted.")


def _delete_paragraphs_impl(start: int, end: int | None):
    doc = backend.word.document()
    start, end = check_span(doc, start, end)
    before = doc.Paragraphs.Count
    shapes = doc.Shapes.Count
    removed = [clean_text(visible_text(p.Range)[0]) for _, p in paragraphs_in(doc, start, end)]
    tracking = bool(doc.TrackRevisions)
    first = paragraph(doc, start).Range.Start
    last = paragraph(doc, end).Range.End
    if end == before and start > 1:
        # The document's last paragraph mark cannot go. Take the mark before
        # the span instead, so the paragraph above becomes the last one.
        first -= 1
    doc.Range(first, last).Delete()
    after = doc.Paragraphs.Count
    result = {"success": True, "deleted": end - start + 1, "texts": removed,
              "paragraphs": after}
    warnings = lost_shapes(shapes - doc.Shapes.Count)
    if tracking:
        result["note"] = TRACKED_DELETE
    elif before - after != end - start + 1:
        warnings.append(
            f"The document went from {before} to {after} paragraphs, not "
            f"{before - (end - start + 1)}. Read it back with word_get_paragraphs.")
    if warnings:
        result["warnings"] = warnings
    if after:
        follow(doc, paragraph(doc, min(start, after)).Range)
    return result


class FindReplaceInput(Input):
    find: str = Field(..., min_length=1, max_length=255, description="Text to find.")
    replace: str | None = Field(None, max_length=255, description=(
        "Replacement. Leave out to only count and locate the matches."))
    match_case: bool = False
    whole_word: bool = False
    wildcards: bool = Field(False, description="Treat `find` as a Word wildcard pattern.")


#: How long one word_find_replace searches, and how long it works in all. A
#: call has 60 s on Windows and 84 s on the Mac before the worker gives it up,
#: and then it cannot say how far it got; stopping sooner, it can.
SEARCH_SECONDS = 20.0
WORK_SECONDS = 45.0
#: Paragraph numbers are read for this many matches, at one call each.
LISTED_MATCHES = 100


def replace_outcome(result: dict, listed: int, searched_all: bool, replaced: int | None,
                    find: str, replace: str | None, wildcards: bool = False) -> None:
    """Add to a find-and-replace answer what the time limit left undone.

    Calling again is offered only when it cannot replace new text a second
    time. A wildcard pattern can match text that does not contain it ("a?"
    matches "abc"), so for one that is left to the caller to judge.
    """
    if result["matches"] > listed:
        result["paragraphs_cover"] = f"the first {listed} matches"
    if searched_all and (replaced is None or replaced == result["matches"]):
        return
    result["complete"] = False
    found = f"{'at least ' if not searched_all else ''}{result['matches']}"
    if replace is None:
        note = (f"Word was still searching when the time for one call ran out: there are "
                f"{found} matches.")
    else:
        note = (f"The time for one call ran out with {replaced} of {found} matches replaced. ")
        if wildcards:
            note += ("A wildcard pattern can match the new text too, and calling again would "
                     "then replace it a second time. Call again only if the pattern cannot "
                     "match the new text; otherwise narrow the search.")
        elif find.lower() in replace.lower():
            note += ("The new text holds what was searched for, so calling again would "
                     "replace those too: narrow the search.")
        else:
            note += "Call again with the same arguments for the rest."
    result.setdefault("warnings", []).append(note)


def _find_all(doc, find, match_case, whole_word, wildcards, deadline):
    """(start, end, paragraph index) of each match, in order, and whether all were found.

    The paragraph index is None after the first LISTED_MATCHES.

    Outside wildcard searches Word reads ^ as the start of a special code
    (^p, ^t, ^2...), so a caret is doubled to find itself: "x^2" finds x^2.
    """
    if not wildcards:
        find = find.replace("^", "^^")
    deleted = deletion_spans(doc.Content)  # Word's search finds these too
    rng = doc.Content
    f = rng.Find
    f.ClearFormatting()
    found, reached = [], -1
    while f.Execute(FindText=find, MatchCase=match_case, MatchWholeWord=whole_word,
                    MatchWildcards=wildcards, Forward=True, Wrap=WD_FIND_STOP,
                    Format=False):
        if rng.End <= reached:  # a search that stops moving would never end
            raise ValueError(f"Word's search stopped moving forward at position {rng.End}, "
                             f"after {len(found)} matches. Nothing was changed. Try a "
                             "narrower search.")
        reached = rng.End
        if not inside_spans(rng.Start, rng.End, deleted):
            found.append((rng.Start, rng.End,
                          index_of(doc, rng) if len(found) < LISTED_MATCHES else None))
        if time.monotonic() > deadline:
            return found, False
        rng.Collapse(WD_COLLAPSE_END)
    return found, True


def _find_replace_impl(find: str, replace: str | None, match_case: bool,
                       whole_word: bool, wildcards: bool):
    doc = backend.word.document()
    began = time.monotonic()
    matches, searched_all = _find_all(doc, find, match_case, whole_word, wildcards,
                                      began + SEARCH_SECONDS)
    listed = matches[:LISTED_MATCHES]
    result = {"success": True, "matches": len(matches),
              "paragraphs": sorted({m[2] for m in listed})}
    if replace is None or not matches:
        replace_outcome(result, len(listed), searched_all, None, find, None)
        return result
    # Each match is replaced through its own range, last first so the
    # positions of the ones before it hold. Word's own ReplaceAll answered
    # success and replaced nothing on Word 16 for Windows, and a range is
    # what both platforms can do; each replacement is read back.
    text = replace.replace("\n", "\v")
    length = len(text.encode("utf-16-le")) // 2  # Word counts UTF-16 units
    done, missed = 0, 0
    for start, end, _ in reversed(matches):
        if time.monotonic() > began + WORK_SECONDS:
            break
        target = doc.Range(start, end)
        target.Text = text
        done += 1
        if (doc.Range(start, start + length).Text or "") != text:
            missed += 1
    result["replaced"] = done - missed
    if missed:
        result["warnings"] = [f"{missed} of {done} replacements read back "
                              "differently. Read the paragraphs back."]
    replace_outcome(result, len(listed), searched_all, done, find, replace, wildcards)
    follow(doc, paragraph(doc, matches[0][2]).Range)
    return result


def register(mcp) -> None:
    @mcp.tool(name="word_get_document_info",
              annotations=annotations("Document info", READ_ONLY))
    async def word_get_document_info() -> str:
        """The document's size (paragraphs, pages, words), page setup, and
        whether it is saved and tracking changes."""
        return await call(_get_document_info_impl)

    @mcp.tool(name="word_get_outline", annotations=annotations("Outline", READ_ONLY))
    async def word_get_outline(params: OutlineInput) -> str:
        """The document's headings, with their level and paragraph index.

        Start here to find your way around a document; then read the parts
        you need with word_get_paragraphs. Headings are the paragraphs Word
        lists as headings (heading styles); `every_paragraph` also finds
        paragraphs given an outline level directly.
        """
        return await call(_get_outline_impl, params.every_paragraph)

    @mcp.tool(name="word_get_paragraphs", annotations=annotations("Read paragraphs", READ_ONLY))
    async def word_get_paragraphs(params: GetParagraphsInput) -> str:
        """Read paragraphs by index: text, style and outline level, and with
        `detail` their formatting too. A paragraph in a table carries the
        table's number (`table`): read those with word_get_table. Text marked
        as a tracked deletion is left out of `text` and given in
        `deleted_text`. A read that runs out of time answers what it has,
        `complete` false and where to go on in `more`."""
        return await call(_get_paragraphs_impl, params.start, params.end,
                                   params.detail)

    @mcp.tool(name="word_insert_paragraphs",
              annotations=annotations("Insert paragraphs", WRITES))
    async def word_insert_paragraphs(params: InsertParagraphsInput) -> str:
        """Insert paragraphs, each with its own style, after or before a paragraph
        or at the end.

        Writes through a range, so the user's cursor stays where it is. The
        result reads each new paragraph back with its index and style.
        """
        return await call(
            _insert_paragraphs_impl,
            [p.model_dump() for p in params.paragraphs], params.after, params.before,
            undo="Insert paragraphs")

    @mcp.tool(name="word_set_paragraph_text",
              annotations=annotations("Rewrite a paragraph", WRITES))
    async def word_set_paragraph_text(params: SetParagraphTextInput) -> str:
        """Replace one paragraph's text, keeping its style and formatting."""
        return await call(_set_paragraph_text_impl, params.index, params.text,
                                   undo="Rewrite paragraph")

    @mcp.tool(name="word_delete_paragraphs",
              annotations=annotations("Delete paragraphs", DESTRUCTIVE))
    async def word_delete_paragraphs(params: DeleteParagraphsInput) -> str:
        """Delete a paragraph or a run of them, and say what was removed."""
        return await call(_delete_paragraphs_impl, params.start, params.end,
                                   undo="Delete paragraphs")

    @mcp.tool(name="word_find_replace", annotations=annotations("Find and replace", WRITES))
    async def word_find_replace(params: FindReplaceInput) -> str:
        """Find text in the whole document, and replace it if `replace` is given.

        Answers how many matched, in which paragraphs, and how many were
        replaced. Very many matches can take longer than one call may run:
        then the answer has `complete` false and a warning saying how far it
        got and what to do next.
        """
        return await call(
            _find_replace_impl, params.find, params.replace, params.match_case,
            params.whole_word, params.wildcards,
            undo=None if params.replace is None else "Replace")
