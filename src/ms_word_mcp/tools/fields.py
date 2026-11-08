"""Fields and tables of contents: insert them, list them, update them.

A table of contents is a TOC field, made the way Word's own command makes it
(\\o "1-3" \\h \\z \\u); one for part of the document only takes \\b with a
bookmark over that part. Fields are added as wdFieldEmpty (-1) with their
whole code, which is how a user's Ctrl+F9 field is made
(https://learn.microsoft.com/en-us/office/vba/api/word.wdfieldtype).
"""

from __future__ import annotations

import re
from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, Input, annotations, call, follow, mac_impl
from ._word import check_span, clean_text, paragraph, span_range

WD_FIELD_EMPTY = -1


def toc_code(upper: int, lower: int, hyperlinks: bool, page_numbers: bool,
             bookmark: str | None) -> str:
    code = f'TOC \\o "{upper}-{lower}"'
    if hyperlinks:
        code += " \\h"
    code += " \\z \\u"
    if not page_numbers:
        code += " \\n"
    if bookmark:
        code += f" \\b {bookmark}"
    return code


class InsertTocInput(Input):
    after: int | None = Field(None, ge=0, description=(
        "Put the table after this paragraph (0 = at the very start)."))
    before: int | None = Field(None, ge=1, description="Put it before this paragraph.")
    upper_level: int = Field(1, ge=1, le=9, description="Highest heading level listed.")
    lower_level: int = Field(3, ge=1, le=9, description="Lowest heading level listed.")
    hyperlinks: bool = Field(True, description="Entries link to their headings.")
    page_numbers: bool = Field(True, description="Show page numbers.")
    from_paragraph: int | None = Field(None, ge=1, description=(
        "List only the headings from this paragraph..."))
    to_paragraph: int | None = Field(None, ge=1, description=(
        "...to this one: a table of contents for one part, such as a chapter."))

    @model_validator(mode="after")
    def _consistent(self):
        if self.after is not None and self.before is not None:
            raise ValueError("Give `after` or `before`, not both.")
        if self.lower_level < self.upper_level:
            raise ValueError("lower_level must be at least upper_level.")
        if (self.from_paragraph is None) != (self.to_paragraph is None):
            raise ValueError("Give both from_paragraph and to_paragraph for a part, or "
                             "neither for the whole document.")
        if self.from_paragraph and self.to_paragraph < self.from_paragraph:
            raise ValueError("to_paragraph comes before from_paragraph.")
        return self


def _toc_summary(doc, field) -> dict:
    result = field.Result
    lines = [clean_text(line) for line in (result.Text or "").split("\r") if line.strip()]
    return {"code": field.Code.Text.strip(), "entries": len(lines),
            "first_entries": [re.sub(r"\s+", " ", x)[:80] for x in lines[:8]]}


def _insert_toc_impl(after, before, upper_level, lower_level, hyperlinks, page_numbers,
                     from_paragraph, to_paragraph):
    from . import text

    doc = backend.word.document()
    bookmark = None
    if from_paragraph is not None:
        check_span(doc, from_paragraph, to_paragraph)
    written = text._insert_paragraphs_impl([{"text": "", "style": None}], after, before)
    index = written["inserted"][0]["index"]
    if from_paragraph is not None:
        # The part's paragraphs moved down by one if the table went above them.
        shift = 1 if index <= from_paragraph else 0
        span = span_range(doc, from_paragraph + shift, to_paragraph + shift)
        bookmark = f"_McpToc{doc.Bookmarks.Count + 1}"
        doc.Bookmarks.Add(bookmark, span)
    p = paragraph(doc, index)
    at = doc.Range(p.Range.Start, p.Range.Start)
    field = doc.Fields.Add(at, WD_FIELD_EMPTY,
                           toc_code(upper_level, lower_level, hyperlinks, page_numbers,
                                    bookmark), False)
    field.Update()
    summary = _toc_summary(doc, field)
    follow(doc, field.Result)
    result = {"success": True, "index": index, **summary}
    if bookmark:
        result["bookmark"] = bookmark
    if summary["entries"] == 0:
        result["warnings"] = ["The table lists no headings: give headings a heading style "
                              "(word_apply_style \"Heading 1\") and update it."]
    return result


class UpdateFieldsInput(Input):
    what: Literal["all", "tables_of_contents", "page_numbers", "numbers"] = Field(
        "all", description=(
            "all: every field in the main text. tables_of_contents: rebuild the tables "
            "of contents (entries and page numbers). page_numbers: only the page numbers "
            "of the tables of contents, leaving their entries. numbers: SEQ fields "
            "(equation, figure and table numbers)."))
    start: int | None = Field(None, ge=1, description=(
        "Only the fields in these paragraphs: first paragraph..."))
    end: int | None = Field(None, ge=1, description="...last paragraph.")

    @model_validator(mode="after")
    def _span(self):
        if self.end is not None and self.start is None:
            raise ValueError("Give start with end.")
        return self


def _fields_in(doc, start, end):
    fields = doc.Fields
    if start is not None:
        s, e = check_span(doc, start, end)
        fields = span_range(doc, s, e).Fields
    return [fields.Item(i) for i in range(1, fields.Count + 1)]


def _kind(code: str) -> str:
    word = (code or "").strip().split(" ")[0].upper()
    return word or "?"


def _update_fields_impl(what, start, end):
    doc = backend.word.document()
    updated, failed = [], []
    if what in ("tables_of_contents", "page_numbers") and start is None:
        tocs = doc.TablesOfContents
        for i in range(1, tocs.Count + 1):
            toc = tocs.Item(i)
            try:
                if what == "page_numbers":
                    toc.UpdatePageNumbers()
                else:
                    toc.Update()
                updated.append(f"table of contents {i}")
            except Exception as exc:  # noqa: BLE001 - reported
                failed.append(f"table of contents {i}: {exc}")
    else:
        for i, f in enumerate(_fields_in(doc, start, end), start=1):
            kind = _kind(f.Code.Text)
            if what == "numbers" and kind != "SEQ":
                continue
            if what in ("tables_of_contents", "page_numbers") and kind != "TOC":
                continue
            # Field.Update answers True when the update worked (the page's own
            # example tests for it).
            ok = f.Update()
            (updated if ok in (True, 1, -1, None) else failed).append(f"{kind} field {i}")
    result = {"success": True, "updated": len(updated), "which": updated[:50]}
    if failed:
        result["warnings"] = [f"Not updated: {', '.join(failed[:20])}."]
    if not updated and not failed:
        result["note"] = "There was nothing of that kind to update."
    return result


class ListFieldsInput(Input):
    start: int | None = Field(None, ge=1, description="First paragraph (optional).")
    end: int | None = Field(None, ge=1, description="Last paragraph (optional).")


def _list_fields_impl(start, end):
    doc = backend.word.document()
    out = []
    for f in _fields_in(doc, start, end):
        code = f.Code.Text.strip()
        out.append({"index": doc.Range(0, f.Code.Start + 1).Paragraphs.Count,
                    "kind": _kind(code), "code": code,
                    "result": clean_text(f.Result.Text or "")[:120]})
    return {"success": True, "fields": out}


class InsertFieldInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the field goes in.")
    code: str = Field(..., min_length=1, max_length=255, description=(
        "The field code without braces: PAGE, NUMPAGES, DATE \\@ \"yyyy/M/d\", "
        "STYLEREF \"Heading 1\" \\n, SEQ Figure, REF bookmark \\h ..."))
    position: Literal["end", "start"] = Field("end", description=(
        "At the end of the paragraph's text or at its start."))
    match: str | None = Field(None, min_length=1, max_length=255, description=(
        "Instead of `position`: replace the first occurrence of this text in the "
        "paragraph with the field."))


def _insert_field_impl(index, code, position, match):
    doc = backend.word.document()
    p = paragraph(doc, index)
    rng = p.Range
    text = rng.Text
    if match is not None:
        at = text.find(match)
        if at < 0:
            raise ValueError(f"'{match}' is not in paragraph {index}. Nothing was inserted.")
        s = rng.Start + len(text[:at].encode("utf-16-le")) // 2
        target = doc.Range(s, s + len(match.encode("utf-16-le")) // 2)
    elif position == "start":
        target = doc.Range(rng.Start, rng.Start)
    else:
        target = doc.Range(rng.End - 1, rng.End - 1)
    field = doc.Fields.Add(target, WD_FIELD_EMPTY, code, False)
    field.Update()
    follow(doc, field.Result)
    return {"success": True, "index": index, "code": field.Code.Text.strip(),
            "result": clean_text(field.Result.Text or "")}


# -- Mac -------------------------------------------------------------------------------

def _mac_fields(doc):
    from appscript import k

    return [doc.fields[i] for i in range(1, int(doc.count(each=k.field)) + 1)]


def _mac_new_field(app, doc, start, end, code):
    """Make a field over [start, end) and find it again by its position."""
    from appscript import k

    app.create_new_field(text_range=doc.create_range(start=start, end_=end),
                         field_type=k.field_empty, field_text=code,
                         preserve_formatting=False)
    for f in _mac_fields(doc):
        s = f.field_code.start_of_content.get()
        if start - 1 <= s <= start + 2:
            return f
    raise RuntimeError("Word answered without error, but the field is not where it was put.")


@mac_impl(_insert_toc_impl)
def _mac_insert_toc(after, before, upper_level, lower_level, hyperlinks, page_numbers,
                    from_paragraph, to_paragraph):
    from appscript import k

    from . import mac as m
    from . import text
    from ._common import MAC_IMPLS

    word = backend.word
    doc = word.document()
    if from_paragraph is not None:
        m._span(doc, from_paragraph, to_paragraph)
    written = MAC_IMPLS[text._insert_paragraphs_impl.__name__](
        [{"text": "", "style": None}], after, before)
    index = written["inserted"][0]["index"]
    bookmark = None
    if from_paragraph is not None:
        shift = 1 if index <= from_paragraph else 0
        s, _ = m._bounds(doc, from_paragraph + shift)
        _, e = m._bounds(doc, to_paragraph + shift)
        bookmark = f"_McpToc{int(doc.count(each=k.bookmark)) + 1}"
        doc.make(new=k.bookmark, at=doc, with_properties={
            k.name: bookmark, k.text_object: doc.create_range(start=s, end_=e)})
    start, _ = m._bounds(doc, index)
    f = _mac_new_field(word.get_app(), doc, start, start,
                       toc_code(upper_level, lower_level, hyperlinks, page_numbers, bookmark))
    f.update_field()
    lines = [m._clean(x) for x in (f.result_range.content.get() or "").split("\r")
             if x.strip()]
    result = {"success": True, "index": index,
              "code": (f.field_code.content.get() or "").strip(), "entries": len(lines),
              "first_entries": [re.sub(r"\s+", " ", x)[:80] for x in lines[:8]]}
    if bookmark:
        result["bookmark"] = bookmark
    if not lines:
        result["warnings"] = ["The table lists no headings: give headings a heading style "
                              "(word_apply_style \"Heading 1\") and update it."]
    return result


@mac_impl(_update_fields_impl)
def _mac_update_fields(what, start, end):
    from . import mac as m

    doc = backend.word.document()
    lo, hi = None, None
    if start is not None:
        s, e = m._span(doc, start, end)
        lo, hi = m._bounds(doc, s)[0], m._bounds(doc, e)[1]
    updated, failed = [], []
    for i, f in enumerate(_mac_fields(doc), start=1):
        kind = _kind(f.field_code.content.get())
        if lo is not None and not lo <= f.field_code.start_of_content.get() < hi:
            continue
        if what == "numbers" and kind != "SEQ":
            continue
        if what in ("tables_of_contents", "page_numbers") and kind != "TOC":
            continue
        try:
            f.update_field()
            updated.append(f"{kind} field {i}")
        except Exception as exc:  # noqa: BLE001 - reported
            failed.append(f"{kind} field {i}: {str(exc).splitlines()[0]}")
    result = {"success": True, "updated": len(updated), "which": updated[:50]}
    if what == "page_numbers" and updated:
        result["note"] = ("On macOS a table of contents is updated whole: its entries "
                          "are rebuilt along with its page numbers.")
    if failed:
        result["warnings"] = [f"Not updated: {', '.join(failed[:20])}."]
    if not updated and not failed:
        result["note"] = "There was nothing of that kind to update."
    return result


@mac_impl(_list_fields_impl)
def _mac_list_fields(start, end):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    lo, hi = None, None
    if start is not None:
        s, e = m._span(doc, start, end)
        lo, hi = m._bounds(doc, s)[0], m._bounds(doc, e)[1]
    out = []
    for f in _mac_fields(doc):
        at = f.field_code.start_of_content.get()
        if lo is not None and not lo <= at < hi:
            continue
        code = (f.field_code.content.get() or "").strip()
        out.append({"index": int(doc.create_range(start=0, end_=at + 1)
                                 .count(each=k.paragraph)),
                    "kind": _kind(code), "code": code,
                    "result": m._clean(f.result_range.content.get() or "")[:120]})
    return {"success": True, "fields": out}


@mac_impl(_insert_field_impl)
def _mac_insert_field(index, code, position, match):
    from . import mac as m

    word = backend.word
    doc = word.document()
    m._para(doc, index)
    start, end = m._bounds(doc, index)
    if match is not None:
        content = doc.paragraphs[index].text_object.content.get()
        at = content.find(match)
        if at < 0:
            raise ValueError(f"'{match}' is not in paragraph {index}. Nothing was inserted.")
        s = start + m._u16(content, at)
        e = s + m._u16(match, len(match))
    elif position == "start":
        s = e = start
    else:
        s = e = end - 1
    f = _mac_new_field(word.get_app(), doc, s, e, code)
    f.update_field()
    return {"success": True, "index": index, "code": (f.field_code.content.get() or "").strip(),
            "result": m._clean(f.result_range.content.get() or "")}


def register(mcp) -> None:
    @mcp.tool(name="word_insert_table_of_contents",
              annotations=annotations("Insert a table of contents", WRITES))
    async def word_insert_table_of_contents(params: InsertTocInput) -> str:
        """Insert a table of contents built from the heading styles, as Word's
        own command makes it, for the whole document or for part of it (a
        chapter: from_paragraph to to_paragraph). Reads back its entries."""
        p = params
        return await call(_insert_toc_impl, p.after, p.before, p.upper_level, p.lower_level,
                          p.hyperlinks, p.page_numbers, p.from_paragraph, p.to_paragraph,
                          undo="Insert table of contents")

    @mcp.tool(name="word_update_fields", annotations=annotations("Update fields", WRITES))
    async def word_update_fields(params: UpdateFieldsInput) -> str:
        """Update fields: all of them, the tables of contents (whole, or their
        page numbers only), or the SEQ numbers of equations, figures and tables;
        in the whole document or in some paragraphs."""
        p = params
        return await call(_update_fields_impl, p.what, p.start, p.end, undo="Update fields")

    @mcp.tool(name="word_list_fields", annotations=annotations("List fields", READ_ONLY))
    async def word_list_fields(params: ListFieldsInput) -> str:
        """The fields in the document or some paragraphs: kind, code and what
        each shows."""
        p = params
        return await call(_list_fields_impl, p.start, p.end)

    @mcp.tool(name="word_insert_field", annotations=annotations("Insert a field", WRITES))
    async def word_insert_field(params: InsertFieldInput) -> str:
        """Insert a field by its code (PAGE, DATE, STYLEREF, SEQ, REF...) at the
        start or end of a paragraph, or in place of some of its text, and read
        what it shows."""
        p = params
        return await call(_insert_field_impl, p.index, p.code, p.position, p.match,
                          undo="Insert field")
