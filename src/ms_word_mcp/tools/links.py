"""Hyperlinks and the user's own bookmarks.

A hyperlink is a HYPERLINK field over some text, to an address or to a
bookmark in the document.
https://learn.microsoft.com/en-us/office/vba/api/word.hyperlinks.add
https://learn.microsoft.com/en-us/office/vba/api/word.bookmarks.add

Bookmarks whose names start with an underscore are hidden: Word keeps them for
cross-references and tables of contents. They are left out here by name, since
Word itself lists them when the user has turned on Show Hidden Bookmarks.

Word keeps a link's address once it has one: setting it to "" leaves it as it
was, on Windows and the Mac (measured). So a link to an address that is to go
to a bookmark is taken off and made again over the same text. On Word for Mac,
counting the links in a range answers 0 (measured).
"""

from __future__ import annotations

import re

from pydantic import Field, field_validator, model_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, Trimmed, annotations, call, follow, mac_impl
from ._word import clean_text, locate_spans, paragraph

NAME = re.compile(r"[^\W\d_]\w{0,39}")
SHOWN = 120


def _bookmark_name(name: str | None) -> str | None:
    if name is not None and NAME.fullmatch(name) is None:
        raise ValueError("A bookmark's name starts with a letter and holds only letters, "
                         "digits and underscores, at most 40 of them.")
    return name


class InsertHyperlinkInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the link goes in.")
    address: Trimmed | None = Field(None, min_length=1, max_length=2000, description=(
        "Where it goes: a full address (https://..., mailto:...) or a file path."))
    bookmark: Trimmed | None = Field(None, description=(
        "Or: the name of a bookmark in this document."))
    match: str | None = Field(None, min_length=1, max_length=255, description=(
        "Make the first occurrence of this text in the paragraph the link."))
    text: str | None = Field(None, min_length=1, max_length=2000, description=(
        "Or: add this text at the end of the paragraph, as the link. "
        "With neither, the paragraph's whole text becomes the link."))
    screen_tip: str | None = Field(None, min_length=1, max_length=255, description=(
        "Shown when the pointer is over the link."))

    _name = field_validator("bookmark")(_bookmark_name)

    @model_validator(mode="after")
    def _one_each(self):
        if (self.address is None) == (self.bookmark is None):
            raise ValueError("Pass either `address` or `bookmark`.")
        if self.match is not None and self.text is not None:
            raise ValueError("Pass `match` or `text`, not both.")
        return self


class ListHyperlinksInput(Input):
    start: int = Field(1, ge=1, description="The link to start from (numbered from 1).")
    limit: int = Field(100, ge=1, le=200, description="How many to read at most.")


class SetHyperlinkInput(Input):
    number: int = Field(..., ge=1, description="Which link, as word_list_hyperlinks numbers them.")
    address: Trimmed | None = Field(None, min_length=1, max_length=2000, description="A new address.")
    bookmark: Trimmed | None = Field(None, description=(
        "Or: a bookmark in this document to go to instead."))
    text: str | None = Field(None, min_length=1, max_length=2000, description="New visible text.")
    screen_tip: str | None = Field(None, max_length=255, description=(
        "A new screen tip; an empty string takes it off."))

    _name = field_validator("bookmark")(_bookmark_name)

    @model_validator(mode="after")
    def _something(self):
        if self.address is not None and self.bookmark is not None:
            raise ValueError("Pass `address` or `bookmark`, not both.")
        if all(v is None for v in (self.address, self.bookmark, self.text, self.screen_tip)):
            raise ValueError("Nothing to change: pass address, bookmark, text or screen_tip.")
        return self


class RemoveHyperlinkInput(Input):
    number: int = Field(..., ge=1, description="Which link, as word_list_hyperlinks numbers them.")


class AddBookmarkInput(Input):
    name: Trimmed = Field(..., description=(
        "Starts with a letter; letters, digits and underscores, at most 40. "
        "A bookmark already of this name moves here."))
    index: int = Field(..., ge=1, description="The paragraph it marks.")
    match: str | None = Field(None, min_length=1, max_length=255, description=(
        "Mark the first occurrence of this text in the paragraph. "
        "Left out: the paragraph's whole text."))

    _name = field_validator("name")(_bookmark_name)


class DeleteBookmarkInput(Input):
    name: Trimmed = Field(..., min_length=1, description="The bookmark's name.")


def _u16(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


def check_link(number: int, count: int) -> None:
    if not 1 <= number <= count:
        raise ValueError(f"There is no hyperlink {number}: the document has {count} "
                         "(numbered from 1; word_list_hyperlinks lists them). Nothing was changed.")


def no_bookmark(name: str) -> ValueError:
    return ValueError(f"There is no bookmark named '{name}' (word_list_bookmarks lists them). "
                      "Nothing was changed.")


def not_in_paragraph(match: str, index: int) -> ValueError:
    return ValueError(f"'{match}' is not in paragraph {index}. Nothing was changed.")


def nothing_to_link(index: int) -> ValueError:
    return ValueError(f"Paragraph {index} has no text to make a link of: pass `text`.")


def link_listing(out: list[dict], count: int, start: int) -> dict:
    result = {"success": True, "count": count, "hyperlinks": out}
    last = start + len(out) - 1
    if out and last < count:
        result["more"] = f"Links {last + 1} to {count} follow; pass start={last + 1}."
    return result


def bookmark_answer(info: dict, moved: bool) -> dict:
    result = {"success": True, **info}
    if moved:
        result["note"] = "A bookmark of this name was elsewhere; it has moved here."
    return result


def _span(doc, index, match):
    """Word's (start, end) of `match` in the paragraph, or of its whole text."""
    rng = paragraph(doc, index).Range
    if match is None:
        return rng.Start, rng.End - 1
    raw = rng.Text
    if raw.find(match) < 0:
        raise not_in_paragraph(match, index)
    hidden = max(0, rng.End - rng.Start - _u16(raw))
    return locate_spans(raw, [match], lambda a, b: doc.Range(a, b).Text, rng.Start, hidden)[0]


def _describe_link(doc, link, number) -> dict:
    return {"number": number, "index": doc.Range(0, link.Range.Start + 1).Paragraphs.Count,
            "text": link.TextToDisplay, "address": link.Address or None,
            "bookmark": link.SubAddress or None, "screen_tip": link.ScreenTip or None}


def _link(doc, number):
    check_link(number, doc.Hyperlinks.Count)
    return doc.Hyperlinks.Item(number)


def _insert_hyperlink_impl(index, address, bookmark, match, text, screen_tip):
    doc = backend.word.document()
    if bookmark is not None and not doc.Bookmarks.Exists(bookmark):
        raise no_bookmark(bookmark)
    start, end = _span(doc, index, match)
    if text is not None:
        doc.Range(end, end).InsertAfter(text)
        start, end = end, end + _u16(text)
    if start == end:
        raise nothing_to_link(index)
    link = doc.Hyperlinks.Add(doc.Range(start, end), address or "", bookmark or "")
    if screen_tip is not None:
        link.ScreenTip = screen_tip
    follow(doc, link.Range)
    links = doc.Hyperlinks
    number = next(i for i in range(1, links.Count + 1) if links.Item(i).Range.Start == start)
    return {"success": True, **_describe_link(doc, links.Item(number), number)}


def _list_hyperlinks_impl(start, limit):
    doc = backend.word.document()
    links = doc.Hyperlinks
    out = [_describe_link(doc, links.Item(i), i)
           for i in range(start, min(links.Count, start + limit - 1) + 1)]
    return link_listing(out, links.Count, start)


def _remake_to_bookmark(doc, link, bookmark) -> None:
    """Take a link to an address off, and make one to `bookmark` over its text."""
    start, text, tip = link.Range.Start, link.TextToDisplay, link.ScreenTip
    link.Delete()
    made = doc.Hyperlinks.Add(doc.Range(start, start + _u16(text)), "", bookmark)
    if tip:
        made.ScreenTip = tip


def _set_hyperlink_impl(number, address, bookmark, text, screen_tip):
    doc = backend.word.document()
    link = _link(doc, number)
    if bookmark is not None and not doc.Bookmarks.Exists(bookmark):
        raise no_bookmark(bookmark)
    if address is not None:
        link.Address, link.SubAddress = address, ""
    if bookmark is not None:
        if link.Address:
            _remake_to_bookmark(doc, link, bookmark)
            link = doc.Hyperlinks.Item(number)
        else:
            link.SubAddress = bookmark
    if text is not None:
        link.TextToDisplay = text
    if screen_tip is not None:
        link.ScreenTip = screen_tip
    return {"success": True, **_describe_link(doc, doc.Hyperlinks.Item(number), number)}


def _remove_hyperlink_impl(number):
    doc = backend.word.document()
    link = _link(doc, number)
    text = link.TextToDisplay
    link.Delete()
    return {"success": True, "removed": number, "text_kept": text, "left": doc.Hyperlinks.Count}


def _shown_bookmarks(doc) -> list:
    marks = doc.Bookmarks
    found = (marks.Item(i) for i in range(1, marks.Count + 1))
    return [b for b in found if not b.Name.startswith("_")]


def _describe_bookmark(doc, mark) -> dict:
    return {"name": mark.Name, "index": doc.Range(0, mark.Start + 1).Paragraphs.Count,
            "text": clean_text(mark.Range.Text)[:SHOWN], "empty": bool(mark.Empty)}


def _add_bookmark_impl(name, index, match):
    doc = backend.word.document()
    moved = bool(doc.Bookmarks.Exists(name))
    start, end = _span(doc, index, match)
    doc.Bookmarks.Add(name, doc.Range(start, end))
    return bookmark_answer(_describe_bookmark(doc, doc.Bookmarks.Item(name)), moved)


def _list_bookmarks_impl():
    doc = backend.word.document()
    found = sorted(_shown_bookmarks(doc), key=lambda b: b.Start)
    return {"success": True, "bookmarks": [_describe_bookmark(doc, b) for b in found]}


def _delete_bookmark_impl(name):
    doc = backend.word.document()
    if not doc.Bookmarks.Exists(name):
        raise no_bookmark(name)
    doc.Bookmarks.Item(name).Delete()
    return {"success": True, "deleted": name, "left": len(_shown_bookmarks(doc))}


# -- Mac -------------------------------------------------------------------------------

def _mac_value(value):
    return value if isinstance(value, str) and value else None


def _mac_span(doc, index, match):
    from . import mac as m

    m._para(doc, index)
    start, end = m._bounds(doc, index)
    if match is None:
        return start, end - 1
    raw = doc.paragraphs[index].text_object.content.get()
    if raw.find(match) < 0:
        raise not_in_paragraph(match, index)

    def read(a, b):
        try:
            return doc.create_range(start=a, end_=b).content.get()
        except Exception:  # noqa: BLE001 - half a character: not here
            return None

    hidden = max(0, end - start - m._u16(raw, len(raw)))
    return locate_spans(raw, [match], read, start, hidden)[0]


def _mac_link_count(doc) -> int:
    from appscript import k

    return int(doc.count(each=k.hyperlink_object))


def _mac_link(doc, number):
    check_link(number, _mac_link_count(doc))
    return doc.hyperlink_objects[number]


def _mac_describe_link(doc, link, number) -> dict:
    from appscript import k

    at = link.text_object.start_of_content.get()
    return {"number": number,
            "index": int(doc.create_range(start=0, end_=at + 1).count(each=k.paragraph)),
            "text": link.text_to_display.get(),
            "address": _mac_value(link.hyperlink_address.get()),
            "bookmark": _mac_value(link.sub_address.get()),
            "screen_tip": _mac_value(link.screen_tip.get())}


def _mac_has_bookmark(doc, name) -> bool:
    return bool(doc.bookmarks[name].exists())


@mac_impl(_insert_hyperlink_impl)
def _mac_insert_hyperlink(index, address, bookmark, match, text, screen_tip):
    from appscript import k

    doc = backend.word.document()
    if bookmark is not None and not _mac_has_bookmark(doc, bookmark):
        raise no_bookmark(bookmark)
    start, end = _mac_span(doc, index, match)
    if text is not None:
        doc.create_range(start=end, end_=end).content.set(text)
        start, end = end, end + _u16(text)
    if start == end:
        raise nothing_to_link(index)
    props = {k.text_object: doc.create_range(start=start, end_=end)}
    if address is not None:
        props[k.hyperlink_address] = address
    if bookmark is not None:
        props[k.sub_address] = bookmark
    if screen_tip is not None:
        props[k.screen_tip] = screen_tip
    doc.make(new=k.hyperlink_object, at=doc, with_properties=props)
    for i in range(1, _mac_link_count(doc) + 1):
        if doc.hyperlink_objects[i].text_object.start_of_content.get() == start:
            return {"success": True, **_mac_describe_link(doc, doc.hyperlink_objects[i], i)}
    raise RuntimeError("Word answered without error, but the link is not where it was put.")


@mac_impl(_list_hyperlinks_impl)
def _mac_list_hyperlinks(start, limit):
    doc = backend.word.document()
    count = _mac_link_count(doc)
    out = [_mac_describe_link(doc, doc.hyperlink_objects[i], i)
           for i in range(start, min(count, start + limit - 1) + 1)]
    return link_listing(out, count, start)


def _mac_remake_to_bookmark(doc, link, bookmark) -> None:
    from appscript import k

    start = link.text_object.start_of_content.get()
    text, tip = link.text_to_display.get(), _mac_value(link.screen_tip.get())
    link.delete()
    props = {k.text_object: doc.create_range(start=start, end_=start + _u16(text)),
             k.sub_address: bookmark}
    if tip:
        props[k.screen_tip] = tip
    doc.make(new=k.hyperlink_object, at=doc, with_properties=props)


@mac_impl(_set_hyperlink_impl)
def _mac_set_hyperlink(number, address, bookmark, text, screen_tip):
    doc = backend.word.document()
    link = _mac_link(doc, number)
    if bookmark is not None:
        if not _mac_has_bookmark(doc, bookmark):
            raise no_bookmark(bookmark)
        if _mac_value(link.hyperlink_address.get()) is not None:
            _mac_remake_to_bookmark(doc, link, bookmark)
            link = doc.hyperlink_objects[number]
        else:
            link.sub_address.set(bookmark)
    if address is not None:
        link.hyperlink_address.set(address)
        link.sub_address.set("")
    if text is not None:
        link.text_to_display.set(text)
    if screen_tip is not None:
        link.screen_tip.set(screen_tip)
    return {"success": True, **_mac_describe_link(doc, doc.hyperlink_objects[number], number)}


@mac_impl(_remove_hyperlink_impl)
def _mac_remove_hyperlink(number):
    doc = backend.word.document()
    link = _mac_link(doc, number)
    text = link.text_to_display.get()
    link.delete()
    return {"success": True, "removed": number, "text_kept": text, "left": _mac_link_count(doc)}


def _mac_shown_bookmarks(doc) -> list:
    from appscript import k

    found = [doc.bookmarks[i] for i in range(1, int(doc.count(each=k.bookmark)) + 1)]
    return [b for b in found if not b.name.get().startswith("_")]


def _mac_describe_bookmark(doc, mark) -> dict:
    from appscript import k

    from . import mac as m

    at = mark.start_of_bookmark.get()
    return {"name": mark.name.get(),
            "index": int(doc.create_range(start=0, end_=at + 1).count(each=k.paragraph)),
            "text": m._clean(mark.text_object.content.get())[:SHOWN],
            "empty": bool(mark.empty.get())}


@mac_impl(_add_bookmark_impl)
def _mac_add_bookmark(name, index, match):
    from appscript import k

    doc = backend.word.document()
    moved = _mac_has_bookmark(doc, name)
    start, end = _mac_span(doc, index, match)
    doc.make(new=k.bookmark, at=doc, with_properties={
        k.name: name, k.text_object: doc.create_range(start=start, end_=end)})
    return bookmark_answer(_mac_describe_bookmark(doc, doc.bookmarks[name]), moved)


@mac_impl(_list_bookmarks_impl)
def _mac_list_bookmarks():
    doc = backend.word.document()
    found = sorted(_mac_shown_bookmarks(doc), key=lambda b: b.start_of_bookmark.get())
    return {"success": True, "bookmarks": [_mac_describe_bookmark(doc, b) for b in found]}


@mac_impl(_delete_bookmark_impl)
def _mac_delete_bookmark(name):
    doc = backend.word.document()
    if not _mac_has_bookmark(doc, name):
        raise no_bookmark(name)
    doc.bookmarks[name].delete()
    return {"success": True, "deleted": name, "left": len(_mac_shown_bookmarks(doc))}


def register(mcp) -> None:
    @mcp.tool(name="word_insert_hyperlink", annotations=annotations("Insert a hyperlink", WRITES))
    async def word_insert_hyperlink(params: InsertHyperlinkInput) -> str:
        """Make a link to an address or to a bookmark in this document: over
        some text of a paragraph, over its whole text, or as new text at its
        end. Word applies its Hyperlink style."""
        p = params
        return await call(_insert_hyperlink_impl, p.index, p.address, p.bookmark, p.match, p.text,
                          p.screen_tip, undo="Insert hyperlink")

    @mcp.tool(name="word_list_hyperlinks", annotations=annotations("List hyperlinks", READ_ONLY))
    async def word_list_hyperlinks(params: ListHyperlinksInput) -> str:
        """The document's links in order: visible text, paragraph, address or
        bookmark, screen tip. The entries of a table of contents are links too."""
        p = params
        return await call(_list_hyperlinks_impl, p.start, p.limit)

    @mcp.tool(name="word_set_hyperlink", annotations=annotations("Change a hyperlink", WRITES))
    async def word_set_hyperlink(params: SetHyperlinkInput) -> str:
        """Change where a link goes, its visible text or its screen tip."""
        p = params
        return await call(_set_hyperlink_impl, p.number, p.address, p.bookmark, p.text,
                          p.screen_tip, undo="Change hyperlink")

    @mcp.tool(name="word_remove_hyperlink", annotations=annotations("Remove a hyperlink", DESTRUCTIVE))
    async def word_remove_hyperlink(params: RemoveHyperlinkInput) -> str:
        """Take a link off. Its text stays in the document."""
        return await call(_remove_hyperlink_impl, params.number, undo="Remove hyperlink")

    @mcp.tool(name="word_add_bookmark", annotations=annotations("Add a bookmark", WRITES))
    async def word_add_bookmark(params: AddBookmarkInput) -> str:
        """Name a place in the document: a paragraph or some text in it. A link
        (word_insert_hyperlink) or a field (REF name, PAGEREF name) can then
        point to it."""
        p = params
        return await call(_add_bookmark_impl, p.name, p.index, p.match, undo="Add bookmark")

    @mcp.tool(name="word_list_bookmarks", annotations=annotations("List bookmarks", READ_ONLY))
    async def word_list_bookmarks(params: Input) -> str:
        """The document's bookmarks in order, each with its paragraph and the
        text it marks. Word's hidden ones (for cross-references) are left out."""
        return await call(_list_bookmarks_impl)

    @mcp.tool(name="word_delete_bookmark", annotations=annotations("Delete a bookmark", DESTRUCTIVE))
    async def word_delete_bookmark(params: DeleteBookmarkInput) -> str:
        """Delete a bookmark. The text it marked stays; links and REF fields
        that point to it stop working."""
        return await call(_delete_bookmark_impl, params.name, undo="Delete bookmark")
