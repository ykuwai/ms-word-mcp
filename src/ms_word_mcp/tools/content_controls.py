"""Content controls: the tagged boxes of templates, forms and add-ins.

A content control holds a stretch of text and carries a tag and a title that
the reader does not see. Templates mark their fields with them, and add-ins
keep their data in their tags (Mendeley Cite's citations, for one: a tag of
tens of thousands of characters).

The helpers here (``controls``, ``add``, ``update``, ``remove``, ``place``)
work on either platform for the document a call works on, so a tool built on
content controls has one body: on Windows through COM, on the Mac through
ms-word-mcp's VBA bridge, since Word for Mac's scripting has no content
controls at all.
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, annotations, call, follow, mac_impl
from ._word import clean_text, paragraph

#: WdContentControlType.
TYPES = {0: "rich_text", 1: "plain_text", 2: "picture", 3: "combo_box", 4: "drop_down_list",
         5: "building_block_gallery", 6: "date", 7: "group", 8: "check_box",
         9: "repeating_section"}
KINDS = {"rich_text": 0, "plain_text": 1}

NO_BRIDGE = ("Word for Mac's scripting has no content controls, and ms-word-mcp's VBA bridge, "
             "which reaches them, cannot be used with this Word ({why}).")
NO_CONTROL = "There is no content control {nth}: the document has {count}{which}."


# -- the bridge's wire format -------------------------------------------------------------

def wire(text: str) -> str:
    """As the bridge's macros read a tag, title or text: plain ASCII as it is
    ("r:"), anything else as UTF-16 codes ("e:")."""
    from ..backend.bridge import encode

    if all(32 <= ord(c) <= 126 for c in text):
        return "r:" + text
    return "e:" + encode(text)


def unwire(field: str) -> str:
    from ..backend.bridge import decode

    return decode(field[2:]) if field.startswith("e:") else field[2:]


# -- the controls of the document a call works on ------------------------------------------

def _bridge(doc):
    from . import _vba
    from . import mac as m
    from ..backend.bridge import encode

    why = _vba.usable()
    if why:
        raise ValueError(NO_BRIDGE.format(why=why))
    word = backend.word
    return lambda what, *fields: word.vba.call(word.get_app(), what, encode(m._full(doc)), *fields)


def _entry(start, end, index, kind, locks, title, text, tag) -> dict:
    return {"start": int(start), "end": int(end), "index": int(index),
            "type": TYPES.get(int(kind), int(kind)), "lock_control": locks[0] == "1",
            "lock_contents": locks[1] == "1", "title": title, "text": clean_text(text), "tag": tag}


def controls(prefix: str = "") -> list[dict]:
    """The document's content controls whose tag starts with `prefix`, in
    reading order: start and end of their text, paragraph, type, locks,
    title, text and tag."""
    doc = backend.word.document()
    out = []
    if backend.IS_MACOS:
        for line in _bridge(doc)("ContentControlList", wire(prefix)):
            start, end, index, kind, locks, title, text, tag = line.split("|", 7)
            out.append(_entry(start, end, index, kind, locks, unwire(title), unwire(text),
                              unwire(tag)))
    else:
        found = doc.ContentControls
        for i in range(1, found.Count + 1):
            cc = found.Item(i)
            tag = cc.Tag or ""
            if not tag.startswith(prefix):
                continue
            r = cc.Range
            locks = ("1" if cc.LockContentControl else "0") + ("1" if cc.LockContents else "0")
            out.append(_entry(r.Start, r.End, doc.Range(0, r.Start).Paragraphs.Count, cc.Type,
                              locks, cc.Title or "", r.Text or "", tag))
    out.sort(key=lambda c: c["start"])
    for n, c in enumerate(out, 1):
        c["nth"] = n
    return out


def _com_at(doc, start: int):
    found = doc.ContentControls
    for i in range(1, found.Count + 1):
        if found.Item(i).Range.Start == start:
            return found.Item(i)
    raise ValueError(f"No content control's text starts at {start}.")


def add(at: int, kind: str = "rich_text", tag: str = "", title: str = "", text: str = "",
        end: int | None = None) -> tuple[int, int]:
    """A content control at `at` (an empty place) holding `text`, or, given
    `end`, around the text from `at` to `end`, which stays as it is (its
    formatting too). The start and end of its text."""
    doc = backend.word.document()
    stop = at if end is None else end
    if backend.IS_MACOS:
        line = _bridge(doc)("ContentControlAdd", str(at), str(stop), str(KINDS[kind]), wire(tag),
                            wire(title), "-" if end is not None else wire(text))[0]
        start, finish = line.split("|")
        return int(start), int(finish)
    cc = doc.ContentControls.Add(KINDS[kind], doc.Range(at, stop))
    cc.Tag = tag
    cc.Title = title
    if end is None:
        cc.Range.Text = text
    return cc.Range.Start, cc.Range.End


def update(start: int, tag=None, title=None, text=None, lock_control=None, lock_contents=None) -> tuple[int, int]:
    """Change the control whose text starts at `start`; None leaves a property."""
    doc = backend.word.document()
    if backend.IS_MACOS:
        flag = {True: "1", False: "0", None: "-"}
        line = _bridge(doc)("ContentControlSet", str(start),
                            "-" if tag is None else wire(tag), "-" if title is None else wire(title),
                            "-" if text is None else wire(text), flag[lock_control],
                            flag[lock_contents])[0]
        s, e = line.split("|")
        return int(s), int(e)
    cc = _com_at(doc, start)
    held = cc.LockContents
    if tag is not None:
        cc.Tag = tag
    if title is not None:
        cc.Title = title
    if text is not None:
        cc.LockContents = False
        cc.Range.Text = text
        cc.LockContents = held
    if lock_control is not None:
        cc.LockContentControl = lock_control
    if lock_contents is not None:
        cc.LockContents = lock_contents
    return cc.Range.Start, cc.Range.End


def remove(start: int, with_text: bool) -> None:
    doc = backend.word.document()
    if backend.IS_MACOS:
        _bridge(doc)("ContentControlDelete", str(start), "1" if with_text else "0")
        return
    cc = _com_at(doc, start)
    cc.LockContentControl = False
    if with_text:
        cc.LockContents = False
    cc.Delete(with_text)


def place(index: int, after: str | None, before: str | None, position: str) -> int:
    """Where in paragraph `index` something goes: right after or before the
    first occurrence of a text, or at the start or end of its text."""
    if after or before:
        s, e = find(index, after or before)
        return e if after else s
    return _find(index, None, position)


def find(index: int, text: str) -> tuple[int, int]:
    """Where the first occurrence of `text` in paragraph `index` starts and ends."""
    return _find(index, text, "end")


def _find(index: int, anchor: str | None, position: str):
    from ._word import locate_spans

    doc = backend.word.document()
    if backend.IS_MACOS:
        from . import mac as m

        m._para(doc, index)
        start, end = m._bounds(doc, index)
        if anchor is None:
            return start if position == "start" else end - 1
        content = doc.paragraphs[index].text_object.content.get() or ""

        def read(a, b):
            try:
                return doc.create_range(start=a, end_=b).content.get()
            except Exception:  # noqa: BLE001 - half a character: not here
                return None

        hidden = max(0, end - start - m._u16(content, len(content)))
    else:
        rng = paragraph(doc, index).Range
        start, end = rng.Start, rng.End
        if anchor is None:
            return start if position == "start" else end - 1
        content = rng.Text or ""
        hidden = max(0, end - start - len(content.encode("utf-16-le")) // 2)

        def read(a, b):
            return doc.Range(a, b).Text

    if content.find(anchor) < 0:
        raise ValueError(f"'{anchor}' is not in paragraph {index}. Nothing was changed.")
    return locate_spans(content, [anchor], read, start, hidden)[0]


def show(start: int, end: int) -> None:
    """Scroll the document's window to a control, as every write does."""
    doc = backend.word.document()
    if backend.IS_MACOS:
        from . import mac as m

        m._scroll(doc, start, end)
    else:
        follow(doc, doc.Range(start, end))


def view(c: dict, full_tag: bool) -> dict:
    """A control as a tool answers it: a long tag cut, its length given."""
    out = {"nth": c["nth"], "index": c["index"], "type": c["type"], "text": c["text"][:500]}
    if c["title"]:
        out["title"] = c["title"]
    if c["tag"]:
        out["tag"] = c["tag"] if full_tag or len(c["tag"]) <= 200 else c["tag"][:200] + "…"
        if len(c["tag"]) > 200:
            out["tag_length"] = len(c["tag"])
    if c["lock_control"]:
        out["lock_control"] = True
    if c["lock_contents"]:
        out["lock_contents"] = True
    return out


def nth_control(nth: int, prefix: str = "") -> dict:
    found = controls(prefix)
    if not 1 <= nth <= len(found):
        which = f" whose tag starts with '{prefix}'" if prefix else ""
        raise ValueError(NO_CONTROL.format(nth=nth, count=len(found), which=which))
    return found[nth - 1]


# -- inputs -----------------------------------------------------------------------------------

class _Place(Input):
    after: str | None = Field(None, min_length=1, max_length=255, description=(
        "Right after the first occurrence of this text in the paragraph."))
    before: str | None = Field(None, min_length=1, max_length=255, description=(
        "Or right before it."))
    position: Literal["end", "start"] = Field("end", description=(
        "Without `after` or `before`: at the end of the paragraph's text, or its start."))

    @model_validator(mode="after")
    def _one_place(self):
        if self.after is not None and self.before is not None:
            raise ValueError("Give `after` or `before`, not both.")
        return self


class ListContentControlsInput(Input):
    tag_prefix: str | None = Field(None, max_length=200, description=(
        "Only the controls whose tag starts with this."))
    start: int = Field(1, ge=1, description="The control to start from (`nth`, from 1).")
    limit: int = Field(50, ge=1, le=200)
    full_tags: bool = Field(False, description=(
        "Give tags of more than 200 characters whole (add-ins keep their data in them)."))


class InsertContentControlInput(_Place):
    index: int = Field(..., ge=1, description="The paragraph the control goes in.")
    text: str = Field("", max_length=32767, description="The text it holds.")
    around: str | None = Field(None, min_length=1, max_length=32767, description=(
        "Instead of `text` and a place: make the first occurrence of this text in the "
        "paragraph the control's, as it is, formatting and all."))
    kind: Literal[tuple(KINDS)] = Field("rich_text", description=(  # type: ignore[valid-type]
        "rich_text holds formatted text; plain_text holds text of one format."))
    tag: str = Field("", max_length=100000, description="A tag, for a program to find it by.")
    title: str = Field("", max_length=255, description="The title Word shows on its box.")
    lock_control: bool = Field(False, description="The control itself cannot be deleted.")
    lock_contents: bool = Field(False, description="Its text cannot be edited.")

    @model_validator(mode="after")
    def _around_alone(self):
        if self.around is not None and (self.text or self.after or self.before):
            raise ValueError("`around` makes text already there the control's: give it "
                             "without `text`, `after` or `before`.")
        return self


class SetContentControlInput(Input):
    control: int = Field(..., ge=1, description="Which control (`nth` in word_list_content_controls).")
    tag_prefix: str | None = Field(None, max_length=200, description=(
        "Count `control` among the controls whose tag starts with this."))
    text: str | None = Field(None, max_length=32767)
    tag: str | None = Field(None, max_length=100000)
    title: str | None = Field(None, max_length=255)
    lock_control: bool | None = None
    lock_contents: bool | None = None

    @model_validator(mode="after")
    def _something(self):
        if all(v is None for v in (self.text, self.tag, self.title, self.lock_control,
                                   self.lock_contents)):
            raise ValueError("Give at least one of text, tag, title, lock_control, lock_contents.")
        return self


class DeleteContentControlInput(Input):
    control: int = Field(..., ge=1, description="Which control (`nth` in word_list_content_controls).")
    tag_prefix: str | None = Field(None, max_length=200)
    keep_text: bool = Field(True, description=(
        "Keep its text in the document, as plain text; false deletes the text too."))


# -- the tools' bodies: the same on both platforms, through the helpers above -----------------

def _list_content_controls_impl(tag_prefix, start, limit, full_tags):
    found = controls(tag_prefix or "")
    out = {"success": True, "count": len(found),
           "controls": [view(c, full_tags) for c in found if start <= c["nth"] < start + limit]}
    if start + limit <= len(found):
        out["more"] = f"Controls {start + limit} to {len(found)} follow; pass start={start + limit}."
    return out


def _insert_content_control_impl(index, text, kind, tag, title, lock_control, lock_contents,
                                 after, before, position, around):
    if around is not None:
        s, e = find(index, around)
        start, end = add(s, kind, tag, title, end=e)
    else:
        start, end = add(place(index, after, before, position), kind, tag, title, text)
    if lock_control or lock_contents:
        start, end = update(start, lock_control=lock_control or None,
                            lock_contents=lock_contents or None)
    show(start, end)
    made = next(c for c in controls() if c["start"] == start)
    return {"success": True, **view(made, False)}


def _set_content_control_impl(control, tag_prefix, text, tag, title, lock_control, lock_contents):
    c = nth_control(control, tag_prefix or "")
    start, end = update(c["start"], tag, title, text, lock_control, lock_contents)
    show(start, end)
    now = next(x for x in controls() if x["start"] == start)
    return {"success": True, **view(now, False)}


def _delete_content_control_impl(control, tag_prefix, keep_text):
    c = nth_control(control, tag_prefix or "")
    remove(c["start"], not keep_text)
    return {"success": True, "deleted": c["nth"], "kept_text": c["text"][:500] if keep_text else "",
            "left": len(controls())}


for _impl in (_list_content_controls_impl, _insert_content_control_impl,
              _set_content_control_impl, _delete_content_control_impl):
    mac_impl(_impl)(_impl)  # one body: the helpers choose COM or the bridge
del _impl


# -- the tools --------------------------------------------------------------------------------

def register(mcp) -> None:
    @mcp.tool(name="word_list_content_controls",
              annotations=annotations("List content controls", READ_ONLY))
    async def word_list_content_controls(params: ListContentControlsInput) -> str:
        """The document's content controls (the boxes of templates, forms and
        add-ins), in order: each one's `nth`, paragraph (`index`), type, title,
        tag, text and locks; only those whose tag starts with `tag_prefix` if
        given. On the Mac this goes through ms-word-mcp's VBA bridge."""
        p = params
        return await call(_list_content_controls_impl, p.tag_prefix, p.start, p.limit, p.full_tags)

    @mcp.tool(name="word_insert_content_control",
              annotations=annotations("Insert a content control", WRITES))
    async def word_insert_content_control(params: InsertContentControlInput) -> str:
        """Insert a content control holding `text`, with a tag and title,
        right after or before a text of a paragraph, or at its start or end;
        or, with `around`, around a text already there, which keeps its
        formatting. Optionally locked. On the Mac this goes through
        ms-word-mcp's VBA bridge."""
        p = params
        return await call(_insert_content_control_impl, p.index, p.text, p.kind, p.tag, p.title,
                          p.lock_control, p.lock_contents, p.after, p.before, p.position,
                          p.around, undo="Insert content control")

    @mcp.tool(name="word_set_content_control",
              annotations=annotations("Change a content control", WRITES))
    async def word_set_content_control(params: SetContentControlInput) -> str:
        """Change a content control's text, tag, title or locks; what is not
        given stays. A control whose contents are locked takes a new text all
        the same, and stays locked. On the Mac this goes through ms-word-mcp's
        VBA bridge."""
        p = params
        return await call(_set_content_control_impl, p.control, p.tag_prefix, p.text, p.tag,
                          p.title, p.lock_control, p.lock_contents, undo="Change content control")

    @mcp.tool(name="word_delete_content_control",
              annotations=annotations("Delete a content control", DESTRUCTIVE))
    async def word_delete_content_control(params: DeleteContentControlInput) -> str:
        """Delete a content control: by default its text stays in the document
        as plain text; `keep_text: false` deletes the text too. On the Mac this
        goes through ms-word-mcp's VBA bridge."""
        p = params
        return await call(_delete_content_control_impl, p.control, p.tag_prefix, p.keep_text,
                          undo="Delete content control")
