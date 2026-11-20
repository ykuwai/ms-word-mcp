"""Comments.

A comment is about some text in the document, its scope, and Word numbers the
comments in the order of their scopes. On Windows a reply is a comment of its
own in the same list, with its comment as its Ancestor, and a comment can be
marked done.
https://learn.microsoft.com/en-us/office/vba/api/word.comments.add
https://learn.microsoft.com/en-us/office/vba/api/word.comment

On Word for Mac a comment is made with `make new Word comment at <document>`,
its scope and its text both given as properties. Made with the scope alone,
Word for Mac ended its process (measured on 16.114, 2026-11-20, and in
scripts/probes/mac_comments_revisions.py). Its scripting has no replies and no
done state; they go through the VBA bridge (backend/bridge.py), and when it
cannot be used those two tools say why.
"""

from __future__ import annotations

from pydantic import Field, field_validator, model_validator

from .. import backend
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, annotations, call, follow, mac_impl
from ._word import clean_text, locate_spans, paragraph
from .revisions import when

#: How much of the text a comment is about an answer holds.
ON_SHOWN = 200

NO_REPLIES_ON_MAC = (
    "Word for Mac's scripting has no replies to comments, and ms-word-mcp's VBA bridge, "
    "which adds them, cannot be used with this Word ({why}). Add a comment of its own on "
    "the same text instead (word_add_comment).")
NO_DONE_ON_MAC = (
    "Word for Mac's scripting has no done state for comments, and ms-word-mcp's VBA "
    "bridge, which adds it, cannot be used with this Word ({why}). Delete the comment "
    "once it is dealt with instead (word_delete_comments).")


def _one_paragraph(text: str) -> str:
    if "\r" in text or "\n" in text:
        raise ValueError("A comment is written as one paragraph here: take the line breaks out.")
    return text


class AddCommentInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the comment is about.")
    text: str = Field(..., min_length=1, max_length=5000, description="The comment.")
    on: str | None = Field(None, min_length=1, max_length=255, description=(
        "The text in the paragraph the comment is about: its first occurrence. Left out: "
        "the whole paragraph."))

    _text = field_validator("text")(_one_paragraph)


class ListCommentsInput(Input):
    start: int = Field(1, ge=1, description="The comment to start from (numbered from 1).")
    limit: int = Field(100, ge=1, le=200, description="How many to read at most.")


class ReplyInput(Input):
    number: int = Field(..., ge=1, description="Which comment, as word_list_comments numbers them.")
    text: str = Field(..., min_length=1, max_length=5000, description="The reply.")

    _text = field_validator("text")(_one_paragraph)


class ResolveCommentInput(Input):
    number: int = Field(..., ge=1, description="Which comment, as word_list_comments numbers them.")
    done: bool = Field(True, description="True to mark it done, false to open it again.")


class DeleteCommentsInput(Input):
    numbers: list[int] | None = Field(None, min_length=1, max_length=200, description=(
        "Which comments, as word_list_comments numbers them."))
    every: bool = Field(False, description="Every comment in the document.")

    @model_validator(mode="after")
    def _one_way(self):
        if (self.numbers is None) == (not self.every):
            raise ValueError("Give `numbers`, or `every` true: one of the two.")
        if self.numbers is not None and min(self.numbers) < 1:
            raise ValueError("Comments are numbered from 1.")
        return self


def check_number(number: int, count: int) -> None:
    if not 1 <= number <= count:
        raise ValueError(f"There is no comment {number}: the document has {count} "
                         "(numbered from 1; word_list_comments lists them). Nothing was changed.")


def highest_first(numbers: list[int], count: int) -> list[int]:
    """The numbers from the highest down, so that each holds until its turn."""
    for number in numbers:
        check_number(number, count)
    return sorted(set(numbers), reverse=True)


def not_in_paragraph(on: str, index: int) -> ValueError:
    return ValueError(f"'{on}' is not in paragraph {index}. Nothing was added.")


def shown_on(text) -> dict:
    """The text a comment is about, cut short when it is long."""
    text = clean_text(text)
    if len(text) <= ON_SHOWN:
        return {"on": text}
    return {"on": text[:ON_SHOWN], "on_cut": True}


def listing(out: list[dict], count: int, start: int) -> dict:
    result = {"success": True, "count": count, "comments": out}
    last = start + len(out) - 1
    if out and last < count:
        result["more"] = f"Comments {last + 1} to {count} follow; pass start={last + 1}."
    return result


def deleted(asked: int, before: int, left: int) -> dict:
    """What a deletion did: `asked` comments named, `before` and `left` counted."""
    result = {"success": True, "deleted": before - left, "left": left}
    if before - left > asked:
        result["note"] = (f"{asked} comment{'s' if asked != 1 else ''} named, {before - left} "
                          "gone: deleting a comment deletes its replies too.")
    elif left:
        result["note"] = "The comments after these have new numbers; list them again."
    return result


# -- Windows ---------------------------------------------------------------------------

def _scope(doc, index, on):
    """The range a new comment is about: `on` in paragraph `index`, or its text."""
    rng = paragraph(doc, index).Range
    if on is None:
        return doc.Range(rng.Start, max(rng.Start, rng.End - 1))
    raw = rng.Text
    if raw.find(on) < 0:
        raise not_in_paragraph(on, index)
    hidden = max(0, rng.End - rng.Start - len(raw.encode("utf-16-le")) // 2)
    start, end = locate_spans(raw, [on], lambda a, b: doc.Range(a, b).Text, rng.Start, hidden)[0]
    return doc.Range(start, end)


def _numbered(doc):
    """The document's comments in Word's order, and each one's number by where its
    text starts in the comments' story: two references to one comment do not
    compare equal over COM, and no two comments start at one place."""
    comments = doc.Comments
    items = [comments.Item(i) for i in range(1, comments.Count + 1)]
    return items, {c.Range.Start: n for n, c in enumerate(items, 1)}


def _describe(doc, number: int, comment, numbers: dict) -> dict:
    scope = comment.Scope
    ancestor = comment.Ancestor
    described = {"number": number, "index": doc.Range(0, scope.Start + 1).Paragraphs.Count,
                 **shown_on(scope.Text), "text": clean_text(comment.Range.Text),
                 "author": comment.Author, "initials": comment.Initial,
                 "date": when(comment.Date), "done": bool(comment.Done)}
    if ancestor is not None:
        described["reply_to"] = numbers.get(ancestor.Range.Start)
    return described


def _find(doc, comment) -> dict:
    items, numbers = _numbered(doc)
    number = numbers[comment.Range.Start]
    return _describe(doc, number, items[number - 1], numbers)


def _add_comment_impl(index, text, on):
    doc = backend.word.document()
    scope = _scope(doc, index, on)
    comment = doc.Comments.Add(scope, text)
    follow(doc, comment.Scope)
    return {"success": True, **_find(doc, comment), "count": doc.Comments.Count}


def _list_comments_impl(start, limit):
    doc = backend.word.document()
    items, numbers = _numbered(doc)
    last = min(len(items), start + limit - 1)
    out = [_describe(doc, n, items[n - 1], numbers) for n in range(start, last + 1)]
    return listing(out, len(items), start)


def _reply_to_comment_impl(number, text):
    doc = backend.word.document()
    check_number(number, doc.Comments.Count)
    comment = doc.Comments.Item(number)
    while comment.Ancestor is not None:  # a reply to a reply goes to its comment
        comment = comment.Ancestor
    reply = comment.Replies.Add(comment.Scope, text)
    return {"success": True, **_find(doc, reply), "count": doc.Comments.Count}


def _resolve_comment_impl(number, done):
    doc = backend.word.document()
    check_number(number, doc.Comments.Count)
    comment = doc.Comments.Item(number)
    comment.Done = done
    return {"success": True, **_find(doc, comment)}


def _delete_comments_impl(numbers, every):
    doc = backend.word.document()
    comments = doc.Comments
    before = comments.Count
    order = list(range(before, 0, -1)) if every else highest_first(numbers, before)
    asked = 0
    for number in order:
        if number <= comments.Count:  # not taken already, as another's reply
            comments.Item(number).DeleteRecursively()
            asked += 1
    return deleted(len(order) if every else asked, before, comments.Count)


# -- Mac -------------------------------------------------------------------------------

def _mac_count(doc) -> int:
    from . import mac as m

    return int(doc.count(each=m._k().Word_comment))


def _mac_scope(doc, index, on):
    from . import mac as m

    m._para(doc, index)
    start, end = m._bounds(doc, index)
    if on is None:
        return doc.create_range(start=start, end_=max(start, end - 1))
    raw = doc.paragraphs[index].text_object.content.get()
    if raw.find(on) < 0:
        raise not_in_paragraph(on, index)

    def read(a, b):
        try:
            return doc.create_range(start=a, end_=b).content.get()
        except Exception:  # noqa: BLE001 - half a character: not here
            return None

    hidden = max(0, end - start - m._u16(raw, len(raw)))
    a, b = locate_spans(raw, [on], read, start, hidden)[0]
    return doc.create_range(start=a, end_=b)


def _mac_describe(doc, number: int, comment) -> dict:
    from . import mac as m

    k = m._k()
    scope = comment.scope
    at = scope.start_of_content.get()
    return {"number": number, "index": int(doc.create_range(start=0, end_=at + 1).count(each=k.paragraph)),
            **shown_on(m._clean(scope.content.get())),
            "text": m._clean(comment.comment_text.content.get()),
            "author": comment.author.get(), "initials": comment.initials.get(),
            "date": when(comment.date_value.get())}


def _mac_bridge(doc, what: str, *fields: str, refused: str | None = None):
    """Ask the VBA bridge about this document. None when it cannot be used, or
    with `refused`, that message, saying why."""
    from . import mac as m
    from ..backend.bridge import BridgeError, Unavailable, encode

    word = backend.word
    try:
        vba = getattr(word, "vba", None)
        if vba is None:
            raise Unavailable("this server has no bridge")
        return vba.call(word.get_app(), what, encode(m._full(doc)), *fields)
    except Unavailable as exc:
        if refused:
            raise ValueError(refused.format(why=exc)) from None
        return None
    except BridgeError as exc:
        raise ValueError(f"Word refused it: {exc}") from None


def _mac_states(doc) -> dict | None:
    """Each comment's done state and the comment it replies to, by number;
    None when the bridge cannot be used."""
    lines = _mac_bridge(doc, "CommentState")
    if lines is None:
        return None
    rows = [[int(x) for x in line.split(",")] for line in lines]
    by_start = {own: n for n, (_, _, own) in enumerate(rows, 1)}
    return {n: {"done": bool(done), **({"reply_to": by_start.get(parent)} if parent >= 0 else {})}
            for n, (done, parent, _) in enumerate(rows, 1)}


def _mac_find(doc, number: int, states=None) -> dict:
    described = _mac_describe(doc, number, doc.Word_comments[number])
    states = _mac_states(doc) if states is None else states
    if states:
        described.update(states.get(number, {}))
    return described


@mac_impl(_add_comment_impl)
def _mac_add_comment(index, text, on):
    from . import mac as m

    k = m._k()
    doc = backend.word.document()
    scope = _mac_scope(doc, index, on)
    # The text goes in with the scope: made with the scope alone, Word for Mac
    # ended its process (module docstring).
    comment = doc.make(new=k.Word_comment, at=doc,
                       with_properties={k.scope: scope, k.comment_text: text})
    number = int(comment.comment_index.get())
    return {"success": True, **_mac_find(doc, number), "count": _mac_count(doc)}


@mac_impl(_list_comments_impl)
def _mac_list_comments(start, limit):
    doc = backend.word.document()
    count = _mac_count(doc)
    last = min(count, start + limit - 1)
    states = _mac_states(doc) or {}
    out = [{**_mac_describe(doc, n, doc.Word_comments[n]), **states.get(n, {})}
           for n in range(start, last + 1)]
    return listing(out, count, start)


@mac_impl(_reply_to_comment_impl)
def _mac_reply_to_comment(number, text):
    doc = backend.word.document()
    check_number(number, _mac_count(doc))
    from ..backend.bridge import encode

    lines = _mac_bridge(doc, "Reply", str(number), encode(text), refused=NO_REPLIES_ON_MAC)
    return {"success": True, **_mac_find(doc, int(lines[0])), "count": int(lines[1])}


@mac_impl(_resolve_comment_impl)
def _mac_resolve_comment(number, done):
    doc = backend.word.document()
    check_number(number, _mac_count(doc))
    _mac_bridge(doc, "SetDone", str(number), "1" if done else "0", refused=NO_DONE_ON_MAC)
    return {"success": True, **_mac_find(doc, number)}


@mac_impl(_delete_comments_impl)
def _mac_delete_comments(numbers, every):
    doc = backend.word.document()
    before = _mac_count(doc)
    order = list(range(before, 0, -1)) if every else highest_first(numbers, before)
    # Through the bridge when it is there: Word for Mac's `delete` leaves a
    # comment that has replies where it is (measured on 16.114).
    lines = _mac_bridge(doc, "DeleteComments", ",".join(map(str, order))) if order else None
    if lines is not None:
        left, asked = int(lines[0]), int(lines[1])
        return deleted(len(order) if every else asked, before, left)
    # One at a time from the last: `delete all comments` is refused (-1708).
    for number in order:
        doc.Word_comments[number].delete()
    return deleted(len(order), before, _mac_count(doc))


def register(mcp) -> None:
    @mcp.tool(name="word_add_comment", annotations=annotations("Add a comment", WRITES))
    async def word_add_comment(params: AddCommentInput) -> str:
        """Add a comment about some text: a paragraph's whole text, or the first
        occurrence of `on` in it. Word numbers comments in the order of the text
        they are about, so the ones after it renumber; `number` is where this one
        landed. Word signs it with the name of the Office user, which `author`
        gives back."""
        p = params
        return await call(_add_comment_impl, p.index, p.text, p.on, undo="Add comment")

    @mcp.tool(name="word_list_comments", annotations=annotations("List comments", READ_ONLY))
    async def word_list_comments(params: ListCommentsInput) -> str:
        """The comments, in order: each one's number, the paragraph its text
        starts in (`index`), the text it is about (`on`), the comment, its
        author, initials and date, whether it is done, and for a reply, the
        comment it answers (`reply_to`). On the Mac, the last two need
        ms-word-mcp's VBA bridge, and are left out when it cannot be used."""
        p = params
        return await call(_list_comments_impl, p.start, p.limit)

    @mcp.tool(name="word_reply_to_comment", annotations=annotations("Reply to a comment", WRITES))
    async def word_reply_to_comment(params: ReplyInput) -> str:
        """Reply to a comment; a reply to a reply goes to its comment. On the
        Mac this goes through ms-word-mcp's VBA bridge."""
        p = params
        return await call(_reply_to_comment_impl, p.number, p.text, undo="Reply to comment")

    @mcp.tool(name="word_resolve_comment", annotations=annotations("Resolve a comment", WRITES))
    async def word_resolve_comment(params: ResolveCommentInput) -> str:
        """Mark a comment done, or open it again. On the Mac this goes through
        ms-word-mcp's VBA bridge."""
        p = params
        return await call(_resolve_comment_impl, p.number, p.done, undo="Resolve comment")

    @mcp.tool(name="word_delete_comments", annotations=annotations("Delete comments", DESTRUCTIVE))
    async def word_delete_comments(params: DeleteCommentsInput) -> str:
        """Delete comments, by number or every one. Deleting a comment deletes
        its replies too, and the comments after it renumber, so list them again
        before the next call."""
        p = params
        return await call(_delete_comments_impl, p.numbers, p.every, undo="Delete comments")
