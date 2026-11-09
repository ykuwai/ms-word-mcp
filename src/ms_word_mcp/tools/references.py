"""Pictures, captions and cross-references.

A picture goes in a paragraph of its own, in line with the text. A caption is
a paragraph in the Caption style holding a SEQ field, the way Word's own
Insert Caption makes one, under any label in any language ("Figure", "図",
"Abbildung"). A cross-reference is a hidden bookmark over what it shows, and
a REF or PAGEREF field to the bookmark, which is what Word's own dialog makes
(https://learn.microsoft.com/en-us/office/vba/api/word.bookmarks.showhidden).
Both are built from fields and bookmarks rather than Range.InsertCaption and
Range.InsertCrossReference, whose labels must already exist in this Word's
language; the Mac takes only the built-in ones.

Word for Mac is sandboxed and asks the user about each file outside its own
container; the call waits and times out (-1712, measured). So a picture is
copied into the container for the moment it is read.
"""

from __future__ import annotations

import itertools
import os
import shutil
import tempfile
import time
from typing import Literal

from pydantic import Field, field_validator, model_validator

from .. import backend
from ._common import (READ_ONLY, WRITES, Input, Trimmed, annotations, call, follow, local_path,
                      mac_impl)
from ._word import ALIGNMENTS, clean_text, paragraph
from .shapes import PICTURE_LOOK, _PictureLook, _Position, _Wrap, faded

CONTAINER = os.path.expanduser("~/Library/Containers/com.microsoft.Word/Data/Documents")
ALIGN_CHOICES = Literal["left", "center", "right"]
_bookmark_numbers = itertools.count(int(time.time()) % 10 ** 8 * 10)


def _bookmark_name() -> str:
    return f"_Ref{next(_bookmark_numbers) % 10 ** 9:09d}"


def _u16(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


# -- arguments ---------------------------------------------------------------------

class InsertPictureInput(Input):
    path: Trimmed = Field(..., min_length=1, description="Full path of the picture file.")
    after: int | None = Field(None, ge=0, description=(
        "Put it after this paragraph (0 = at the very start). Neither after nor "
        "before: at the end."))
    before: int | None = Field(None, ge=1, description="Put it before this paragraph.")
    width: float | None = Field(None, gt=0, description=(
        "Width in points (72 = 1 inch). Given alone, the height follows in proportion."))
    height: float | None = Field(None, gt=0, description=(
        "Height in points. Given alone, the width follows in proportion."))
    alignment: ALIGN_CHOICES | None = Field(None, description="Its paragraph's alignment.")
    alt_text: str | None = Field(None, description=(
        "What the picture shows, for readers who cannot see it."))

    @field_validator("path")
    @classmethod
    def _absolute(cls, value: str) -> str:
        if not os.path.isabs(value):
            raise ValueError("path must be a full path.")
        return local_path(value)

    @model_validator(mode="after")
    def _where(self):
        if self.after is not None and self.before is not None:
            raise ValueError("Give `after` or `before`, not both.")
        return self


class SetPictureInput(_Position, _Wrap, _PictureLook):
    picture: int = Field(..., ge=1, description="The picture, numbered from 1 in the document.")
    width: float | None = Field(None, gt=0, description=(
        "Width in points. Given alone, the height follows in proportion."))
    height: float | None = Field(None, gt=0, description=(
        "Height in points. Given alone, the width follows in proportion."))
    alignment: ALIGN_CHOICES | None = None
    alt_text: str | None = None
    name: Trimmed | None = Field(None, min_length=1, max_length=255, description=(
        "With wrap: a name of its own once it floats."))

    @model_validator(mode="after")
    def _something(self):
        given = self.model_dump(exclude_none=True, exclude={"picture"})
        if not given:
            raise ValueError("Give at least one thing to change.")
        floating = [n for n in ("left", "top", "horizontal_from", "vertical_from", "wrap_side",
                                "wrap_distance", "name") if n in given]
        if self.wrap is None and floating:
            raise ValueError(f"{', '.join(floating)} only mean something for a picture that "
                             "floats. Give `wrap` too, such as square or behind.")
        if self.wrap == "inline":
            raise ValueError("It is in line with the text already. A floating picture goes "
                             "back with word_set_shape and wrap inline.")
        return self


class InsertCaptionInput(Input):
    text: str = Field("", description="The caption's words, after its label and number.")
    label: Trimmed = Field("Figure", min_length=1, max_length=40, description=(
        "The label, in the document's language: Figure, Table, 図, 表, Abbildung... "
        "Each label is numbered on its own."))
    separator: str = Field(" ", max_length=10, description=(
        "What comes between the number and the words, such as \": \"."))
    for_picture: int | None = Field(None, ge=1, description=(
        "Put the caption under this picture."))
    for_table: int | None = Field(None, ge=1, description="Put the caption above this table.")
    after: int | None = Field(None, ge=0, description="Or after this paragraph.")
    before: int | None = Field(None, ge=1, description="Or before this paragraph.")

    @model_validator(mode="after")
    def _where(self):
        given = [n for n in ("for_picture", "for_table", "after", "before")
                 if getattr(self, n) is not None]
        if len(given) != 1:
            raise ValueError("Give one of for_picture, for_table, after or before.")
        return self


class CrossReferenceInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the reference goes in.")
    to: int = Field(..., ge=1, description=(
        "The paragraph it refers to: a caption or numbered equation "
        "(word_list_captions), or a heading (word_get_outline)."))
    show: Literal["label_and_number", "number", "text", "page"] = Field(
        "label_and_number", description=(
            "label_and_number: \"Figure 1\". number: \"1\". text: the whole paragraph, as "
            "for a heading. page: its page number."))
    position: Literal["end", "start"] = Field("end", description=(
        "At the end of the paragraph's text or at its start."))
    match: str | None = Field(None, min_length=1, max_length=255, description=(
        "Instead of `position`: replace the first occurrence of this text in the "
        "paragraph, such as a placeholder: \"see FIG\" with match \"FIG\". Several "
        "references can go in one paragraph, in any order."))


class NoInput(Input):
    pass


def seq_name(label: str) -> str:
    """The SEQ field's identifier for a label: one word, as a field needs."""
    return "_".join(label.split())


def proportional(width, height, old_width, old_height) -> tuple[float | None, float | None]:
    """The new size, the missing side following the given one in proportion."""
    if width is not None and height is None and old_width:
        return width, round(old_height * width / old_width, 2)
    if height is not None and width is None and old_height:
        return round(old_width * height / old_height, 2), height
    return width, height


def _place(count: int, after, before) -> tuple[int | None, int | None]:
    if before is not None and before > count:
        raise ValueError(f"There is no paragraph {before}: the document has {count}.")
    if after is not None and after > count:
        raise ValueError(f"There is no paragraph {after}: the document has {count}.")
    return after, before


def _too_wide(width, text_width) -> list[str]:
    if text_width and width > text_width + 1:
        return [f"The picture is {width:g} pt wide and the text is {text_width:g} pt wide, "
                "so it runs into the right margin. Narrow it with word_set_picture."]
    return []


def _reference_field(bookmark: str, show: str) -> str:
    return f"{'PAGEREF' if show == 'page' else 'REF'} {bookmark} \\h"


NO_NUMBER = ("Paragraph {index} has no number to refer to: it holds no SEQ field. "
             "word_list_captions lists the captions and numbered equations; for a "
             "heading use show \"text\" or \"page\".")


# -- Windows -----------------------------------------------------------------------

def _shape(doc, number: int):
    count = doc.InlineShapes.Count
    if not 1 <= number <= count:
        raise ValueError(f"There is no picture {number}: the document has {count} "
                         "(numbered from 1; word_list_pictures lists them; pictures "
                         "that float are in word_list_shapes).")
    return doc.InlineShapes.Item(number)


def _describe_picture(doc, shape, number: int) -> dict:
    return {"picture": number, "index": doc.Range(0, shape.Range.Start + 1).Paragraphs.Count,
            "width": round(shape.Width, 2), "height": round(shape.Height, 2),
            "alt_text": shape.AlternativeText or ""}


def _text_width(doc, rng) -> float:
    setup = rng.Sections.Item(1).PageSetup
    return round(setup.PageWidth - setup.LeftMargin - setup.RightMargin - setup.Gutter, 1)


def _apply_picture(doc, shape, width, height, alignment, alt_text) -> list[str]:
    width, height = proportional(width, height, shape.Width, shape.Height)
    if width is not None or height is not None:
        shape.LockAspectRatio = 0  # both sides are set, whatever they were
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height
    if alt_text is not None:
        shape.AlternativeText = alt_text
    if alignment is not None:
        shape.Range.ParagraphFormat.Alignment = ALIGNMENTS[alignment]
    return _too_wide(shape.Width, _text_width(doc, shape.Range))


def _insert_picture_impl(path, after, before, width, height, alignment, alt_text):
    from . import text

    if not os.path.isfile(path):
        raise ValueError(f"There is no file at '{path}'.")
    doc = backend.word.document()
    after, before = _place(doc.Paragraphs.Count, after, before)
    index = text._insert_paragraphs_impl([{"text": "", "style": "Normal"}], after,
                                         before)["inserted"][0]["index"]
    start = paragraph(doc, index).Range.Start
    shape = doc.InlineShapes.AddPicture(path, False, True, doc.Range(start, start))
    warnings = _apply_picture(doc, shape, width, height, alignment, alt_text)
    follow(doc, shape.Range)
    number = doc.Range(0, shape.Range.Start + 1).InlineShapes.Count
    result = {"success": True, **_describe_picture(doc, shape, number),
              "paragraphs": doc.Paragraphs.Count}
    if warnings:
        result["warnings"] = warnings
    return result


def _list_pictures_impl():
    doc = backend.word.document()
    return {"success": True, "pictures": [
        _describe_picture(doc, doc.InlineShapes.Item(n), n)
        for n in range(1, doc.InlineShapes.Count + 1)]}


def _set_picture_impl(picture, width, height, alignment, alt_text, spec):
    from .shapes import _check_name, adjust_picture, float_picture

    doc = backend.word.document()
    shape = _shape(doc, picture)
    if "wrap" in spec:
        _check_name(doc, spec.get("name"))
    if any(key in spec for key in PICTURE_LOOK):  # first: if refused, nothing else has changed
        adjust_picture(shape, spec)
    warnings = _apply_picture(doc, shape, width, height, alignment, alt_text)
    if "wrap" in spec:
        out = float_picture(doc, shape, {k: v for k, v in spec.items() if k not in PICTURE_LOOK})
        if warnings:
            out["warnings"] = [*warnings, *out.get("warnings", [])]
        return out
    follow(doc, shape.Range)
    result = {"success": True, **_describe_picture(doc, shape, picture),
              **faded(shape.PictureFormat.Brightness, shape.PictureFormat.Contrast)}
    if warnings:
        result["warnings"] = warnings
    return result


def _caption_where(doc, for_picture, for_table, after, before):
    from .tables import _span, _table

    if for_picture is not None:
        return doc.Range(0, _shape(doc, for_picture).Range.Start + 1).Paragraphs.Count, None
    if for_table is not None:
        first = _span(doc, _table(doc, for_table))["first"]
        if first == 1:
            raise ValueError("That table is the first thing in the document, so there is "
                             "no paragraph to put the caption after. Insert a paragraph "
                             "above it first, or caption it from below with `after`.")
        return first - 1, None
    return _place(doc.Paragraphs.Count, after, before)


def _insert_caption_impl(text_, label, separator, for_picture, for_table, after, before):
    from . import fields, text

    doc = backend.word.document()
    after, before = _caption_where(doc, for_picture, for_table, after, before)
    words = f"{label} X{separator}{text_}".replace("\n", "\v")
    index = text._insert_paragraphs_impl([{"text": words, "style": "Caption"}], after,
                                         before)["inserted"][0]["index"]
    at = paragraph(doc, index).Range.Start + _u16(label) + 1
    field = doc.Fields.Add(doc.Range(at, at + 1), fields.WD_FIELD_EMPTY,
                           f"SEQ {seq_name(label)} \\* ARABIC", False)
    field.Update()
    number = field.Result.Text
    fields._update_fields_impl("numbers", None, None)
    fields._update_fields_impl("references", None, None)
    # Later captions have renumbered, so the number is read again.
    p = paragraph(doc, index)
    number = _seq_in(p).Result.Text or number
    follow(doc, p.Range)
    return {"success": True, "index": index, "label": label, "number": number,
            "text": clean_text(p.Range.Text), "paragraphs": doc.Paragraphs.Count}


def _seq_in(p):
    from .fields import _kind

    found = p.Range.Fields
    for i in range(1, found.Count + 1):
        f = found.Item(i)
        if _kind(f.Code.Text) == "SEQ":
            return f
    return None


def _list_captions_impl():
    from .fields import _fields_in, _kind

    doc = backend.word.document()
    out = []
    for f in _fields_in(doc, None, None):
        code = f.Code.Text.split()
        if _kind(f.Code.Text) != "SEQ" or len(code) < 2:
            continue
        index = doc.Range(0, f.Code.Start + 1).Paragraphs.Count
        p = paragraph(doc, index)
        out.append({"index": index, "label": code[1], "number": f.Result.Text,
                    "equation": bool(p.Range.OMaths.Count), "text": clean_text(p.Range.Text)})
    return {"success": True, "captions": out}


def _insert_cross_reference_impl(index, to, show, position, match):
    from . import fields

    doc = backend.word.document()
    target = paragraph(doc, to)
    rng = target.Range
    if show in ("label_and_number", "number"):
        seq = _seq_in(target)
        if seq is None:
            raise ValueError(NO_NUMBER.format(index=to))
        end = seq.Result.End + 1  # past the field's end mark
        whole = show == "label_and_number" and not rng.OMaths.Count
        start = rng.Start if whole else seq.Code.Start - 1  # the field's begin mark
    else:
        start, end = rng.Start, rng.End - 1
    name = _bookmark_name()
    doc.Bookmarks.Add(name, doc.Range(start, end))
    result = fields._insert_field_impl(index, _reference_field(name, show), position, match)
    return {**result, "to": to, "show": show}


# -- Mac -------------------------------------------------------------------------------

def _mac_stage(path: str) -> tuple[str, bool]:
    """A copy inside Word's container, and whether one was made."""
    if os.path.realpath(path).startswith(os.path.realpath(CONTAINER) + os.sep):
        return path, False
    handle, staged = tempfile.mkstemp(prefix="ms-word-mcp-picture-",
                                      suffix=os.path.splitext(path)[1], dir=CONTAINER)
    os.close(handle)
    shutil.copyfile(path, staged)
    return staged, True


def _mac_shape(doc, number: int):
    from appscript import k

    count = int(doc.count(each=k.inline_shape))
    if not 1 <= number <= count:
        raise ValueError(f"There is no picture {number}: the document has {count} "
                         "(numbered from 1; word_list_pictures lists them; pictures "
                         "that float are in word_list_shapes).")
    return doc.inline_shapes[number]


def _mac_describe_picture(doc, shape, number: int) -> dict:
    from . import mac as m

    return {"picture": number, "index": m._index_at(doc, shape.text_object.start_of_content.get()),
            "width": round(shape.width.get(), 2), "height": round(shape.height.get(), 2),
            "alt_text": shape.alternative_text.get() or ""}


def _mac_apply_picture(doc, shape, width, height, alignment, alt_text) -> list[str]:
    from appscript import k

    from . import mac as m

    width, height = proportional(width, height, shape.width.get(), shape.height.get())
    # Setting one side does not keep the proportions even with the lock on
    # (measured), so both are set.
    if width is not None:
        shape.width.set(width)
    if height is not None:
        shape.height.set(height)
    if alt_text is not None:
        shape.alternative_text.set(alt_text)
    if alignment is not None:
        shape.text_object.paragraph_format.alignment.set(getattr(k, m.ALIGN_K[alignment]))
    setup = doc.page_setup
    return _too_wide(shape.width.get(), round(
        setup.page_width.get() - setup.left_margin.get() - setup.right_margin.get()
        - (setup.gutter.get() or 0), 1))


@mac_impl(_insert_picture_impl)
def _mac_insert_picture(path, after, before, width, height, alignment, alt_text):
    from appscript import k

    from . import mac as m
    from . import text
    from ._common import MAC_IMPLS

    if not os.path.isfile(path):
        raise ValueError(f"There is no file at '{path}'.")
    doc = backend.word.document()
    after, before = _place(m._count(doc), after, before)
    pictures = int(doc.count(each=k.inline_shape))
    index = MAC_IMPLS[text._insert_paragraphs_impl.__name__](
        [{"text": "", "style": "Normal"}], after, before)["inserted"][0]["index"]
    start, _ = m._bounds(doc, index)
    staged, copied = _mac_stage(path)
    try:
        doc.make(new=k.inline_picture, at=doc.create_range(start=start, end_=start),
                 with_properties={k.file_name: staged, k.link_to_file: False,
                                  k.save_with_document: True})
    finally:
        if copied:
            os.remove(staged)
    if int(doc.count(each=k.inline_shape)) != pictures + 1:
        raise RuntimeError("Word answered without error, but no picture was added.")
    number = next(n for n in range(1, pictures + 2)
                  if doc.inline_shapes[n].text_object.start_of_content.get() >= start)
    shape = doc.inline_shapes[number]
    warnings = _mac_apply_picture(doc, shape, width, height, alignment, alt_text)
    m._scroll(doc, start, start + 1)
    result = {"success": True, **_mac_describe_picture(doc, shape, number),
              "paragraphs": m._count(doc)}
    if warnings:
        result["warnings"] = warnings
    return result


@mac_impl(_list_pictures_impl)
def _mac_list_pictures():
    from appscript import k

    doc = backend.word.document()
    return {"success": True, "pictures": [
        _mac_describe_picture(doc, doc.inline_shapes[n], n)
        for n in range(1, int(doc.count(each=k.inline_shape)) + 1)]}


@mac_impl(_set_picture_impl)
def _mac_set_picture(picture, width, height, alignment, alt_text, spec):
    from .shapes import _mac_check_name, mac_adjust_picture, mac_float_picture

    doc = backend.word.document()
    shape = _mac_shape(doc, picture)
    if "wrap" in spec:
        _mac_check_name(doc, spec.get("name"))
    if any(key in spec for key in PICTURE_LOOK):  # first: if refused, nothing else has changed
        mac_adjust_picture(shape, spec)
    warnings = _mac_apply_picture(doc, shape, width, height, alignment, alt_text)
    if "wrap" in spec:
        out = mac_float_picture(doc, picture,
                                {k: v for k, v in spec.items() if k not in PICTURE_LOOK})
        if warnings:
            out["warnings"] = [*warnings, *out.get("warnings", [])]
        return out
    fmt = shape.picture_format
    result = {"success": True, **_mac_describe_picture(doc, shape, picture),
              **faded(fmt.brightness.get(), fmt.contrast.get())}
    if warnings:
        result["warnings"] = warnings
    return result


def _mac_caption_where(doc, for_picture, for_table, after, before):
    from . import mac as m
    from .tables import _mac_span, _mac_table

    if for_picture is not None:
        return m._index_at(doc, _mac_shape(doc, for_picture).text_object.start_of_content.get()), None
    if for_table is not None:
        first = _mac_span(doc, _mac_table(doc, for_table))["first"]
        if first == 1:
            raise ValueError("That table is the first thing in the document, so there is "
                             "no paragraph to put the caption after. Insert a paragraph "
                             "above it first, or caption it from below with `after`.")
        return first - 1, None
    return _place(m._count(doc), after, before)


def _mac_seq_in(doc, index):
    from . import mac as m
    from .fields import _kind, _mac_fields

    start, end = m._bounds(doc, index)
    for f in _mac_fields(doc):
        if start <= f.field_code.start_of_content.get() < end and _kind(
                f.field_code.content.get()) == "SEQ":
            return f
    return None


@mac_impl(_insert_caption_impl)
def _mac_insert_caption(text_, label, separator, for_picture, for_table, after, before):
    from . import fields
    from . import mac as m
    from . import text
    from ._common import MAC_IMPLS

    word = backend.word
    doc = word.document()
    after, before = _mac_caption_where(doc, for_picture, for_table, after, before)
    words = f"{label} X{separator}{text_}".replace("\n", "\v")
    index = MAC_IMPLS[text._insert_paragraphs_impl.__name__](
        [{"text": words, "style": "Caption"}], after, before)["inserted"][0]["index"]
    at = m._bounds(doc, index)[0] + _u16(label) + 1
    fields._mac_new_field(word.get_app(), doc, at, at + 1,
                          f"SEQ {seq_name(label)} \\* ARABIC").update_field()
    update = MAC_IMPLS[fields._update_fields_impl.__name__]
    update("numbers", None, None)
    update("references", None, None)
    seq = _mac_seq_in(doc, index)
    return {"success": True, "index": index, "label": label,
            "number": m._clean(seq.result_range.content.get()),
            "text": m._clean(doc.paragraphs[index].text_object.content.get()),
            "paragraphs": m._count(doc)}


@mac_impl(_list_captions_impl)
def _mac_list_captions():
    from appscript import k

    from . import mac as m
    from .fields import _kind, _mac_fields

    doc = backend.word.document()
    out = []
    for f in _mac_fields(doc):
        code = f.field_code.content.get() or ""
        parts = code.split()
        if _kind(code) != "SEQ" or len(parts) < 2:
            continue
        index = m._index_at(doc, f.field_code.start_of_content.get())
        rng = doc.paragraphs[index].text_object
        out.append({"index": index, "label": parts[1],
                    "number": m._clean(f.result_range.content.get()),
                    "equation": int(rng.count(each=k.math_object)) > 0,
                    "text": m._clean(rng.content.get())})
    return {"success": True, "captions": out}


@mac_impl(_insert_cross_reference_impl)
def _mac_insert_cross_reference(index, to, show, position, match):
    from appscript import k

    from . import fields
    from . import mac as m
    from ._common import MAC_IMPLS

    doc = backend.word.document()
    m._para(doc, to)
    start, end = m._bounds(doc, to)
    if show in ("label_and_number", "number"):
        seq = _mac_seq_in(doc, to)
        if seq is None:
            raise ValueError(NO_NUMBER.format(index=to))
        stop = seq.result_range.end_of_content.get() + 1  # past the field's end mark
        equation = int(doc.paragraphs[to].text_object.count(each=k.math_object)) > 0
        begin = start if show == "label_and_number" and not equation else (
            seq.field_code.start_of_content.get() - 1)  # the field's begin mark
    else:
        begin, stop = start, end - 1
    name = _bookmark_name()
    doc.make(new=k.bookmark, at=doc, with_properties={
        k.name: name, k.text_object: doc.create_range(start=begin, end_=stop)})
    result = MAC_IMPLS[fields._insert_field_impl.__name__](
        index, _reference_field(name, show), position, match)
    return {**result, "to": to, "show": show}


# -- tools -----------------------------------------------------------------------------

def register(mcp) -> None:
    @mcp.tool(name="word_insert_picture", annotations=annotations("Insert a picture", WRITES))
    async def word_insert_picture(params: InsertPictureInput) -> str:
        """Insert a picture from a file in a paragraph of its own, after or
        before a paragraph (default: at the end), at a size in points. Answers
        its number and size."""
        p = params
        return await call(_insert_picture_impl, p.path, p.after, p.before, p.width, p.height,
                          p.alignment, p.alt_text, undo="Insert picture")

    @mcp.tool(name="word_list_pictures", annotations=annotations("List pictures", READ_ONLY))
    async def word_list_pictures(params: NoInput) -> str:
        """The pictures in the text: number, paragraph, size and alt text."""
        return await call(_list_pictures_impl)

    @mcp.tool(name="word_set_picture", annotations=annotations("Change a picture", WRITES))
    async def word_set_picture(params: SetPictureInput) -> str:
        """Change a picture in the text: size, alignment, alt text, crop, brightness
        and contrast. Give
        `wrap` to take it out of the line of text so it floats: square, tight,
        through or top_and_bottom for text round it, behind for behind the text
        (背面), front for over it, with left, top, horizontal_from and
        vertical_from for its place. It answers the name word_set_shape takes
        from then on. To delete a picture in the text, delete its paragraph."""
        p = params
        return await call(_set_picture_impl, p.picture, p.width, p.height, p.alignment,
                          p.alt_text,
                          p.model_dump(exclude_none=True, exclude={
                              "picture", "width", "height", "alignment", "alt_text"}),
                          undo="Change picture")

    @mcp.tool(name="word_insert_caption", annotations=annotations("Insert a caption", WRITES))
    async def word_insert_caption(params: InsertCaptionInput) -> str:
        """Insert a numbered caption ("Figure 1 ...", "表 1 ...") in the Caption
        style, under a picture, above a table, or at a paragraph. Each label
        counts on its own, and the numbers renumber themselves."""
        p = params
        return await call(_insert_caption_impl, p.text, p.label, p.separator, p.for_picture,
                          p.for_table, p.after, p.before, undo="Insert caption")

    @mcp.tool(name="word_list_captions", annotations=annotations("List captions", READ_ONLY))
    async def word_list_captions(params: NoInput) -> str:
        """The captions and numbered equations: paragraph, label, number and
        text, to refer to with word_insert_cross_reference."""
        return await call(_list_captions_impl)

    @mcp.tool(name="word_insert_cross_reference",
              annotations=annotations("Insert a cross-reference", WRITES))
    async def word_insert_cross_reference(params: CrossReferenceInput) -> str:
        """Insert a live reference to a caption, numbered equation or heading:
        its label and number, number, text or page. It follows its target
        when the numbers change and word_update_fields is run."""
        p = params
        return await call(_insert_cross_reference_impl, p.index, p.to, p.show, p.position,
                          p.match, undo="Insert cross-reference")
