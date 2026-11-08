"""The page: its setup, vertical text, headers and footers, page numbers, ruby.

Setting a section's text vertical turns its page on its side as well
(measured on both platforms: A4 portrait becomes 842 x 595), so unless
`orientation` says otherwise the page is put back as it was, which keeps the
paper the user chose. Text in a header or footer is written whole, its {page}
and {pages} becoming PAGE and NUMPAGES fields, last first so the characters
before them stay where they are.

Ruby is Range.PhoneticGuide
(https://learn.microsoft.com/en-us/office/vba/api/word.range.phoneticguide),
which Word stores as an EQ field: the text it covers then reads as a single
\\x15 on both platforms, so positions are found again by reading Word.
"""

from __future__ import annotations

import re
from contextlib import contextmanager
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, Input, annotations, call, error_text, follow, mac_impl, mac_with_active
from ._word import (
    ALIGNMENTS, ALIGNMENT_NAMES, clean_text, locate_spans, paragraph, paragraph_info,
    ruby_to_drop)

WD_FIELD_EMPTY = -1
ORIENTATIONS = {"portrait": 0, "landscape": 1}  # WdOrientation
PAPERS = {"a3": 6, "a4": 7, "a5": 9, "b4": 10, "b5": 11, "letter": 2, "legal": 4}  # WdPaperSize
GRIDS = {"none": 0, "lines_and_characters": 1, "lines": 2, "snap_to_characters": 3}  # WdLayoutMode
DIRECTIONS = {"horizontal": 0, "vertical": 1}  # WdTextOrientation
WHICH = {"primary": 1, "first": 2, "even": 3}  # WdHeaderFooterIndex
NUMBER_STYLES = {  # WdPageNumberStyle
    "arabic": 0, "uppercase_roman": 1, "lowercase_roman": 2, "uppercase_letter": 3,
    "lowercase_letter": 4, "kanji": 10, "kanji_digit": 11, "arabic_full_width": 14,
    "number_in_circle": 18, "number_in_dash": 57}
RUBY_ALIGNMENTS = {"center": 0, "zero_one_zero": 1, "one_two_one": 2, "left": 3,
                   "right": 4}  # WdPhoneticGuideAlignmentType

_NAMES = {name: {v: k for k, v in table.items()}
          for name, table in (("orientation", ORIENTATIONS), ("paper", PAPERS), ("grid", GRIDS),
                              ("number_style", NUMBER_STYLES))}

MAC_GRIDS = {"none": "layout_mode_default", "lines_and_characters": "layout_mode_grid",
             "lines": "layout_mode_line_grid", "snap_to_characters": "layout_mode_genko"}
MAC_WHICH = {"primary": "header_footer_primary", "first": "header_footer_first_page",
             "even": "header_footer_even_pages"}


class SetPageSetupInput(Input):
    section: int | None = Field(None, ge=1, description=(
        "One section, numbered from 1. Default: every section."))
    paper: Literal[tuple(PAPERS)] | None = Field(None, description="A paper size.")  # type: ignore[valid-type]
    width: float | None = Field(None, gt=0, description="Or the page's width in points.")
    height: float | None = Field(None, gt=0, description="And its height in points.")
    orientation: Literal["portrait", "landscape"] | None = None
    text_direction: Literal["horizontal", "vertical"] | None = Field(None, description=(
        "vertical: East Asian vertical text (lines run top to bottom, right to left). "
        "The page keeps its orientation unless `orientation` is given too."))
    top_margin: float | None = Field(None, ge=0, description="In points.")
    bottom_margin: float | None = Field(None, ge=0, description="In points.")
    left_margin: float | None = Field(None, ge=0, description="In points.")
    right_margin: float | None = Field(None, ge=0, description="In points.")
    header_distance: float | None = Field(None, ge=0, description=(
        "From the page's edge to the header, in points."))
    footer_distance: float | None = Field(None, ge=0, description=(
        "From the page's edge to the footer, in points."))
    columns: int | None = Field(None, ge=1, le=45, description="Columns of text.")
    grid: Literal[tuple(GRIDS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "The document grid, for East Asian layouts: none, lines (a set number of "
        "lines a page), lines_and_characters (and characters a line), "
        "snap_to_characters (text sits on the characters of the grid)."))
    lines_per_page: float | None = Field(None, gt=0, description="Lines a page in the grid.")
    characters_per_line: float | None = Field(None, gt=0, description=(
        "Characters a line in the grid."))
    different_first_page: bool | None = Field(None, description=(
        "The first page has its own header and footer, both blank until written "
        "with word_set_header_footer `which` first."))
    different_odd_even: bool | None = Field(None, description=(
        "Even pages have their own header and footer, both blank until written "
        "with `which` even."))

    @model_validator(mode="after")
    def _something(self):
        if all(v is None for k, v in self.__dict__.items() if k != "section"):
            raise ValueError("Give at least one setting to change.")
        if self.paper is not None and (self.width is not None or self.height is not None):
            raise ValueError("Give paper, or width and height, not both.")
        return self


class SetHeaderFooterInput(Input):
    part: Literal["header", "footer"] = Field(..., description="Header or footer.")
    text: str = Field(..., description=(
        "Its whole text, or empty to clear it. {page} becomes the page number and "
        "{pages} the number of pages: \"Page {page} of {pages}\". A tab moves to the "
        "next tab stop of the Header or Footer style, by default the centre then "
        "the right."))
    which: Literal["primary", "first", "even"] = Field("primary", description=(
        "primary: every page. first: the first page, which then has its own header "
        "and footer. even: even pages, which then have their own. Once the first or "
        "even pages have their own, the other part (header or footer) needs writing "
        "for them too, or they are blank there."))
    section: int = Field(1, ge=1, description=(
        "The section. Later sections show the same until a section has its own; "
        "writing to a later section gives it its own."))
    alignment: Literal["left", "center", "right"] | None = None


class SetPageNumbersInput(Input):
    section: int = Field(1, ge=1, description="The section, numbered from 1.")
    style: Literal[tuple(NUMBER_STYLES)] | None = Field(None, description=(  # type: ignore[valid-type]
        "How numbers are written: arabic 1, lowercase_roman i, kanji 一, "
        "arabic_full_width 1, number_in_circle ①, number_in_dash - 1 -..."))
    start_at: int | None = Field(None, ge=0, description=(
        "Start this section's numbering at this number."))
    continue_previous: bool | None = Field(None, description=(
        "True: go on from the section before instead of restarting."))

    @model_validator(mode="after")
    def _something(self):
        if self.style is None and self.start_at is None and self.continue_previous is None:
            raise ValueError("Give style, start_at or continue_previous.")
        if self.start_at is not None and self.continue_previous:
            raise ValueError("Give start_at, or continue_previous, not both.")
        return self


class RubySpan(BaseModel):
    model_config = ConfigDict(extra="forbid")

    text: str = Field(..., min_length=1, max_length=255, description=(
        "The text to put ruby on, as it stands in the paragraph. Text that already "
        "has ruby takes the new reading."))
    ruby: str = Field(..., min_length=1, max_length=255, description="The reading above it.")


class AddRubyInput(Input):
    index: int = Field(..., ge=1, description="The paragraph.")
    spans: list[RubySpan] = Field(..., min_length=1, max_length=100, description=(
        "Each piece of text and its reading, in the order they come in the "
        "paragraph: [{\"text\": \"漢字\", \"ruby\": \"かんじ\"}]. Each entry covers one "
        "occurrence: to cover a word that comes twice, list it twice."))
    alignment: Literal[tuple(RUBY_ALIGNMENTS)] = Field("center", description=(  # type: ignore[valid-type]
        "How the reading sits over the text."))
    size: int | None = Field(None, ge=1, le=127, description=(
        "The reading's size in points. Default: half the text's."))


class NoInput(Input):
    pass


def parse_placeholders(text: str) -> tuple[str, list[tuple[int, str]]]:
    """The text with each {page}/{pages} as one character, and (offset, field code)."""
    out, fields = [], []
    length = 0
    for piece in re.split(r"(\{pages?\})", text.replace("\r\n", "\n").replace("\n", "\v")):
        if piece in ("{page}", "{pages}"):
            fields.append((length, "PAGE" if piece == "{page}" else "NUMPAGES"))
            out.append("#")
            length += 1
        else:
            out.append(piece)
            length += len(piece)
    return "".join(out), fields


def _attempt(warnings: list[str], name: str, action) -> None:
    """One setting; one Word refuses goes in the warnings, and the rest still land."""
    try:
        action()
    except Exception as exc:  # noqa: BLE001 - reported
        warnings.append(f"{name} was not set: {error_text(exc).splitlines()[0][:160]}")


# -- Windows -----------------------------------------------------------------------

def _sections(doc, section) -> list[int]:
    count = doc.Sections.Count
    if section is None:
        return list(range(1, count + 1))
    if section > count:
        raise ValueError(f"There is no section {section}: the document has {count}.")
    return [section]


def _describe_section(doc, n: int) -> dict:
    sec = doc.Sections.Item(n)
    ps = sec.PageSetup
    return {
        "section": n, "paper": _NAMES["paper"].get(ps.PaperSize, "custom"),
        "width": round(ps.PageWidth, 1), "height": round(ps.PageHeight, 1),
        "orientation": _NAMES["orientation"].get(ps.Orientation),
        "text_direction": "vertical" if sec.Range.Orientation in (1, 5) else "horizontal",
        "margins": {"top": round(ps.TopMargin, 1), "bottom": round(ps.BottomMargin, 1),
                    "left": round(ps.LeftMargin, 1), "right": round(ps.RightMargin, 1)},
        "header_distance": round(ps.HeaderDistance, 1),
        "footer_distance": round(ps.FooterDistance, 1), "columns": ps.TextColumns.Count,
        "grid": _NAMES["grid"].get(ps.LayoutMode), "lines_per_page": ps.LinesPage,
        "characters_per_line": ps.CharsLine,
        "different_first_page": bool(ps.DifferentFirstPageHeaderFooter),
        "different_odd_even": bool(ps.OddAndEvenPagesHeaderFooter)}


def _get_page_setup_impl():
    doc = backend.word.document()
    return {"success": True, "sections": [_describe_section(doc, n)
                                          for n in range(1, doc.Sections.Count + 1)]}


def _set_page_setup_impl(section, spec):
    doc = backend.word.document()
    numbers = _sections(doc, section)
    warnings: list[str] = []
    for n in numbers:
        sec = doc.Sections.Item(n)
        ps = sec.PageSetup

        def put(name, attr, value):
            if value is not None:
                _attempt(warnings, f"{name} (section {n})", lambda: setattr(ps, attr, value))

        tables, kept = [], []
        put("paper", "PaperSize", PAPERS.get(spec.get("paper")))
        put("width", "PageWidth", spec.get("width"))
        put("height", "PageHeight", spec.get("height"))
        if spec.get("text_direction"):
            was = ps.Orientation
            # Changing the direction floats the section's tables, with text round
            # them, and sends them to the left (measured). Vertical text keeps
            # them floating; back in horizontal text they are put in line again.
            tables = [sec.Range.Tables.Item(i) for i in range(1, sec.Range.Tables.Count + 1)]
            kept = [t.Rows.Alignment for t in tables]
            _attempt(warnings, f"text_direction (section {n})", lambda: setattr(
                sec.Range, "Orientation", DIRECTIONS[spec["text_direction"]]))
            if not spec.get("orientation") and ps.Orientation != was:
                ps.Orientation = was
        put("orientation", "Orientation", ORIENTATIONS.get(spec.get("orientation")))
        for t, alignment in zip(tables, kept):  # after the page has turned too
            if spec["text_direction"] == "horizontal":
                t.Rows.WrapAroundText = False  # before the alignment, which it resets
            t.Rows.Alignment = alignment
        put("top_margin", "TopMargin", spec.get("top_margin"))
        put("bottom_margin", "BottomMargin", spec.get("bottom_margin"))
        put("left_margin", "LeftMargin", spec.get("left_margin"))
        put("right_margin", "RightMargin", spec.get("right_margin"))
        put("header_distance", "HeaderDistance", spec.get("header_distance"))
        put("footer_distance", "FooterDistance", spec.get("footer_distance"))
        if spec.get("columns"):
            _attempt(warnings, f"columns (section {n})",
                     lambda: ps.TextColumns.SetCount(spec["columns"]))
        put("grid", "LayoutMode", GRIDS.get(spec.get("grid")))
        put("lines_per_page", "LinesPage", spec.get("lines_per_page"))
        put("characters_per_line", "CharsLine", spec.get("characters_per_line"))
        if spec.get("different_first_page") is not None:
            put("different_first_page", "DifferentFirstPageHeaderFooter",
                bool(spec["different_first_page"]))
        if spec.get("different_odd_even") is not None:
            put("different_odd_even", "OddAndEvenPagesHeaderFooter",
                bool(spec["different_odd_even"]))
    result = {"success": True, "sections": [_describe_section(doc, n) for n in numbers]}
    if warnings:
        result["warnings"] = warnings
    return result


def _header_footer(sec, part, which):
    return (sec.Headers if part == "header" else sec.Footers).Item(WHICH[which])


def _set_header_footer_impl(part, text, which, section, alignment):
    doc = backend.word.document()
    _sections(doc, section)
    sec = doc.Sections.Item(section)
    if which == "first":
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    elif which == "even":
        sec.PageSetup.OddAndEvenPagesHeaderFooter = True
    hf = _header_footer(sec, part, which)
    result = {"success": True, "part": part, "which": which, "section": section}
    if section > 1 and hf.LinkToPrevious:
        hf.LinkToPrevious = False
        result["note"] = f"Section {section} now has its own {part}."
    plain, fields = parse_placeholders(text)
    hf.Range.Text = plain
    for offset, code in reversed(fields):
        doc.Fields.Add(hf.Range.Characters.Item(offset + 1), WD_FIELD_EMPTY, code, False)
    if alignment is not None:
        hf.Range.ParagraphFormat.Alignment = ALIGNMENTS[alignment]
    result["reads"] = clean_text(hf.Range.Text)
    return result


def _hf_info(hf) -> dict:
    return {"text": clean_text(hf.Range.Text),
            "alignment": ALIGNMENT_NAMES.get(hf.Range.ParagraphFormat.Alignment),
            "linked_to_previous": bool(hf.LinkToPrevious)}


def _get_headers_footers_impl():
    doc = backend.word.document()
    out = []
    for n in range(1, doc.Sections.Count + 1):
        sec = doc.Sections.Item(n)
        ps = sec.PageSetup
        kinds = ["primary"]
        if ps.DifferentFirstPageHeaderFooter:
            kinds.append("first")
        if ps.OddAndEvenPagesHeaderFooter:
            kinds.append("even")
        numbers = sec.Footers.Item(1).PageNumbers
        out.append({"section": n,
                    "header": {w: _hf_info(_header_footer(sec, "header", w)) for w in kinds},
                    "footer": {w: _hf_info(_header_footer(sec, "footer", w)) for w in kinds},
                    "page_numbers": {
                        "style": _NAMES["number_style"].get(numbers.NumberStyle,
                                                            numbers.NumberStyle),
                        "restart": bool(numbers.RestartNumberingAtSection),
                        "start_at": numbers.StartingNumber}})
    return {"success": True, "sections": out}


def _set_page_numbers_impl(section, style, start_at, continue_previous):
    doc = backend.word.document()
    _sections(doc, section)
    numbers = doc.Sections.Item(section).Footers.Item(1).PageNumbers
    if style is not None:
        numbers.NumberStyle = NUMBER_STYLES[style]
    if start_at is not None:
        numbers.RestartNumberingAtSection = True
        numbers.StartingNumber = start_at
    elif continue_previous is not None:
        numbers.RestartNumberingAtSection = not continue_previous
    return {"success": True, "section": section,
            "style": _NAMES["number_style"].get(numbers.NumberStyle, numbers.NumberStyle),
            "restart": bool(numbers.RestartNumberingAtSection),
            "start_at": numbers.StartingNumber,
            "note": "Page numbers show where {page} is in a header or footer "
                    "(word_set_header_footer)."}


@contextmanager
def _saying_what_was_taken_off(words: list[str]):
    """An error after old readings came off says so: it is not "nothing changed"."""
    try:
        yield
    except Exception as exc:
        if not words:
            raise
        said = error_text(exc).replace("Nothing was changed.", "").strip()
        raise ValueError(f"{said} The readings on {', '.join(words)} had been taken off by "
                         "then: read the paragraph, and add them again.") from exc


def _drop_ruby(doc, p, spans) -> list[str]:
    """Turn the ruby of the spans' text, if it has any, back into plain text.

    Refuses first, with nothing taken off, when the spans do not fit the
    paragraph. Answers the words whose readings came off.
    """
    rng = p.Range
    fields = rng.Fields
    items = [fields.Item(i) for i in range(1, fields.Count + 1)]
    found = ruby_to_drop(clean_text(rng.Text), [f.Code.Text for f in items],
                         [s["text"] for s in spans])
    for i, base in reversed(found):
        at = items[i].Code.Start - 1  # the field's begin mark
        items[i].Delete()
        doc.Range(at, at).InsertAfter(base)
    return [base for _, base in found]


def _add_ruby_impl(index, spans, alignment, size):
    doc = backend.word.document()
    p = paragraph(doc, index)
    taken_off = _drop_ruby(doc, p, spans)
    with _saying_what_was_taken_off(taken_off):
        p = paragraph(doc, index)
        rng = p.Range
        raw = rng.Text or ""
        hidden = max(0, rng.End - rng.Start - len(raw.encode("utf-16-le")) // 2)
        places = locate_spans(raw, [s["text"] for s in spans],
                              lambda a, b: doc.Range(a, b).Text, rng.Start, hidden)
        app = backend.word.get_app()
        try:
            was = app.ActiveDocument
        except Exception:  # noqa: BLE001 - no active document
            was = None
        # Last first, so the positions before it hold. PhoneticGuide makes the
        # document the active one (measured), so the user's is given back.
        try:
            for (start, end), span in zip(reversed(places), reversed(spans)):
                target = doc.Range(start, end)
                if size is None:
                    target.PhoneticGuide(span["ruby"], RUBY_ALIGNMENTS[alignment])
                else:
                    target.PhoneticGuide(span["ruby"], RUBY_ALIGNMENTS[alignment], 0, size)
        finally:
            if was is not None and was.FullName != doc.FullName:
                was.Activate()
    p = paragraph(doc, index)
    follow(doc, p.Range)
    return {"success": True, **paragraph_info(index, p)}


# -- Mac -------------------------------------------------------------------------------

def _mac_sections(doc, section) -> list[int]:
    from appscript import k

    count = int(doc.count(each=k.section))
    if section is None:
        return list(range(1, count + 1))
    if section > count:
        raise ValueError(f"There is no section {section}: the document has {count}.")
    return [section]


def _kname(value) -> str:
    return getattr(value, "name", None) or str(value).removeprefix("k.")


def _mac_describe_section(doc, n: int) -> dict:
    from appscript import k

    sec = doc.sections[n]
    ps = sec.page_setup
    paper = _kname(ps.paper_size.get()).removeprefix("paper_")
    grid = _kname(ps.layout_mode.get())
    direction = _kname(sec.text_object.orientation.get())
    return {
        "section": n, "paper": paper if paper in PAPERS else "custom",
        "width": round(ps.page_width.get(), 1), "height": round(ps.page_height.get(), 1),
        "orientation": "landscape" if _kname(ps.orientation.get()) == "orient_landscape"
        else "portrait",
        "text_direction": "vertical" if "vertical" in direction else "horizontal",
        "margins": {"top": round(ps.top_margin.get(), 1),
                    "bottom": round(ps.bottom_margin.get(), 1),
                    "left": round(ps.left_margin.get(), 1),
                    "right": round(ps.right_margin.get(), 1)},
        "header_distance": round(ps.header_distance.get(), 1),
        "footer_distance": round(ps.footer_distance.get(), 1),
        "columns": int(ps.count(each=k.text_column)),
        "grid": next((g for g, name in MAC_GRIDS.items() if name == grid), grid),
        "lines_per_page": ps.lines_page.get(), "characters_per_line": ps.chars_line.get(),
        "different_first_page": bool(ps.different_first_page_header_footer.get()),
        "different_odd_even": bool(ps.odd_and_even_pages_header_footer.get())}


@mac_impl(_get_page_setup_impl)
def _mac_get_page_setup():
    from appscript import k

    doc = backend.word.document()
    return {"success": True, "sections": [
        _mac_describe_section(doc, n) for n in range(1, int(doc.count(each=k.section)) + 1)]}


@mac_impl(_set_page_setup_impl)
def _mac_set_page_setup(section, spec):
    from appscript import k

    doc = backend.word.document()
    numbers = _mac_sections(doc, section)
    warnings: list[str] = []
    for n in numbers:
        sec = doc.sections[n]
        ps = sec.page_setup

        def put(name, prop, value):
            if value is not None:
                _attempt(warnings, f"{name} (section {n})", lambda: prop.set(value))

        tables, kept = [], []
        if spec.get("paper"):
            put("paper", ps.paper_size, getattr(k, "paper_" + spec["paper"]))
        put("width", ps.page_width, spec.get("width"))
        put("height", ps.page_height, spec.get("height"))
        if spec.get("text_direction"):
            was = ps.orientation.get()
            value = (k.text_orientation_vertical_east_asian
                     if spec["text_direction"] == "vertical" else k.text_orientation_horizontal)
            # Changing the direction floats the section's tables, with text round
            # them, and sends them to the left (measured). Vertical text keeps
            # them floating; back in horizontal text they are put in line again.
            low = sec.text_object.start_of_content.get()
            high = sec.text_object.end_of_content.get()
            tables = [doc.tables[i] for i in range(1, int(doc.count(each=k.table)) + 1)
                      if low <= doc.tables[i].text_object.start_of_content.get() < high]
            kept = [t.row_options.alignment.get() for t in tables]
            _attempt(warnings, f"text_direction (section {n})",
                     lambda: sec.text_object.orientation.set(value))
            if not spec.get("orientation") and ps.orientation.get() != was:
                ps.orientation.set(was)
        if spec.get("orientation"):
            put("orientation", ps.orientation, getattr(k, "orient_" + spec["orientation"]))
        for t, alignment in zip(tables, kept):  # after the page has turned too
            if spec["text_direction"] == "horizontal":
                t.row_options.wrap_around_text.set(False)  # before the alignment, which it resets
            if alignment != k.missing_value:  # mixed: nothing one value can put back
                t.row_options.alignment.set(alignment)
        put("top_margin", ps.top_margin, spec.get("top_margin"))
        put("bottom_margin", ps.bottom_margin, spec.get("bottom_margin"))
        put("left_margin", ps.left_margin, spec.get("left_margin"))
        put("right_margin", ps.right_margin, spec.get("right_margin"))
        put("header_distance", ps.header_distance, spec.get("header_distance"))
        put("footer_distance", ps.footer_distance, spec.get("footer_distance"))
        if spec.get("columns"):
            _attempt(warnings, f"columns (section {n})", lambda: ps.set_number_of_text_columns(
                number_of_columns=spec["columns"]))
        if spec.get("grid"):
            put("grid", ps.layout_mode, getattr(k, MAC_GRIDS[spec["grid"]]))
        put("lines_per_page", ps.lines_page, spec.get("lines_per_page"))
        put("characters_per_line", ps.chars_line, spec.get("characters_per_line"))
        if spec.get("different_first_page") is not None:
            put("different_first_page", ps.different_first_page_header_footer,
                bool(spec["different_first_page"]))
        if spec.get("different_odd_even") is not None:
            put("different_odd_even", ps.odd_and_even_pages_header_footer,
                bool(spec["different_odd_even"]))
    result = {"success": True, "sections": [_mac_describe_section(doc, n) for n in numbers]}
    if warnings:
        result["warnings"] = warnings
    return result


def _mac_hf(sec, part, which):
    from appscript import k

    get = sec.get_header if part == "header" else sec.get_footer
    return get(index=getattr(k, MAC_WHICH[which]))


@mac_impl(_set_header_footer_impl)
def _mac_set_header_footer(part, text, which, section, alignment):
    from appscript import k

    from . import mac as m

    word = backend.word
    app = word.get_app()
    doc = word.document()
    _mac_sections(doc, section)
    sec = doc.sections[section]
    if which == "first":
        sec.page_setup.different_first_page_header_footer.set(True)
    elif which == "even":
        sec.page_setup.odd_and_even_pages_header_footer.set(True)
    hf = _mac_hf(sec, part, which)
    result = {"success": True, "part": part, "which": which, "section": section}
    if section > 1 and hf.link_to_previous.get():
        hf.link_to_previous.set(False)
        result["note"] = f"Section {section} now has its own {part}."
    plain, fields = parse_placeholders(text)
    hf.text_object.content.set(plain)
    for offset, code in reversed(fields):
        target = hf.text_object.characters[offset + 1]
        mac_with_active(app, doc, lambda c=code, t=target: app.create_new_field(
            text_range=t, field_type=k.field_empty, field_text=c, preserve_formatting=False))
    if alignment is not None:
        hf.text_object.paragraph_format.alignment.set(getattr(k, m.ALIGN_K[alignment]))
    result["reads"] = m._clean(hf.text_object.content.get())
    return result


def _mac_hf_info(hf) -> dict:
    from . import mac as m

    return {"text": m._clean(hf.text_object.content.get()),
            "alignment": m._from_k(m.ALIGN_K, hf.text_object.paragraph_format.alignment.get()),
            "linked_to_previous": bool(hf.link_to_previous.get())}


def _mac_number_style(value):
    name = _kname(value).removeprefix("page_number_style_")
    return name


@mac_impl(_get_headers_footers_impl)
def _mac_get_headers_footers():
    from appscript import k

    doc = backend.word.document()
    out = []
    for n in range(1, int(doc.count(each=k.section)) + 1):
        sec = doc.sections[n]
        ps = sec.page_setup
        kinds = ["primary"]
        if ps.different_first_page_header_footer.get():
            kinds.append("first")
        if ps.odd_and_even_pages_header_footer.get():
            kinds.append("even")
        numbers = _mac_hf(sec, "footer", "primary").page_number_options
        out.append({"section": n,
                    "header": {w: _mac_hf_info(_mac_hf(sec, "header", w)) for w in kinds},
                    "footer": {w: _mac_hf_info(_mac_hf(sec, "footer", w)) for w in kinds},
                    "page_numbers": {
                        "style": _mac_number_style(numbers.number_style.get()),
                        "restart": bool(numbers.restart_numbering_at_section.get()),
                        "start_at": numbers.starting_number.get()}})
    return {"success": True, "sections": out}


@mac_impl(_set_page_numbers_impl)
def _mac_set_page_numbers(section, style, start_at, continue_previous):
    from appscript import k

    doc = backend.word.document()
    _mac_sections(doc, section)
    numbers = _mac_hf(doc.sections[section], "footer", "primary").page_number_options
    if style is not None:
        numbers.number_style.set(getattr(k, "page_number_style_" + style))
    if start_at is not None:
        numbers.restart_numbering_at_section.set(True)
        numbers.starting_number.set(start_at)
    elif continue_previous is not None:
        numbers.restart_numbering_at_section.set(not continue_previous)
    return {"success": True, "section": section,
            "style": _mac_number_style(numbers.number_style.get()),
            "restart": bool(numbers.restart_numbering_at_section.get()),
            "start_at": numbers.starting_number.get(),
            "note": "Page numbers show where {page} is in a header or footer "
                    "(word_set_header_footer)."}


def _mac_drop_ruby(doc, index, spans) -> list[str]:
    from appscript import k

    rng = doc.paragraphs[index].text_object
    codes = [rng.fields[i].field_code.content.get()
             for i in range(1, int(rng.count(each=k.field)) + 1)]
    found = ruby_to_drop(clean_text(rng.content.get()), codes, [s["text"] for s in spans])
    for i, base in reversed(found):
        at = rng.fields[i + 1].field_code.start_of_content.get() - 1  # the field's begin mark
        rng.fields[i + 1].delete()
        doc.create_range(start=at, end_=at).content.set(base)
    return [base for _, base in found]


@mac_impl(_add_ruby_impl)
def _mac_add_ruby(index, spans, alignment, size):
    from appscript import k

    from . import mac as m

    word = backend.word
    app = word.get_app()
    doc = word.document()
    m._para(doc, index)
    taken_off = _mac_drop_ruby(doc, index, spans)
    with _saying_what_was_taken_off(taken_off):
        start, end = m._bounds(doc, index)
        raw = doc.paragraphs[index].text_object.content.get() or ""
        hidden = max(0, end - start - len(raw.encode("utf-16-le")) // 2)

        def read(a, b):
            try:
                return doc.create_range(start=a, end_=b).content.get()
            except Exception:  # noqa: BLE001 - half a character: not here
                return None

        places = locate_spans(raw, [s["text"] for s in spans], read, start, hidden)
        value = getattr(k, "phonetic_guide_alignment_" + {
            "center": "center", "zero_one_zero": "zero_one_zero", "one_two_one": "one_two_one",
            "left": "left", "right": "right"}[alignment])
        for (a, b), span in zip(reversed(places), reversed(spans)):
            target = doc.create_range(start=a, end_=b)
            points = size or max(1, round(float(target.font_object.font_size.get()) / 2))
            # `raise` 0 lets Word place the reading; every parameter is required here.
            mac_with_active(app, doc, lambda t=target, s=span, p=points: t.phonetic_guide(
                text=s["ruby"], alignment_type=value, raise_=0, font_size=p, font_name=""))
    m._scroll(doc, start, end)
    return {"success": True, **m._info(doc, index)}


# -- tools -----------------------------------------------------------------------------

def register(mcp) -> None:
    @mcp.tool(name="word_get_page_setup", annotations=annotations("Page setup", READ_ONLY))
    async def word_get_page_setup(params: NoInput) -> str:
        """Each section's paper, orientation, margins, text direction, columns
        and document grid."""
        return await call(_get_page_setup_impl)

    @mcp.tool(name="word_set_page_setup", annotations=annotations("Set page setup", WRITES))
    async def word_set_page_setup(params: SetPageSetupInput) -> str:
        """Set the paper, orientation, margins, vertical text (text_direction),
        columns, document grid, and whether the first or even pages have their
        own headers, for one section or all."""
        p = params.model_dump(exclude={"section"}, exclude_none=True)
        return await call(_set_page_setup_impl, params.section, p, undo="Set page setup")

    @mcp.tool(name="word_get_headers_footers",
              annotations=annotations("Headers and footers", READ_ONLY))
    async def word_get_headers_footers(params: NoInput) -> str:
        """Each section's headers and footers and how its pages are numbered."""
        return await call(_get_headers_footers_impl)

    @mcp.tool(name="word_set_header_footer",
              annotations=annotations("Set a header or footer", WRITES))
    async def word_set_header_footer(params: SetHeaderFooterInput) -> str:
        """Write a header or footer, with {page} and {pages} for the page
        number and the number of pages."""
        p = params
        return await call(_set_header_footer_impl, p.part, p.text, p.which, p.section,
                          p.alignment, undo="Set header or footer")

    @mcp.tool(name="word_set_page_numbers", annotations=annotations("Set page numbering", WRITES))
    async def word_set_page_numbers(params: SetPageNumbersInput) -> str:
        """Set how a section's pages are numbered: the style of the numbers and
        where they start. The number itself is {page} in a header or footer."""
        p = params
        return await call(_set_page_numbers_impl, p.section, p.style, p.start_at,
                          p.continue_previous, undo="Set page numbering")

    @mcp.tool(name="word_add_ruby", annotations=annotations("Add ruby", WRITES))
    async def word_add_ruby(params: AddRubyInput) -> str:
        """Add ruby (furigana, phonetic guides) above pieces of a paragraph's
        text. Word reads such text back with its readings in `ruby`."""
        p = params
        return await call(_add_ruby_impl, p.index, [s.model_dump() for s in p.spans],
                          p.alignment, p.size, undo="Add ruby")
