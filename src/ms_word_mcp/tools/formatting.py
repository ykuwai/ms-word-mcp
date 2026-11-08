"""Styles, character formatting and paragraph formatting.

Word keeps three fonts on every run of text: one for Latin text, one for
East Asian text (Japanese, Chinese, Korean) and one for complex scripts
(Arabic, Hebrew and other right-to-left text). Setting only one leaves the
others as they were, which is why Japanese text so often stays in the wrong
font. These tools take the three separately and read all three back.
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, Input, annotations, call, error_text, follow
from ._word import (
    ALIGNMENTS, LINE_SPACING, WD_COLLAPSE_END, WD_FIND_STOP, check_span, clean_text,
    flag, fonts_of, paragraph_info, paragraphs_in, resolve_style, spacing_in_lines,
    span_range, style_name, undefined,
)


class Target(Input):
    """Which text a formatting call works on."""

    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")
    match: str | None = Field(None, min_length=1, max_length=255, description=(
        "Only the occurrences of this text inside those paragraphs, rather than "
        "the paragraphs whole."))


def _target_ranges(doc, start: int, end: int | None, match: str | None):
    start, end = check_span(doc, start, end)
    whole = span_range(doc, start, end)
    if match is None:
        return [whole]
    found = []
    rng = doc.Range(whole.Start, whole.End)
    f = rng.Find
    f.ClearFormatting()
    while f.Execute(FindText=match, MatchCase=True, Forward=True, Wrap=WD_FIND_STOP,
                    Format=False):
        if rng.End > whole.End:
            break
        found.append(doc.Range(rng.Start, rng.End))
        rng.Collapse(WD_COLLAPSE_END)
    if not found:
        raise ValueError(f"'{match}' does not occur in paragraphs {start} to {end}. "
                         "Nothing was changed.")
    return found


# -- character formatting -------------------------------------------------------

class FormatTextInput(Target):
    latin_font: str | None = Field(None, description=(
        "Font for Latin text (alphabet, digits), e.g. \"Segoe UI\"."))
    east_asian_font: str | None = Field(None, description=(
        "Font for East Asian text, e.g. \"BIZ UDPゴシック\" or \"游明朝\". Word ignores "
        "a font without East Asian glyphs here; the read-back says so."))
    complex_script_font: str | None = Field(None, description=(
        "Font for complex scripts: Arabic, Hebrew and other right-to-left text."))
    size: float | None = Field(None, gt=0, le=1638, description=(
        "Size in points for Latin and East Asian text."))
    size_complex_script: float | None = Field(None, gt=0, le=1638, description=(
        "Size in points for complex-script text."))
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    color: str | None = Field(None, pattern=r"^#?[0-9A-Fa-f]{6}$", description=(
        "Text color as #RRGGBB."))

    @model_validator(mode="after")
    def _something(self):
        fields = ("latin_font", "east_asian_font", "complex_script_font", "size",
                  "size_complex_script", "bold", "italic", "underline", "color")
        if all(getattr(self, f) is None for f in fields):
            raise ValueError("Give at least one formatting argument: " + ", ".join(fields) + ".")
        return self


def _bgr(hex_color: str) -> int:
    value = hex_color.lstrip("#")
    r, g, b = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
    return r | (g << 8) | (b << 16)


def _hex(bgr) -> str | None:
    if bgr is None or bgr < 0 or bgr > 0xFFFFFF:
        return None  # automatic, or mixed
    r, g, b = bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF
    return f"#{r:02X}{g:02X}{b:02X}"


def _format_text_impl(start, end, match, latin_font, east_asian_font,
                      complex_script_font, size, size_complex_script, bold, italic,
                      underline, color):
    doc = backend.word.document()
    ranges = _target_ranges(doc, start, end, match)
    refused = {}
    for rng in ranges:
        font = rng.Font
        # Word refuses a font name it will not use for that script (an East
        # Asian font without East Asian glyphs answers 0x800A16D4 on Windows
        # and -1708 on the Mac) and keeps the old one. The rest of the call
        # still applies; the refusal comes back as a warning.
        for script, members, wanted in (
                ("latin", ("NameAscii", "NameOther"), latin_font),
                ("east_asian", ("NameFarEast",), east_asian_font),
                ("complex_script", ("NameBi",), complex_script_font)):
            if wanted is None:
                continue
            for member in members:
                try:
                    setattr(font, member, wanted)
                except Exception as exc:  # noqa: BLE001 - reported below
                    refused[script] = error_text(exc)
        if size is not None:
            font.Size = size
        if size_complex_script is not None:
            font.SizeBi = size_complex_script
        if bold is not None:
            font.Bold = bold
            font.BoldBi = bold
        if italic is not None:
            font.Italic = italic
            font.ItalicBi = italic
        if underline is not None:
            font.Underline = 1 if underline else 0
        if color is not None:
            font.Color = _bgr(color)

    # Read back from Word, and say what did not take.
    warnings = []
    readback = []
    for rng in ranges:
        got = fonts_of(rng)
        font = rng.Font
        got.update({"bold": flag(font.Bold), "italic": flag(font.Italic),
                    "color": _hex(undefined(font.Color))})
        readback.append({"text": clean_text(rng.Text)[:80], **got})
        for key, wanted in (("latin", latin_font), ("east_asian", east_asian_font),
                            ("complex_script", complex_script_font)):
            if wanted is not None and got[key] != wanted:
                why = (" Word refused it" + (": an East Asian font needs East Asian "
                                             "glyphs" if key == "east_asian" else "")
                       + f" ({refused[key]})." if key in refused else "")
                warnings.append(
                    f"The {key.replace('_', ' ')} font reads back as {got[key]!r}, not "
                    f"{wanted!r}.{why}")
        if size is not None and got["size"] != size:
            warnings.append(f"The size reads back as {got['size']}, not {size}.")
    follow(doc, ranges[0])
    result = {"success": True, "ranges": len(ranges), "read_back": readback[:20]}
    if warnings:
        result["warnings"] = sorted(set(warnings))
    return result


# -- paragraph formatting --------------------------------------------------------

class ParagraphFormatInput(Input):
    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")
    alignment: Literal["left", "center", "right", "justify", "distribute"] | None = None
    line_spacing_rule: Literal["single", "one_and_half", "double", "at_least", "exactly",
                               "multiple"] | None = Field(None, description=(
        "How line_spacing is read: \"exactly\" and \"at_least\" take points, "
        "\"multiple\" takes lines (1.2 = 1.2 lines)."))
    line_spacing: float | None = Field(None, gt=0, le=1584, description=(
        "Points for exactly/at_least, lines for multiple."))
    space_before: float | None = Field(None, ge=0, le=1584, description="Points.")
    space_after: float | None = Field(None, ge=0, le=1584, description="Points.")
    first_line_indent: float | None = Field(None, ge=-1584, le=1584, description=(
        "Points; negative for a hanging indent."))
    left_indent: float | None = Field(None, ge=-1584, le=1584, description="Points.")
    right_indent: float | None = Field(None, ge=-1584, le=1584, description="Points.")
    snap_to_grid: bool | None = Field(None, description=(
        "Align lines to the document grid. Japanese templates turn it on, and it "
        "stretches line spacing: turn it off for exact spacing."))
    keep_with_next: bool | None = None
    page_break_before: bool | None = None

    @model_validator(mode="after")
    def _consistent(self):
        if self.line_spacing is not None and self.line_spacing_rule in (
                None, "single", "one_and_half", "double"):
            raise ValueError("line_spacing needs line_spacing_rule \"exactly\", "
                             "\"at_least\" or \"multiple\".")
        if self.line_spacing_rule in ("at_least", "exactly", "multiple") \
                and self.line_spacing is None:
            raise ValueError(f"line_spacing_rule \"{self.line_spacing_rule}\" needs "
                             "line_spacing.")
        return self


def _lines_to_points(lines: float) -> float:
    # Word stores a "multiple" spacing in points, at 12 pt per line.
    return lines * 12.0


def _set_paragraph_format_impl(start, end, alignment, line_spacing_rule, line_spacing,
                               space_before, space_after, first_line_indent, left_indent,
                               right_indent, snap_to_grid, keep_with_next,
                               page_break_before):
    doc = backend.word.document()
    start, end = check_span(doc, start, end)
    rng = span_range(doc, start, end)
    fmt = rng.ParagraphFormat
    if alignment is not None:
        fmt.Alignment = ALIGNMENTS[alignment]
    if line_spacing_rule is not None:
        fmt.LineSpacingRule = LINE_SPACING[line_spacing_rule]
        if line_spacing is not None:
            fmt.LineSpacing = (_lines_to_points(line_spacing)
                               if line_spacing_rule == "multiple" else line_spacing)
    if space_before is not None:
        fmt.SpaceBefore = space_before
    if space_after is not None:
        fmt.SpaceAfter = space_after
    if first_line_indent is not None:
        fmt.FirstLineIndent = first_line_indent
    if left_indent is not None:
        fmt.LeftIndent = left_indent
    if right_indent is not None:
        fmt.RightIndent = right_indent
    if snap_to_grid is not None:
        fmt.DisableLineHeightGrid = not snap_to_grid
    if keep_with_next is not None:
        fmt.KeepWithNext = keep_with_next
    if page_break_before is not None:
        fmt.PageBreakBefore = page_break_before

    out = [paragraph_info(i, p, detail=True) for i, p in paragraphs_in(doc, start, end)]
    warnings = []
    if line_spacing_rule is not None and out:
        got = out[0]
        wanted_lines = spacing_in_lines(
            LINE_SPACING[line_spacing_rule],
            _lines_to_points(line_spacing) if line_spacing_rule == "multiple"
            else line_spacing)
        if wanted_lines is not None:
            if got.get("line_spacing_lines") != wanted_lines:
                warnings.append(f"Line spacing reads back as {got['line_spacing_rule']}, "
                                f"{got.get('line_spacing_lines', got.get('line_spacing_pt'))}.")
        elif (got["line_spacing_rule"], got.get("line_spacing_pt")) != (
                line_spacing_rule, line_spacing):
            warnings.append(f"Line spacing reads back as {got['line_spacing_rule']}, "
                            f"{got.get('line_spacing_pt')} pt.")
    follow(doc, rng)
    result = {"success": True, "paragraphs": out[:50]}
    if warnings:
        result["warnings"] = warnings
    return result


# -- styles ---------------------------------------------------------------------

class ApplyStyleInput(Input):
    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")
    style: str = Field(..., min_length=1, description=(
        "A built-in English name (\"Heading 1\", \"Normal\", \"Title\") that works "
        "in every language of Word, or the document's own name for a style."))
    clear_direct_formatting: bool = Field(False, description=(
        "Also remove manual character formatting, so the text looks as the "
        "style says."))


def _apply_style_impl(start, end, style, clear_direct_formatting):
    doc = backend.word.document()
    start, end = check_span(doc, start, end)
    target = resolve_style(doc, style)
    rng = span_range(doc, start, end)
    rng.Style = target
    if clear_direct_formatting:
        rng.Font.Reset()
    out = [paragraph_info(i, p) for i, p in paragraphs_in(doc, start, end)]
    wanted = style_name(target)
    off = [p["index"] for p in out if p["style"] != wanted]
    follow(doc, rng)
    result = {"success": True, "style": wanted, "paragraphs": out[:50]}
    if off:
        result["warnings"] = [f"Paragraphs {off} read back with another style."]
    return result


class ListStylesInput(Input):
    in_use_only: bool = Field(True, description=(
        "Only the styles the document uses, or has changed. False lists every "
        "style, which is slow."))


# WdStyleType: 1 to 4 are documented for VBA; the interop reference adds 5
# (paragraph only) and 6 (linked) as reserved. Whether a style is linked is
# Style.Linked (docs/reference/word-object-model.md, discrepancy 2).
_STYLE_TYPES = {1: "paragraph", 2: "character", 3: "table", 4: "list", 5: "paragraph",
                6: "paragraph"}


def _list_styles_impl(in_use_only):
    doc = backend.word.document()
    styles = doc.Styles
    out = []
    for i in range(1, styles.Count + 1):
        s = styles.Item(i)
        if in_use_only and not s.InUse:
            continue
        kind = _STYLE_TYPES.get(s.Type, s.Type)
        if kind in ("table", "list"):
            continue
        entry = {"name": s.NameLocal, "type": kind, "builtin": bool(s.BuiltIn)}
        if kind == "paragraph" and s.Linked:
            entry["linked"] = True
        if kind == "paragraph":
            font = s.Font
            entry["fonts"] = {"latin": font.NameAscii, "east_asian": font.NameFarEast,
                              "size": undefined(font.Size)}
        out.append(entry)
    return {"success": True, "styles": out}


def register(mcp) -> None:
    @mcp.tool(name="word_format_text", annotations=annotations("Format text", WRITES))
    async def word_format_text(params: FormatTextInput) -> str:
        """Set fonts, size, bold, italic, underline and color on paragraphs, or on
        each occurrence of `match` inside them.

        Latin, East Asian and complex-script (right-to-left) text each have
        their own font; set the ones you mean. All three are read back, and a
        font Word did not take comes back as a warning.
        """
        p = params
        return await call(_format_text_impl, p.start, p.end, p.match, p.latin_font,
                          p.east_asian_font, p.complex_script_font, p.size,
                          p.size_complex_script, p.bold, p.italic, p.underline, p.color,
                          undo="Format text")

    @mcp.tool(name="word_set_paragraph_format",
              annotations=annotations("Paragraph format", WRITES))
    async def word_set_paragraph_format(params: ParagraphFormatInput) -> str:
        """Set alignment, line spacing, spacing before and after, indents, grid
        snapping and page-flow options on paragraphs, and read them back."""
        p = params
        return await call(_set_paragraph_format_impl, p.start, p.end, p.alignment,
                          p.line_spacing_rule, p.line_spacing, p.space_before,
                          p.space_after, p.first_line_indent, p.left_indent,
                          p.right_indent, p.snap_to_grid, p.keep_with_next,
                          p.page_break_before, undo="Paragraph format")

    @mcp.tool(name="word_apply_style", annotations=annotations("Apply a style", WRITES))
    async def word_apply_style(params: ApplyStyleInput) -> str:
        """Apply a paragraph style to paragraphs, and read the style back.

        Prefer styles to manual formatting: headings with heading styles make
        the outline, the navigation pane and the table of contents work.
        """
        p = params
        return await call(_apply_style_impl, p.start, p.end, p.style,
                          p.clear_direct_formatting, undo="Apply style")

    @mcp.tool(name="word_list_styles", annotations=annotations("List styles", READ_ONLY))
    async def word_list_styles(params: ListStylesInput) -> str:
        """The document's paragraph and character styles, with their fonts."""
        return await call(_list_styles_impl, params.in_use_only)
