"""Styles, character formatting and paragraph formatting.

Word keeps three fonts on every run of text: one for Latin text, one for
East Asian text (Japanese, Chinese, Korean) and one for complex scripts
(Arabic, Hebrew and other right-to-left text). Setting only one leaves the
others as they were, which is why Japanese text so often stays in the wrong
font. These tools take the three separately and read all three back.

Text and styles share one table of properties (tools/_props.py), so a style
can be given everything a paragraph can, and the other way round.
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from . import _props as P
from ._common import READ_ONLY, WRITES, Input, annotations, call, follow
from ._word import (
    WD_COLLAPSE_END, WD_FIND_STOP, check_span, clean_text, paragraph_info,
    paragraphs_in, resolve_style, span_range, style_name,
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


def _needs_something(model, names, what):
    if not any(getattr(model, n, None) is not None for n in names):
        raise ValueError(f"Give at least one {what} to set.")
    return model


# Always read back, whatever was set, so the caller sees the three fonts.
_FONT_ALWAYS = ("latin_font", "east_asian_font", "complex_script_font", "size", "bold",
                "italic", "color")


def font_readback(got: dict, spec: dict) -> dict:
    return {k: v for k, v in got.items() if k in spec or k in _FONT_ALWAYS}


# -- character formatting -------------------------------------------------------

class FormatTextInput(Target, P.FontSpec):
    @model_validator(mode="after")
    def _something(self):
        return _needs_something(self, P.FONT_NAMES, "formatting argument")


def _format_text_impl(start, end, match, font):
    doc = backend.word.document()
    spec = P.with_latin_extended(P.given(font, P.FONT_NAMES))
    ranges = _target_ranges(doc, start, end, match)
    warnings, readback = [], []
    for rng in ranges:
        refused = P.com_apply(rng.Font, P.FONT_PROPS, spec)
        got = P.com_read(rng.Font, P.FONT_PROPS)
        warnings += P.mismatches(spec, got, refused)
        readback.append({"text": clean_text(rng.Text)[:80], **font_readback(got, spec)})
    follow(doc, ranges[0])
    result = {"success": True, "ranges": len(ranges), "read_back": readback[:20]}
    if warnings:
        result["warnings"] = sorted(set(warnings))
    return result


# -- paragraph formatting --------------------------------------------------------

class ParagraphFormatInput(P.ParagraphSpec):
    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")

    @model_validator(mode="after")
    def _something(self):
        return _needs_something(self, P.PARAGRAPH_NAMES, "paragraph setting")


def paragraph_readback(got: dict, spec: dict) -> dict:
    keep = set(spec) | {"alignment", "space_before", "space_after", "line_spacing_rule",
                        "line_spacing_lines", "line_spacing_pt"}
    return {k: v for k, v in got.items() if k in keep}


def _set_paragraph_format_impl(start, end, paragraph):
    doc = backend.word.document()
    spec = P.given(paragraph, P.PARAGRAPH_NAMES)
    start, end = check_span(doc, start, end)
    rng = span_range(doc, start, end)
    fmt = rng.ParagraphFormat
    refused = P.com_apply(fmt, P.PARAGRAPH_PROPS, spec)
    P.com_apply_line_spacing(fmt, spec.get("line_spacing_rule"), spec.get("line_spacing"))
    out = []
    warnings = []
    for index, p in paragraphs_in(doc, start, min(end, start + 49)):
        got = {**P.com_read(p.Format, P.PARAGRAPH_PROPS),
               **P.com_read_line_spacing(p.Format)}
        if not warnings:
            warnings = P.mismatches(spec, got, refused)
        out.append({"index": index, "text": clean_text(p.Range.Text)[:80],
                    "style": style_name(p.Style), **paragraph_readback(got, spec)})
    follow(doc, rng)
    result = {"success": True, "paragraphs": out}
    if warnings:
        result["warnings"] = warnings
    return result


# -- applying a style ------------------------------------------------------------

STYLE_NAME_HELP = (
    "A built-in English name (\"Heading 1\", \"Normal\", \"Title\") that works in "
    "every language of Word, or the document's own name for a style.")


class ApplyStyleInput(Input):
    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")
    style: str = Field(..., min_length=1, description=STYLE_NAME_HELP)
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


# -- listing, reading and defining styles ------------------------------------------

class ListStylesInput(Input):
    in_use_only: bool = Field(True, description=(
        "Only the styles the document uses, or has changed. False lists every "
        "style, which is slow."))


# WdStyleType: 1 to 4 are documented for VBA; the interop reference adds 5
# (paragraph only) and 6 (linked) as reserved. Whether a style is linked is
# Style.Linked (docs/reference/word-object-model.md, discrepancy 2).
_STYLE_TYPES = {1: "paragraph", 2: "character", 3: "table", 4: "list", 5: "paragraph",
                6: "paragraph"}
_NEW_STYLE_TYPES = {"paragraph": 1, "character": 2}


def _linked(style) -> bool:
    """Style.Linked: readable on paragraph styles only (measured on Windows)."""
    try:
        return bool(style.Linked)
    except Exception:  # noqa: BLE001
        return False


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
        if kind == "paragraph" and _linked(s):
            entry["linked"] = True
        if kind == "paragraph":
            got = P.com_read(s.Font, P.FONT_PROPS)
            entry["fonts"] = {"latin": got.get("latin_font"),
                              "east_asian": got.get("east_asian_font"),
                              "size": got.get("size")}
        out.append(entry)
    return {"success": True, "styles": out}


class GetStyleInput(Input):
    style: str = Field(..., min_length=1, description=STYLE_NAME_HELP)


def _style_or_none(value):
    """BaseStyle and NextParagraphStyle answer a Style, a name, or nothing."""
    if value in (None, ""):
        return None
    return style_name(value)


def _describe_style(s) -> dict:
    kind = _STYLE_TYPES.get(s.Type, s.Type)
    out = {"name": s.NameLocal, "type": kind, "builtin": bool(s.BuiltIn),
           "in_use": bool(s.InUse)}
    for key, member in (("description", "Description"), ("quick_style", "QuickStyle"),
                        ("priority", "Priority"), ("locked", "Locked")):
        try:
            out[key] = getattr(s, member)
        except Exception:  # noqa: BLE001 - not every style type has it
            pass
    if kind == "paragraph":
        out["linked"] = _linked(s)
        for key, member in (("base_style", "BaseStyle"),
                            ("next_style", "NextParagraphStyle")):
            try:
                out[key] = _style_or_none(getattr(s, member))
            except Exception:  # noqa: BLE001
                out[key] = None
        for key, member in (("automatically_update", "AutomaticallyUpdate"),
                            ("no_space_between_same_style",
                             "NoSpaceBetweenParagraphsOfSameStyle")):
            try:
                out[key] = bool(getattr(s, member))
            except Exception:  # noqa: BLE001
                pass
    elif kind == "character":
        try:
            out["base_style"] = _style_or_none(s.BaseStyle)
        except Exception:  # noqa: BLE001
            out["base_style"] = None
    if kind in ("paragraph", "character"):
        out["font"] = P.com_read(s.Font, P.FONT_PROPS)
    if kind == "paragraph":
        fmt = s.ParagraphFormat
        out["paragraph"] = {**P.com_read(fmt, P.PARAGRAPH_PROPS),
                            **P.com_read_line_spacing(fmt)}
    return out


def _get_style_impl(style):
    doc = backend.word.document()
    return {"success": True, "style": _describe_style(resolve_style(doc, style))}


class SetStyleInput(Input):
    style: str = Field(..., min_length=1, description=(
        STYLE_NAME_HELP + " With `create`, the name of the new style."))
    create: Literal["paragraph", "character"] | None = Field(None, description=(
        "Create the style, of this type. Refused if a style of that name exists."))
    base_style: str | None = Field(None, min_length=1, description=(
        "The style this one is based on; it inherits whatever it does not set."))
    next_style: str | None = Field(None, min_length=1, description=(
        "Paragraph styles: the style of the paragraph that follows when the user "
        "presses Enter."))
    font: P.FontSpec | None = Field(None, description="Character formatting of the style.")
    paragraph: P.ParagraphSpec | None = Field(None, description=(
        "Paragraph formatting of the style (paragraph styles)."))
    quick_style: bool | None = Field(None, description="Show it in the Styles gallery.")
    priority: int | None = Field(None, ge=0, le=99, description=(
        "Its place in the Styles pane, lower first."))
    automatically_update: bool | None = Field(None, description=(
        "Paragraph styles: redefine the style when the user formats text in it."))
    no_space_between_same_style: bool | None = Field(None, description=(
        "Paragraph styles: drop the space before and after between paragraphs of "
        "this style (\"Don't add space between paragraphs of the same style\")."))

    @model_validator(mode="after")
    def _consistent(self):
        if self.create == "character" and self.paragraph is not None:
            raise ValueError("A character style has no paragraph formatting.")
        if self.create == "character" and (self.next_style or self.automatically_update
                                           or self.no_space_between_same_style):
            raise ValueError("next_style, automatically_update and "
                             "no_space_between_same_style apply to paragraph styles.")
        if self.create is None and all(getattr(self, f) is None for f in (
                "base_style", "next_style", "font", "paragraph", "quick_style", "priority",
                "automatically_update", "no_space_between_same_style")):
            raise ValueError("Give something to set, or `create` to make the style.")
        return self


def _find_style(doc, name):
    try:
        return resolve_style(doc, name)
    except ValueError:
        return None


# wdStyleNormal and wdStyleDefaultParagraphFont (reference B.2).
_PLAIN_BASE = {"paragraph": -1, "character": -66}


def _start_plain(doc, s, kind, base_style) -> None:
    """Make a style just added look like its base style and nothing else.

    Measured on Word for Windows: Styles.Add takes the paragraph style and the
    manual formatting of wherever the user's cursor is (a style added with
    the cursor in a bold Title came out based on Title, 28 pt and bold), and
    Font.Reset on a style raises 0x800A174B. Copying the base style's Font
    and ParagraphFormat gives a style that differs from its base in nothing.
    """
    from ._word import BUILTIN_STYLES, resolve_style

    # The base as the measured probe set it: a WdBuiltinStyle constant where
    # there is one (BaseStyle takes "the local name …, an integer or a
    # WdBuiltinStyle constant, or an object").
    key = " ".join((base_style or "").lower().split())
    constant = BUILTIN_STYLES.get(key) if base_style else _PLAIN_BASE[kind]
    base = doc.Styles.Item(constant) if constant is not None else resolve_style(doc, base_style)
    s.BaseStyle = constant if constant is not None else base
    s.Font = base.Font.Duplicate
    if kind == "paragraph":
        s.ParagraphFormat = base.ParagraphFormat.Duplicate


def _set_style_impl(style, create, base_style, next_style, font, paragraph, quick_style,
                    priority, automatically_update, no_space_between_same_style):
    doc = backend.word.document()
    existing = _find_style(doc, style)
    if create:
        if existing is not None:
            raise ValueError(f"A style called '{style}' already exists. Leave out "
                             "`create` to change it.")
        s = doc.Styles.Add(style, _NEW_STYLE_TYPES[create])
        _start_plain(doc, s, create, base_style)
    else:
        if existing is None:
            resolve_style(doc, style)  # raises with the names that exist
        s = existing
    kind = _STYLE_TYPES.get(s.Type, s.Type)
    if paragraph and kind != "paragraph":
        raise ValueError(f"'{style}' is a {kind} style, which has no paragraph formatting.")

    refused = {}
    if base_style is not None:
        try:
            s.BaseStyle = resolve_style(doc, base_style)
        except Exception as exc:  # noqa: BLE001
            refused["base_style"] = str(exc)
    if next_style is not None:
        try:
            s.NextParagraphStyle = resolve_style(doc, next_style)
        except Exception as exc:  # noqa: BLE001
            refused["next_style"] = str(exc)
    for key, member, value in (("quick_style", "QuickStyle", quick_style),
                               ("priority", "Priority", priority),
                               ("automatically_update", "AutomaticallyUpdate",
                                automatically_update),
                               ("no_space_between_same_style",
                                "NoSpaceBetweenParagraphsOfSameStyle",
                                no_space_between_same_style)):
        if value is not None:
            try:
                setattr(s, member, value)
            except Exception as exc:  # noqa: BLE001
                refused[key] = str(exc)
    font_spec = P.with_latin_extended(P.given(font, P.FONT_NAMES))
    para_spec = P.given(paragraph, P.PARAGRAPH_NAMES)
    font_refused = P.com_apply(s.Font, P.FONT_PROPS, font_spec) if font_spec else {}
    para_refused = {}
    if para_spec:
        fmt = s.ParagraphFormat
        para_refused = P.com_apply(fmt, P.PARAGRAPH_PROPS, para_spec)
        P.com_apply_line_spacing(fmt, para_spec.get("line_spacing_rule"),
                                 para_spec.get("line_spacing"))

    described = _describe_style(s)
    warnings = [f"{k} was refused: {v}" for k, v in refused.items()]
    for key, wanted in (("base_style", base_style), ("next_style", next_style)):
        if wanted is not None and key not in refused:
            have = described.get(key)
            if have != style_name(resolve_style(doc, wanted)):
                warnings.append(f"{key} reads back as {have!r}, not {wanted!r}.")
    for key, value in (("quick_style", quick_style), ("priority", priority),
                       ("automatically_update", automatically_update),
                       ("no_space_between_same_style", no_space_between_same_style)):
        if value is not None and key not in refused and key in described \
                and described[key] != value:
            warnings.append(f"{key} reads back as {described[key]!r}, not {value!r}.")
    warnings += P.mismatches(font_spec, described.get("font", {}), font_refused)
    warnings += P.mismatches(para_spec, described.get("paragraph", {}), para_refused)
    result = {"success": True, "created": bool(create), "style": described}
    if warnings:
        result["warnings"] = warnings
    return result


def register(mcp) -> None:
    @mcp.tool(name="word_format_text", annotations=annotations("Format text", WRITES))
    async def word_format_text(params: FormatTextInput) -> str:
        """Set character formatting on paragraphs, or on each occurrence of
        `match` inside them: fonts, size, bold, italic, underline style and
        color, color, strikethrough, super- and subscript, caps, hidden,
        character spacing, scaling, kerning, raised or lowered position,
        Japanese emphasis dots, and the character grid.

        Latin, East Asian and complex-script (right-to-left) text each have
        their own font; set the ones you mean. What was set is read back, and
        what Word did not take comes back as a warning. Prefer word_set_style
        for formatting that recurs.
        """
        p = params
        return await call(_format_text_impl, p.start, p.end, p.match,
                          p.model_dump(include=set(P.FONT_NAMES)), undo="Format text")

    @mcp.tool(name="word_set_paragraph_format",
              annotations=annotations("Paragraph format", WRITES))
    async def word_set_paragraph_format(params: ParagraphFormatInput) -> str:
        """Set paragraph formatting and read it back: alignment, line spacing,
        spacing before and after in points or in lines (行), indents in points
        or in characters (字), grid snapping, keep-with-next and page flow,
        outline level, and the East Asian options (line-breaking rules, hanging
        punctuation, space between Japanese and Latin text)."""
        p = params
        return await call(_set_paragraph_format_impl, p.start, p.end,
                          p.model_dump(include=set(P.PARAGRAPH_NAMES)),
                          undo="Paragraph format")

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

    @mcp.tool(name="word_get_style", annotations=annotations("Read a style", READ_ONLY))
    async def word_get_style(params: GetStyleInput) -> str:
        """Everything a style defines: its type, base and next style, gallery
        settings, every character property and every paragraph property, in
        the same names word_set_style takes."""
        return await call(_get_style_impl, params.style)

    @mcp.tool(name="word_set_style", annotations=annotations("Define a style", WRITES))
    async def word_set_style(params: SetStyleInput) -> str:
        """Create a style or change one: base and next style, gallery settings,
        and any character or paragraph property, then read the whole style back.

        Changing a style changes every paragraph in it at once, which is how a
        document stays consistent: define the look in styles, then apply them.
        """
        p = params
        return await call(
            _set_style_impl, p.style, p.create, p.base_style, p.next_style,
            p.font.model_dump() if p.font else None,
            p.paragraph.model_dump() if p.paragraph else None,
            p.quick_style, p.priority, p.automatically_update,
            p.no_space_between_same_style, undo="Define style")
