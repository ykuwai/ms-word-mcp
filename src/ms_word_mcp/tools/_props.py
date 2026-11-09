"""Character and paragraph properties, written once for text and for styles.

A Font is a Font whether it comes from a range or from a style, and so is a
ParagraphFormat, so one table drives word_format_text,
word_set_paragraph_format and word_set_style on both platforms. Each row
names the property as the tools take it, the COM member it sets (checked in
docs/reference/word-object-model.md, sections C.3 and C.4), and the Word for
Mac property (from Word.sdef).

Kinds:
  pt      points (float)            lines   gridlines (LineUnit*)
  chars   characters (CharacterUnit*)  flag  true or false
  int     a whole number            font    a font name
  color   #RRGGBB                   enum    one of the named values
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

from pydantic import Field, model_validator

from ._common import Input, Trimmed

WD_UNDEFINED = 9999999


@dataclass(frozen=True)
class Prop:
    name: str
    com: str | None
    mac: str | None
    kind: str
    values: dict | None = None      # enum: name -> COM value
    mac_values: dict | None = None  # enum: name -> Mac keyword name
    invert: bool = False            # the member means the opposite of the name


UNDERLINES = {"none": 0, "single": 1, "words": 2, "double": 3, "dotted": 4, "thick": 6,
              "dash": 7, "dot_dash": 9, "dot_dot_dash": 10, "wavy": 11,
              "dotted_heavy": 20, "dash_heavy": 23, "dot_dash_heavy": 25,
              "dot_dot_dash_heavy": 26, "wavy_heavy": 27, "dash_long": 39,
              "wavy_double": 43, "dash_long_heavy": 55}
EMPHASIS = {"none": 0, "dot_above": 1, "comma_above": 2, "circle_above": 3,
            "dot_below": 4}
EMPHASIS_MAC = {"none": "emphasis_mark_none", "dot_above": "emphasis_mark_over_solid_circle",
                "comma_above": "emphasis_mark_over_comma",
                "circle_above": "emphasis_mark_over_white_circle",
                "dot_below": "emphasis_mark_under_solid_circle"}
ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3, "distribute": 4}
ALIGNMENTS_MAC = {k: "align_paragraph_" + k for k in ALIGNMENTS}
BASELINES = {"top": 0, "center": 1, "baseline": 2, "east_asian_50": 3, "auto": 4}
BASELINES_MAC = {"top": "baseline_align_top", "center": "baseline_align_center",
                 "baseline": "baseline_align_baseline",
                 "east_asian_50": "baseline_align_east_asian50", "auto": "baseline_align_auto"}
DIRECTIONS = {"left_to_right": 1, "right_to_left": 0}  # WdReadingOrder
TAB_ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "decimal": 3, "bar": 4}
TAB_LEADERS = {"none": 0, "dots": 1, "dashes": 2, "lines": 3, "heavy": 4, "middle_dot": 5}
OUTLINE = {**{str(i): i for i in range(1, 10)}, "body_text": 10}
OUTLINE_MAC = {**{str(i): f"outline_level{i}" for i in range(1, 10)},
               "body_text": "outline_level_body_text"}

# Paragraph borders. WdBorderType, WdLineStyle and WdLineWidth (eighths of a
# point); the Mac names are from Word.sdef.
BORDER_SIDES = {"top": -1, "left": -2, "bottom": -3, "right": -4}
BORDER_STYLES = {"none": 0, "single": 1, "dotted": 2, "dash_small_gap": 3, "dash_large_gap": 4,
                 "dash_dot": 5, "dash_dot_dot": 6, "double": 7, "triple": 8}
BORDER_STYLES_MAC = {"none": "line_style_none", "single": "line_style_single",
                     "dotted": "line_style_dot", "dash_small_gap": "line_style_dash_small_gap",
                     "dash_large_gap": "line_style_dash_large_gap",
                     "dash_dot": "line_style_dash_dot",
                     "dash_dot_dot": "line_style_dash_dot_dot",
                     "double": "line_style_double", "triple": "line_style_triple"}
BORDER_WIDTHS = {0.25: 2, 0.5: 4, 0.75: 6, 1.0: 8, 1.5: 12, 2.25: 18, 3.0: 24, 4.5: 36, 6.0: 48}
BORDER_WIDTHS_MAC = {0.25: "line_width25_point", 0.5: "line_width50_point",
                     0.75: "line_width75_point", 1.0: "line_width100_point",
                     1.5: "line_width150_point", 2.25: "line_width225_point",
                     3.0: "line_width300_point", 4.5: "line_width450_point",
                     6.0: "line_width600_point"}
WD_COLOR_AUTOMATIC = -16777216

FONT_PROPS = (
    Prop("latin_font", "NameAscii", "ascii_name", "font"),
    Prop("latin_font_extended", "NameOther", "other_name", "font"),
    Prop("east_asian_font", "NameFarEast", "east_asian_name", "font"),
    Prop("complex_script_font", "NameBi", "complex_script_name", "font"),
    Prop("size", "Size", "font_size", "pt"),
    Prop("size_complex_script", "SizeBi", None, "pt"),
    Prop("bold", "Bold", "bold", "flag"),
    Prop("bold_complex_script", "BoldBi", "bold_bi", "flag"),
    Prop("italic", "Italic", "italic", "flag"),
    Prop("italic_complex_script", "ItalicBi", "italic_bi", "flag"),
    Prop("underline", "Underline", "underline", "enum", UNDERLINES,
         {k: "underline_" + k for k in UNDERLINES}),
    Prop("underline_color", "UnderlineColor", None, "color"),
    Prop("color", "Color", "color", "color"),
    Prop("strikethrough", "StrikeThrough", "strike_through", "flag"),
    Prop("double_strikethrough", "DoubleStrikeThrough", "double_strike_through", "flag"),
    Prop("superscript", "Superscript", "superscript", "flag"),
    Prop("subscript", "Subscript", "subscript", "flag"),
    Prop("small_caps", "SmallCaps", "small_caps", "flag"),
    Prop("all_caps", "AllCaps", "all_caps", "flag"),
    Prop("hidden", "Hidden", "hidden", "flag"),
    Prop("character_spacing", "Spacing", "spacing", "pt"),
    Prop("scaling", "Scaling", "scaling", "int"),
    Prop("kerning_from", "Kerning", "kerning", "pt"),
    Prop("position", "Position", "font_position", "pt"),
    Prop("emphasis_mark", "EmphasisMark", "emphasis_mark", "enum", EMPHASIS, EMPHASIS_MAC),
    Prop("ignore_character_grid", "DisableCharacterSpaceGrid",
         "disable_character_space_grid", "flag"),
)

PARAGRAPH_PROPS = (
    Prop("alignment", "Alignment", "alignment", "enum", ALIGNMENTS, ALIGNMENTS_MAC),
    # ParagraphFormat.ReadingOrder (WdReadingOrder); Word for Mac's dictionary
    # has no reading order.
    Prop("direction", "ReadingOrder", None, "enum", DIRECTIONS),
    Prop("space_before", "SpaceBefore", "space_before", "pt"),
    Prop("space_after", "SpaceAfter", "space_after", "pt"),
    Prop("space_before_lines", "LineUnitBefore", "line_unit_before", "lines"),
    Prop("space_after_lines", "LineUnitAfter", "line_unit_after", "lines"),
    Prop("space_before_auto", "SpaceBeforeAuto", "space_before_auto", "flag"),
    Prop("space_after_auto", "SpaceAfterAuto", "space_after_auto", "flag"),
    Prop("first_line_indent", "FirstLineIndent", "first_line_indent", "pt"),
    Prop("first_line_indent_chars", "CharacterUnitFirstLineIndent",
         "character_unit_first_line_indent", "chars"),
    Prop("left_indent", "LeftIndent", "paragraph_format_left_indent", "pt"),
    Prop("left_indent_chars", "CharacterUnitLeftIndent", "character_unit_left_indent",
         "chars"),
    Prop("right_indent", "RightIndent", "paragraph_format_right_indent", "pt"),
    Prop("right_indent_chars", "CharacterUnitRightIndent", "character_unit_right_indent",
         "chars"),
    # DisableLineHeightGrid True writes <w:snapToGrid w:val="0"/>: measured on
    # Windows (docs/PROBES.md), whatever the documentation's sentence says.
    Prop("snap_to_grid", "DisableLineHeightGrid", "disable_line_height_grid", "flag",
         invert=True),
    Prop("keep_with_next", "KeepWithNext", "keep_with_next", "flag"),
    Prop("keep_together", "KeepTogether", "keep_together", "flag"),
    Prop("page_break_before", "PageBreakBefore", "page_break_before", "flag"),
    Prop("widow_control", "WidowControl", "widow_control", "flag"),
    Prop("outline_level", "OutlineLevel", "outline_level", "enum", OUTLINE, OUTLINE_MAC),
    Prop("east_asian_line_breaking", "FarEastLineBreakControl",
         "east_asian_line_break_control", "flag"),
    Prop("hanging_punctuation", "HangingPunctuation", "hanging_punctuation", "flag"),
    Prop("half_width_punctuation_at_line_start", "HalfWidthPunctuationOnTopOfLine",
         "half_width_punctuation_on_top_of_line", "flag"),
    Prop("space_between_east_asian_and_latin", "AddSpaceBetweenFarEastAndAlpha",
         "add_space_between_east_asian_and_alpha", "flag"),
    Prop("space_between_east_asian_and_digits", "AddSpaceBetweenFarEastAndDigit",
         "add_space_between_east_asian_and_digit", "flag"),
    Prop("auto_adjust_right_indent", "AutoAdjustRightIndent", "auto_adjust_right_indent",
         "flag"),
    Prop("break_latin_words", "WordWrap", "word_wrap", "flag"),
    Prop("baseline_alignment", "BaseLineAlignment", "base_line_alignment", "enum",
         BASELINES, BASELINES_MAC),
)

# The same spacing or indent in points and in lines or characters. Word keeps
# both; the one in lines or characters wins while it is not zero (measured on
# the Mac: 0.5 lines before read back as 9 pt before).
PAIRED_UNITS = {"space_before": "space_before_lines", "space_after": "space_after_lines",
                "first_line_indent": "first_line_indent_chars",
                "left_indent": "left_indent_chars", "right_indent": "right_indent_chars"}

FONT_NAMES = tuple(p.name for p in FONT_PROPS)
PARAGRAPH_NAMES = tuple(p.name for p in PARAGRAPH_PROPS) + ("line_spacing_rule",
                                                            "line_spacing", "tab_stops",
                                                            "borders", "shading")
LINE_SPACING = {"single": 0, "one_and_half": 1, "double": 2, "at_least": 3, "exactly": 4,
                "multiple": 5}
LINE_SPACING_MAC = {"single": "line_space_single", "one_and_half": "line_space1_pt5",
                    "double": "line_space_double", "at_least": "line_space_at_least",
                    "exactly": "line_space_exactly", "multiple": "line_space_multiple"}
_FIXED_LINES = {"single": 1.0, "one_and_half": 1.5, "double": 2.0}

Pt = float
_PT = dict(ge=-1584, le=1584)


class FontSpec(Input):
    """Character formatting. Every field is optional; leave out what stays."""

    latin_font: Trimmed | None = Field(None, description=(
        "Font for Latin text (alphabet, digits). Also sets latin_font_extended "
        "unless that is given."))
    latin_font_extended: Trimmed | None = Field(None, description=(
        "Font for accented Latin characters (codes 128-255). Defaults to latin_font."))
    east_asian_font: Trimmed | None = Field(None, description=(
        "Font for Japanese, Chinese and Korean text. Word refuses a font without "
        "East Asian glyphs here; the result says so."))
    complex_script_font: Trimmed | None = Field(None, description=(
        "Font for Arabic, Hebrew and other right-to-left text."))
    size: float | None = Field(None, gt=0, le=1638, description="Points, Latin and East Asian.")
    size_complex_script: float | None = Field(None, gt=0, le=1638, description=(
        "Points, complex-script text (Windows only)."))
    bold: bool | None = None
    bold_complex_script: bool | None = None
    italic: bool | None = None
    italic_complex_script: bool | None = None
    underline: Literal[tuple(UNDERLINES)] | None = None  # type: ignore[valid-type]
    underline_color: Trimmed | None = Field(None, pattern=r"^#?[0-9A-Fa-f]{6}$", description=(
        "#RRGGBB (Windows only)."))
    color: Trimmed | None = Field(None, pattern=r"^#?[0-9A-Fa-f]{6}$", description="#RRGGBB.")
    strikethrough: bool | None = None
    double_strikethrough: bool | None = None
    superscript: bool | None = None
    subscript: bool | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None
    hidden: bool | None = None
    character_spacing: float | None = Field(None, ge=-1584, le=1584, description=(
        "Points added between characters; negative condenses."))
    scaling: int | None = Field(None, ge=1, le=600, description=(
        "Horizontal scaling, percent (100 = normal)."))
    kerning_from: float | None = Field(None, ge=0, le=1638, description=(
        "Kern fonts at this size and above, points; 0 turns kerning off."))
    position: float | None = Field(None, ge=-1584, le=1584, description=(
        "Raise (positive) or lower (negative) the text, points."))
    emphasis_mark: Literal[tuple(EMPHASIS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "East Asian emphasis marks, one on each character: dot_above, "
        "dot_below, comma_above or circle_above."))
    ignore_character_grid: bool | None = Field(None, description=(
        "Ignore the document's characters-per-line grid."))

    @model_validator(mode="after")
    def _one_of_super_sub(self):
        if self.superscript and self.subscript:
            raise ValueError("superscript and subscript cannot both be true.")
        return self


class TabStop(Input):
    position: float = Field(..., ge=0, le=1584, description="Points from the left margin.")
    alignment: Literal[tuple(TAB_ALIGNMENTS)] = "left"  # type: ignore[valid-type]
    leader: Literal[tuple(TAB_LEADERS)] = Field("none", description=(  # type: ignore[valid-type]
        "What fills the space before the tab stop: dots for a table of contents "
        "look."))


class Border(Input):
    style: Literal[tuple(BORDER_STYLES)] = Field("single", description=(  # type: ignore[valid-type]
        "The line; none takes this side's border off."))
    width: float | None = Field(None, description=(
        "Points: 0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5 or 6."))
    color: Trimmed | None = Field(None, pattern=r"^#?[0-9A-Fa-f]{6}$", description="#RRGGBB.")

    @model_validator(mode="after")
    def _a_width_word_has(self):
        if self.width is not None and self.width not in BORDER_WIDTHS:
            raise ValueError(f"A border's width is one of {sorted(BORDER_WIDTHS)} points, "
                             f"not {self.width}.")
        return self


class ParagraphBorders(Input):
    """Each side given is set; a side left out stays as it is."""

    top: Border | None = None
    bottom: Border | None = None
    left: Border | None = None
    right: Border | None = None


class ParagraphSpec(Input):
    """Paragraph formatting. Every field is optional; leave out what stays."""

    alignment: Literal[tuple(ALIGNMENTS)] | None = None  # type: ignore[valid-type]
    direction: Literal[tuple(DIRECTIONS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "right_to_left for Arabic, Hebrew and other right-to-left paragraphs: "
        "punctuation and numbers then sit where those languages put them. "
        "Windows only; Word for Mac's scripting has no reading order."))
    tab_stops: list[TabStop] | None = Field(None, max_length=20, description=(
        "The paragraphs' own tab stops, replacing the ones they have; [] clears "
        "them. Windows only for now."))
    borders: ParagraphBorders | None = Field(None, description=(
        "Lines around the paragraphs, by side: {\"bottom\": {\"style\": \"single\", "
        "\"width\": 1.5, \"color\": \"#1F4E79\"}} for a rule under a heading."))
    shading: Trimmed | None = Field(None, pattern=r"^(none|#?[0-9A-Fa-f]{6})$", description=(
        "The paragraphs' background color, #RRGGBB, or none to take it off."))
    line_spacing_rule: Literal[tuple(LINE_SPACING)] | None = Field(None, description=(
        "single / one_and_half / double, or at_least / exactly (line_spacing in "
        "points) or multiple (line_spacing in lines, 1.2 = 1.2 lines)."))
    line_spacing: float | None = Field(None, gt=0, le=1584)
    space_before: float | None = Field(None, ge=0, le=1584, description="Points.")
    space_after: float | None = Field(None, ge=0, le=1584, description="Points.")
    space_before_lines: float | None = Field(None, ge=0, le=100, description=(
        "Lines of the document grid, the way East Asian (Chinese, Japanese, "
        "Korean) templates measure spacing; it takes precedence over "
        "space_before while set."))
    space_after_lines: float | None = Field(None, ge=0, le=100, description=(
        "Lines of the document grid; takes precedence over space_after."))
    space_before_auto: bool | None = None
    space_after_auto: bool | None = None
    first_line_indent: float | None = Field(None, **_PT, description=(
        "Points; negative for a hanging indent."))
    first_line_indent_chars: float | None = Field(None, ge=-100, le=100, description=(
        "Character widths, the unit East Asian layouts give indents in; "
        "negative for a hanging indent. 1 = a one-character first-line indent, "
        "usual in Chinese and Japanese body text."))
    left_indent: float | None = Field(None, **_PT, description="Points.")
    left_indent_chars: float | None = Field(None, ge=-100, le=100, description="Characters.")
    right_indent: float | None = Field(None, **_PT, description="Points.")
    right_indent_chars: float | None = Field(None, ge=-100, le=100, description="Characters.")
    snap_to_grid: bool | None = Field(None, description=(
        "Fit lines to the document grid. Japanese templates turn it on, and it "
        "stretches line spacing: turn it off for exact spacing."))
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    page_break_before: bool | None = None
    widow_control: bool | None = None
    outline_level: Literal[tuple(OUTLINE)] | None = Field(None, description=(  # type: ignore[valid-type]
        "\"1\"-\"9\" or \"body_text\". Paragraphs in Heading 1-9 keep their heading's level."))
    east_asian_line_breaking: bool | None = Field(None, description=(
        "Apply East Asian line-breaking rules: no line starts with a closing "
        "bracket or a small kana, none ends with an opening bracket."))
    hanging_punctuation: bool | None = Field(None, description=(
        "Let punctuation at the end of a line hang past the margin."))
    half_width_punctuation_at_line_start: bool | None = None
    space_between_east_asian_and_latin: bool | None = Field(None, description=(
        "Add space between Japanese and Latin text automatically."))
    space_between_east_asian_and_digits: bool | None = None
    auto_adjust_right_indent: bool | None = Field(None, description=(
        "Adjust the right indent to the characters-per-line grid."))
    break_latin_words: bool | None = Field(None, description=(
        "Allow Latin words to break in the middle at the end of a line."))
    baseline_alignment: Literal[tuple(BASELINES)] | None = None  # type: ignore[valid-type]

    @model_validator(mode="after")
    def _spacing(self):
        for points, units in PAIRED_UNITS.items():
            if getattr(self, points) is not None and getattr(self, units) is not None:
                raise ValueError(f"Give {points} or {units}, not both: Word keeps one "
                                 "measure for each, and the one in lines or characters "
                                 "wins while it is set.")
        if self.line_spacing is not None and self.line_spacing_rule in (
                None, "single", "one_and_half", "double"):
            raise ValueError("line_spacing needs line_spacing_rule \"exactly\", "
                             "\"at_least\" or \"multiple\".")
        if self.line_spacing_rule in ("at_least", "exactly", "multiple") \
                and self.line_spacing is None:
            raise ValueError(f"line_spacing_rule \"{self.line_spacing_rule}\" needs "
                             "line_spacing.")
        return self


def given(spec: dict | None, names) -> dict:
    """The fields of `spec` that were set."""
    return {k: v for k, v in (spec or {}).items() if k in names and v is not None}


# -- color -----------------------------------------------------------------------

def bgr(hex_color: str) -> int:
    v = hex_color.lstrip("#")
    return int(v[0:2], 16) | (int(v[2:4], 16) << 8) | (int(v[4:6], 16) << 16)


def hex_of(value) -> str | None:
    """#RRGGBB for a WdColor, or None for automatic (-16777216) and mixed."""
    if not isinstance(value, int) or value < 0 or value > 0xFFFFFF:
        return None
    return f"#{value & 0xFF:02X}{(value >> 8) & 0xFF:02X}{(value >> 16) & 0xFF:02X}"


# -- COM -------------------------------------------------------------------------

def _com_value(prop: Prop, value):
    if prop.kind == "flag":
        return (not value) if prop.invert else bool(value)
    if prop.kind == "enum":
        return prop.values[value]
    if prop.kind == "color":
        return bgr(value)
    return value


def _com_read(prop: Prop, raw):
    if raw in (WD_UNDEFINED, float(WD_UNDEFINED)):
        return None
    if prop.kind == "flag":
        value = bool(raw)
        return (not value) if prop.invert else value
    if prop.kind == "enum":
        for name, code in prop.values.items():
            if code == raw:
                return name
        return raw
    if prop.kind == "color":
        return hex_of(raw)
    if prop.kind == "font":
        return raw or None
    return raw


def paired_steps(spec: dict) -> tuple[dict, dict]:
    """What to set before and after `spec` so a measure in points takes.

    Word keeps an indent or a space both in points and in characters (or
    lines), and the one in characters wins while it is not zero; a style
    inherits both from its base. So a value in points is set after the
    characters are made 0, and 0 characters also sets 0 points (seen in an
    agent's run on the Mac: Normal's 1-character first-line indent stayed on
    the headings based on it until both were 0).
    """
    before, after = {}, {}
    for points, units in PAIRED_UNITS.items():
        if points in spec and units not in spec:
            before[units] = 0.0
        if units in spec and spec[units] == 0 and points not in spec:
            after[points] = 0.0
    return before, after


def com_apply(obj, props, spec: dict) -> dict:
    """Set each given property on a COM Font or ParagraphFormat.

    Returns {name: reason} for the ones Word refused; the rest still apply.
    """
    from ._common import error_text

    refused = {}
    before, after = paired_steps(spec) if props is PARAGRAPH_PROPS else ({}, {})
    for step in (before, spec, after):
        for prop in props:
            if prop.name not in step:
                continue
            try:
                setattr(obj, prop.com, _com_value(prop, step[prop.name]))
            except Exception as exc:  # noqa: BLE001 - reported to the caller
                if step is spec:
                    refused[prop.name] = error_text(exc)
    return refused


def com_read(obj, props) -> dict:
    out = {}
    for prop in props:
        try:
            out[prop.name] = _com_read(prop, getattr(obj, prop.com))
        except Exception:  # noqa: BLE001 - a property this object does not have
            continue
    return out


def com_apply_line_spacing(fmt, rule, value) -> None:
    if rule is None:
        return
    fmt.LineSpacingRule = LINE_SPACING[rule]
    if value is not None:
        fmt.LineSpacing = value * 12.0 if rule == "multiple" else value


def com_apply_tab_stops(fmt, tab_stops) -> None:
    if tab_stops is None:
        return
    stops = fmt.TabStops
    stops.ClearAll()
    for stop in tab_stops:
        stops.Add(stop["position"], TAB_ALIGNMENTS[stop["alignment"]],
                  TAB_LEADERS[stop["leader"]])


def _name_of(table: dict, value):
    return next((name for name, v in table.items() if v == value), value)


def _border_steps(spec: dict, widths: dict) -> list[tuple[str, object]]:
    """What to set on one side, in the order Word needs: a line before its width."""
    if spec["style"] == "none":
        return [("style", "none")]
    steps = [("style", spec["style"])]
    if spec.get("width") is not None:
        steps.append(("width", widths[spec["width"]]))
    if spec.get("color") is not None:
        steps.append(("color", spec["color"]))
    return steps


def com_apply_borders(fmt, borders: dict | None) -> dict:
    """Set the sides given; answers what Word refused, by name."""
    refused = {}
    members = {"style": "LineStyle", "width": "LineWidth", "color": "Color"}
    for side, spec in (borders or {}).items():
        if spec is None:
            continue
        line = fmt.Borders.Item(BORDER_SIDES[side])
        for key, value in _border_steps(spec, BORDER_WIDTHS):
            value = (BORDER_STYLES[value] if key == "style"
                     else bgr(value) if key == "color" else value)
            try:
                setattr(line, members[key], value)
            except Exception as exc:  # noqa: BLE001 - reported by name
                refused[f"borders.{side}.{key}"] = str(exc).splitlines()[0]
    return refused


def com_read_borders(fmt) -> dict:
    """The sides that have a line."""
    out = {}
    for side, code in BORDER_SIDES.items():
        line = fmt.Borders.Item(code)
        style = line.LineStyle
        if style != BORDER_STYLES["none"]:
            out[side] = {"style": _name_of(BORDER_STYLES, style),
                         "width": _name_of(BORDER_WIDTHS, line.LineWidth),
                         "color": hex_of(line.Color)}
    return out


def com_apply_shading(fmt, shading: str | None) -> None:
    if shading is not None:
        fmt.Shading.BackgroundPatternColor = (WD_COLOR_AUTOMATIC if shading == "none"
                                              else bgr(shading))


def com_read_shading(fmt) -> str | None:
    return hex_of(fmt.Shading.BackgroundPatternColor)


def border_mismatches(spec: dict, got: dict) -> list[str]:
    """A sentence for each side or background asked for that reads back otherwise."""
    out = []
    for side, want in (spec.get("borders") or {}).items():
        if want is None:
            continue
        have = (got.get("borders") or {}).get(side)
        if want["style"] == "none":
            if have:
                out.append(f"The {side} border is still there: {have}.")
            continue
        for key in ("style", "width", "color"):
            wanted = want.get(key)
            value = (have or {}).get(key)
            if wanted is None:
                continue
            same = (wanted.lstrip("#").upper() == (value or "").lstrip("#").upper()
                    if key == "color" else value == wanted)
            if not same:
                out.append(f"The {side} border's {key} reads back as {value!r}, not {wanted!r}.")
    wanted = spec.get("shading")
    if wanted is not None:
        have = got.get("shading")
        if have != (None if wanted == "none" else "#" + wanted.lstrip("#").upper()):
            out.append(f"shading reads back as {have!r}, not {wanted!r}.")
    return out


def com_read_tab_stops(fmt) -> list:
    """The custom tab stops. TabStops also holds Word's default ones (every
    36 pt after the last custom one, measured on Windows), which
    TabStop.CustomTab tells apart."""
    stops = fmt.TabStops
    out = []
    for i in range(1, stops.Count + 1):
        t = stops.Item(i)
        if not t.CustomTab:
            continue
        out.append({"position": round(t.Position, 2),
                    "alignment": next((k for k, v in TAB_ALIGNMENTS.items()
                                       if v == t.Alignment), t.Alignment),
                    "leader": next((k for k, v in TAB_LEADERS.items() if v == t.Leader),
                                   t.Leader)})
    return out


def com_read_line_spacing(fmt) -> dict:
    code, points = fmt.LineSpacingRule, fmt.LineSpacing
    rule = next((k for k, v in LINE_SPACING.items() if v == code), None)
    return describe_line_spacing(rule, points)


def describe_line_spacing(rule, points) -> dict:
    """Line spacing as the tools report it: lines for the line rules, points else.

    Word stores a multiple in points, 12 to a line, and turns a multiple of
    exactly 1, 1.5 or 2 lines into the single, one-and-a-half or double rule.
    """
    if rule is None:
        return {"line_spacing_rule": None}
    if rule in _FIXED_LINES:
        return {"line_spacing_rule": rule, "line_spacing_lines": _FIXED_LINES[rule]}
    if rule == "multiple":
        return {"line_spacing_rule": rule, "line_spacing_lines": round(points / 12.0, 3)}
    return {"line_spacing_rule": rule, "line_spacing_pt": points}


def expected_lines(rule, value) -> float | None:
    if rule in _FIXED_LINES:
        return _FIXED_LINES[rule]
    if rule == "multiple":
        return round(value, 3)
    return None


# -- comparing what was asked with what reads back ---------------------------------

def mismatches(spec: dict, got: dict, refused: dict) -> list[str]:
    """A sentence for each asked property that reads back otherwise."""
    out = []
    for name, wanted in spec.items():
        if name in ("line_spacing_rule", "line_spacing", "tab_stops", "borders", "shading") \
                or name not in got:
            continue
        have = got[name]
        if isinstance(wanted, str) and isinstance(have, str) and name in ("color",
                                                                          "underline_color"):
            same = wanted.lstrip("#").upper() == have.lstrip("#").upper()
        elif isinstance(wanted, float) or isinstance(have, float):
            same = have is not None and abs(float(have) - float(wanted)) < 0.05
        else:
            same = have == wanted
        if not same:
            why = f" Word refused it ({refused[name]})." if name in refused else ""
            if name == "east_asian_font" and name in refused:
                why = (" Word refused it: an East Asian font needs East Asian glyphs "
                       f"({refused[name]}).")
            out.append(f"{name} reads back as {have!r}, not {wanted!r}.{why}")
    for name, reason in refused.items():
        if name not in got:
            out.append(f"{name} was refused: {reason}")
    if "line_spacing_rule" in spec:
        rule, value = spec["line_spacing_rule"], spec.get("line_spacing")
        wanted_lines = expected_lines(rule, value)
        have_lines = got.get("line_spacing_lines")
        if wanted_lines is not None:
            if have_lines is None or abs(have_lines - wanted_lines) > 0.01:
                out.append(f"Line spacing reads back as {got.get('line_spacing_rule')}, "
                           f"{have_lines or got.get('line_spacing_pt')}.")
        elif (got.get("line_spacing_rule"), got.get("line_spacing_pt")) != (rule, value):
            out.append(f"Line spacing reads back as {got.get('line_spacing_rule')}, "
                       f"{got.get('line_spacing_pt')} pt.")
    return out


def with_latin_extended(spec: dict) -> dict:
    """Fill the paired properties a caller expects one argument to cover.

    latin_font also sets latin_font_extended, and bold, italic and size also
    set their complex-script twins, unless those are given: an agent set
    bold on a heading and its Arabic stayed regular.
    """
    out = dict(spec)
    for one, twin in (("latin_font", "latin_font_extended"), ("bold", "bold_complex_script"),
                      ("italic", "italic_complex_script"), ("size", "size_complex_script")):
        if one in out and twin not in out:
            out[twin] = out[one]
    return out


# -- Mac (appscript) ---------------------------------------------------------------

def _k():
    from appscript import k

    return k


def _kname(value) -> str:
    return getattr(value, "name", None) or str(value).removeprefix("k.")


def mac_font_name(value):
    """A font name as Windows gives it: Word for Mac adds " ()" to theme fonts."""
    if not isinstance(value, str) or not value:
        return None
    return value[:-3] if value.endswith(" ()") else value


def _mac_value(prop: Prop, value):
    if prop.kind == "flag":
        return (not value) if prop.invert else bool(value)
    if prop.kind == "enum":
        return getattr(_k(), prop.mac_values[value])
    if prop.kind == "color":
        # Word for Mac takes 16-bit channels: [192, 0, 0] is black, and
        # [49344, 0, 0] is #C00000 (measured, docs/PROBES.md).
        v = value.lstrip("#")
        return [int(v[i:i + 2], 16) * 257 for i in (0, 2, 4)]
    return value


def _mac_read(prop: Prop, raw):
    k = _k()
    if raw == k.missing_value:
        return None
    if prop.kind == "flag":
        if not isinstance(raw, bool):
            return None
        return (not raw) if prop.invert else raw
    if prop.kind == "enum":
        name = _kname(raw)
        for key, kname in prop.mac_values.items():
            if kname == name:
                return key
        return name
    if prop.kind == "color":
        if isinstance(raw, (list, tuple)) and len(raw) == 3:
            r, g, b = (round(int(c) / 257) for c in raw)
            return f"#{r:02X}{g:02X}{b:02X}"
        return None
    if prop.kind == "font":
        return mac_font_name(raw)
    return raw


_TWINS = {"latin_font_extended": "latin_font", "bold_complex_script": "bold",
          "italic_complex_script": "italic", "size_complex_script": "size"}


def _filled_twin(spec: dict, name: str) -> bool:
    one = _TWINS.get(name)
    return one is not None and one in spec and spec[one] == spec[name]


def mac_apply(obj, props, spec: dict) -> dict:
    """Set each given property on a Word for Mac font or paragraph format."""
    refused = {}
    before, after = paired_steps(spec) if props is PARAGRAPH_PROPS else ({}, {})
    for step in (before, spec, after):
        for prop in props:
            if prop.name not in step:
                continue
            if prop.mac is None:
                # A twin filled in from its pair (size -> size_complex_script)
                # is left out quietly; one the caller asked for is refused.
                if step is spec and not _filled_twin(spec, prop.name):
                    refused[prop.name] = ("Word for Mac has no such property; it works "
                                          "on Windows.")
                continue
            try:
                getattr(obj, prop.mac).set(_mac_value(prop, step[prop.name]))
            except Exception as exc:  # noqa: BLE001 - reported to the caller
                if step is spec:
                    refused[prop.name] = str(exc).splitlines()[0]
    return refused


def _mac_rgb(hex_color: str) -> list[int]:
    v = hex_color.lstrip("#")
    return [int(v[i:i + 2], 16) * 257 for i in (0, 2, 4)]


def _mac_hex(raw) -> str | None:
    if isinstance(raw, (list, tuple)) and len(raw) == 3:
        r, g, b = (round(int(c) / 257) for c in raw)
        return f"#{r:02X}{g:02X}{b:02X}"
    return None


def mac_apply_borders(owner, borders: dict | None) -> dict:
    """Set the sides given on a paragraph or a style.

    `get border` answers a border that works from the paragraph or the style.
    From a paragraph format it answers one that every read refuses (-1728,
    measured on Word 16 for Mac).
    """
    k = _k()
    refused = {}
    for side, spec in (borders or {}).items():
        if spec is None:
            continue
        line = owner.get_border(which_border=getattr(k, f"border_{side}"))
        for key, value in _border_steps(spec, BORDER_WIDTHS_MAC):
            try:
                if key == "style":
                    line.line_style.set(getattr(k, BORDER_STYLES_MAC[value]))
                elif key == "width":
                    line.line_width.set(getattr(k, value))
                else:
                    line.color.set(_mac_rgb(value))
            except Exception as exc:  # noqa: BLE001 - reported by name
                refused[f"borders.{side}.{key}"] = str(exc).splitlines()[0]
    return refused


def mac_read_borders(owner) -> dict:
    k = _k()
    out = {}
    for side in BORDER_SIDES:
        line = owner.get_border(which_border=getattr(k, f"border_{side}"))
        style = _name_of(BORDER_STYLES_MAC, _kname(line.line_style.get()))
        if style != "none":
            out[side] = {"style": style,
                         "width": _name_of(BORDER_WIDTHS_MAC, _kname(line.line_width.get())),
                         "color": _mac_hex(line.color.get())}
    return out


def mac_apply_shading(owner, shading: str | None) -> None:
    if shading == "none":
        owner.shading.background_pattern_color_index.set(_k().no_highlight)
    elif shading is not None:
        owner.shading.background_pattern_color.set(_mac_rgb(shading))


def mac_read_shading(owner) -> str | None:
    """None when there is no background: the color then reads as black."""
    shade = owner.shading
    if shade.background_pattern_color_index.get() == _k().no_highlight:
        return None
    return _mac_hex(shade.background_pattern_color.get())


def mac_read(obj, props) -> dict:
    out = {}
    for prop in props:
        if prop.mac is None:
            continue
        try:
            out[prop.name] = _mac_read(prop, getattr(obj, prop.mac).get())
        except Exception:  # noqa: BLE001
            continue
    return out


def mac_apply_line_spacing(fmt, rule, value) -> None:
    if rule is None:
        return
    fmt.line_spacing_rule.set(getattr(_k(), LINE_SPACING_MAC[rule]))
    if value is not None:
        fmt.line_spacing.set(value * 12.0 if rule == "multiple" else value)


def mac_read_line_spacing(fmt) -> dict:
    name = _kname(fmt.line_spacing_rule.get())
    rule = next((k for k, v in LINE_SPACING_MAC.items() if v == name), None)
    return describe_line_spacing(rule, fmt.line_spacing.get())
