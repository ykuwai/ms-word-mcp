"""Bulleted and numbered lists, down to the bullet's character and color.

A list in Word is drawn from a list template: up to nine levels, each with
its bullet or number format, the bullet's own font (and so its color and
size), and the positions of the number and the text. Changing a template
the document's other lists share would change them too, so word_apply_list
makes a template of its own for each list it applies and leaves every other
list as it was (docs/reference/word-object-model.md, D.1).
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from . import _props as P
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, annotations, call, follow
from ._word import check_span, clean_text, paragraph, paragraphs_in, span_range

# WdListNumberStyle, by the names the tools take (reference B.2).
NUMBER_STYLES = {
    "bullet": 23, "arabic": 0, "upper_roman": 1, "lower_roman": 2, "upper_letter": 3,
    "lower_letter": 4, "ordinal": 5, "cardinal_text": 6, "ordinal_text": 7,
    "arabic_full_width": 14, "arabic_leading_zero": 22, "number_in_circle": 18,
    "kanji": 10, "kanji_digit": 11, "kanji_traditional": 16, "aiueo": 20,
    "aiueo_half_width": 12, "iroha": 21, "iroha_half_width": 13, "legal": 253,
    "none": 255,
}
TRAILING = {"tab": 0, "space": 1, "none": 2}
LEVEL_ALIGNMENT = {"left": 0, "center": 1, "right": 2}
LIST_TYPES = {0: "none", 1: "listnum_field", 2: "bullet", 3: "simple_numbering",
              4: "outline_numbering", 5: "mixed_numbering", 6: "picture_bullet"}
WD_WORD10_LIST_BEHAVIOR = 2
WD_LIST_APPLY_TO_WHOLE_LIST = 0


class BulletFont(Input):
    """The bullet's or number's own character formatting."""

    name: str | None = Field(None, description=(
        "Font of the bullet or number. A Unicode bullet such as ● ■ ◆ ✓ shows in "
        "most text fonts; Symbol and Wingdings have their own."))
    color: str | None = Field(None, pattern=r"^#?[0-9A-Fa-f]{6}$", description="#RRGGBB.")
    size: float | None = Field(None, gt=0, le=1638, description="Points.")
    bold: bool | None = None


class LevelSpec(Input):
    level: int = Field(1, ge=1, le=9, description="The list level this sets, 1 to 9.")
    bullet: str | None = Field(None, min_length=1, max_length=1, description=(
        "The bullet character, one character: ●, ■, ◆, ・, ✓, –. Makes the level "
        "a bulleted one."))
    number_style: Literal[tuple(NUMBER_STYLES)] | None = Field(None, description=(  # type: ignore[valid-type]
        "How the number is written: arabic (1), arabic_full_width (１), "
        "number_in_circle (①), kanji (一), aiueo (ア), iroha (イ), upper_roman, "
        "lower_letter and the rest."))
    number_format: str | None = Field(None, min_length=1, max_length=40, description=(
        "The text around the number: %1 is this level's number, %2 the second "
        "level's and so on. \"%1.\" gives 1., \"(%1)\" gives (1), \"第%1章\" gives "
        "第1章, \"%1.%2\" gives 1.1."))
    font: BulletFont | None = None
    number_position: float | None = Field(None, ge=-1584, le=1584, description=(
        "Where the bullet or number sits, points from the margin."))
    text_position: float | None = Field(None, ge=-1584, le=1584, description=(
        "Where the text and its wrapped lines start, points from the margin."))
    tab_position: float | None = Field(None, ge=0, le=1584, description=(
        "The tab stop after the number, points."))
    trailing: Literal[tuple(TRAILING)] | None = Field(None, description=(  # type: ignore[valid-type]
        "What follows the number: tab, space or none."))
    alignment: Literal[tuple(LEVEL_ALIGNMENT)] | None = Field(None, description=(  # type: ignore[valid-type]
        "How the number aligns at number_position."))
    start_at: int | None = Field(None, ge=0, le=32767)
    linked_style: str | None = Field(None, min_length=1, description=(
        "A paragraph style whose paragraphs take this level, such as \"Heading 1\" "
        "for numbered headings."))

    @model_validator(mode="after")
    def _bullet_or_number(self):
        if self.bullet is not None and self.number_style not in (None, "bullet"):
            raise ValueError("Give `bullet` for a bulleted level or `number_style` for "
                             "a numbered one, not both.")
        if self.bullet is not None and self.number_format is not None:
            raise ValueError("A bulleted level's format is its one character: give "
                             "`bullet` only.")
        return self


class ApplyListInput(Input):
    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")
    levels: list[LevelSpec] = Field(..., min_length=1, max_length=9, description=(
        "How each level looks. A plain bulleted list is [{\"bullet\": \"●\"}]; a "
        "numbered one [{\"number_style\": \"arabic\", \"number_format\": \"%1.\"}]."))
    level: int | None = Field(None, ge=1, le=9, description=(
        "Put these paragraphs at this level of the list."))
    continue_previous: bool = Field(False, description=(
        "Continue the numbering of the list before, rather than start at 1."))

    @model_validator(mode="after")
    def _distinct_levels(self):
        numbers = [lv.level for lv in self.levels]
        if len(numbers) != len(set(numbers)):
            raise ValueError("Each level may appear once in `levels`.")
        return self


def _level_spec_defaults(spec: dict) -> dict:
    if spec.get("bullet") is not None:
        spec = {**spec, "number_style": "bullet", "number_format": spec["bullet"]}
    return spec


def _set_level(doc, lv, spec: dict, refused: dict) -> None:
    def put(key, member, value):
        try:
            setattr(lv, member, value)
        except Exception as exc:  # noqa: BLE001 - reported to the caller
            from ._common import error_text

            refused[f"level {spec['level']} {key}"] = error_text(exc)

    # Measured on Word 16 for Windows and Mac: a level's number style turns
    # to bullet only once its number format is the bullet character, and a
    # number format holding %1 is taken only under a numbering style. So a
    # bullet's format goes first, and a number's style.
    style = spec.get("number_style")
    steps = [("number_format", "NumberFormat", spec.get("number_format")),
             ("number_style", "NumberStyle",
              None if style is None else NUMBER_STYLES[style])]
    if style not in (None, "bullet"):
        steps.reverse()
    for key, member, value in steps:
        if value is not None:
            put(key, member, value)
    for key, member in (("number_position", "NumberPosition"),
                        ("text_position", "TextPosition"),
                        ("tab_position", "TabPosition"), ("start_at", "StartAt")):
        if spec.get(key) is not None:
            put(key, member, spec[key])
    if spec.get("trailing") is not None:
        put("trailing", "TrailingCharacter", TRAILING[spec["trailing"]])
    if spec.get("alignment") is not None:
        put("alignment", "Alignment", LEVEL_ALIGNMENT[spec["alignment"]])
    if spec.get("linked_style") is not None:
        from ._word import resolve_style, style_name

        put("linked_style", "LinkedStyle", style_name(resolve_style(doc, spec["linked_style"])))
    font = spec.get("font") or {}
    font_spec = {k: v for k, v in {"latin_font": font.get("name"),
                                   "east_asian_font": None, "color": font.get("color"),
                                   "size": font.get("size"),
                                   "bold": font.get("bold")}.items() if v is not None}
    if font.get("name") is not None:
        font_spec["latin_font_extended"] = font["name"]
    for name, why in P.com_apply(lv.Font, P.FONT_PROPS, font_spec).items():
        refused[f"level {spec['level']} font {name}"] = why


def describe_level(lv) -> dict:
    """A list level, in the names word_apply_list takes."""
    style = lv.NumberStyle
    out = {"level": lv.Index,
           "number_style": next((k for k, v in NUMBER_STYLES.items() if v == style), style),
           "number_format": lv.NumberFormat}
    if style == NUMBER_STYLES["bullet"]:
        out["bullet"] = lv.NumberFormat
    font = P.com_read(lv.Font, P.FONT_PROPS)
    out["font"] = {"name": font.get("latin_font"), "color": font.get("color"),
                   "size": font.get("size"), "bold": font.get("bold")}
    for key, member in (("number_position", "NumberPosition"),
                        ("text_position", "TextPosition"), ("start_at", "StartAt")):
        out[key] = getattr(lv, member)
    tab = lv.TabPosition
    out["tab_position"] = None if tab in (999999, P.WD_UNDEFINED) else tab
    out["trailing"] = next((k for k, v in TRAILING.items() if v == lv.TrailingCharacter),
                           lv.TrailingCharacter)
    out["alignment"] = next((k for k, v in LEVEL_ALIGNMENT.items() if v == lv.Alignment),
                            lv.Alignment)
    out["linked_style"] = lv.LinkedStyle or None
    return out


def describe_paragraph_list(index, p) -> dict:
    fmt = p.Range.ListFormat
    kind = LIST_TYPES.get(fmt.ListType, fmt.ListType)
    out = {"index": index, "text": clean_text(p.Range.Text)[:80], "list": kind}
    if kind != "none":
        out["level"] = fmt.ListLevelNumber
        out["list_string"] = fmt.ListString
    return out


def _apply_list_impl(start, end, levels, level, continue_previous):
    doc = backend.word.document()
    start, end = check_span(doc, start, end)
    specs = [_level_spec_defaults(dict(s)) for s in levels]
    outline = len(specs) > 1 or any(s["level"] > 1 for s in specs) or level not in (None, 1)
    template = doc.ListTemplates.Add(outline)
    refused: dict = {}
    for spec in specs:
        _set_level(doc, template.ListLevels.Item(spec["level"]), spec, refused)
    rng = span_range(doc, start, end)
    rng.ListFormat.ApplyListTemplate(ListTemplate=template,
                                     ContinuePreviousList=continue_previous,
                                     ApplyTo=WD_LIST_APPLY_TO_WHOLE_LIST,
                                     DefaultListBehavior=WD_WORD10_LIST_BEHAVIOR)
    if level is not None:
        for _, p in paragraphs_in(doc, start, end):
            p.Range.ListFormat.ListLevelNumber = level
    out = [describe_paragraph_list(i, p) for i, p in paragraphs_in(doc, start, end)]
    used = paragraph(doc, start).Range.ListFormat.ListTemplate
    described = [describe_level(used.ListLevels.Item(s["level"])) for s in specs]
    warnings = [f"{key} was refused: {why}" for key, why in refused.items()]
    for spec, got in zip(specs, described):
        for key in ("bullet", "number_format", "number_position", "text_position",
                    "start_at"):
            if spec.get(key) is not None and got.get(key) != spec[key]:
                warnings.append(f"Level {spec['level']} {key} reads back as "
                                f"{got.get(key)!r}, not {spec[key]!r}.")
        for key, wanted in (spec.get("font") or {}).items():
            have = got["font"].get(key)
            if wanted is not None and key != "color" and have != wanted:
                warnings.append(f"Level {spec['level']} font {key} reads back as "
                                f"{have!r}, not {wanted!r}.")
            if key == "color" and wanted is not None and (have or "").upper() != \
                    "#" + wanted.lstrip("#").upper():
                warnings.append(f"Level {spec['level']} font color reads back as "
                                f"{have!r}, not {wanted!r}.")
    off = [p["index"] for p in out if p["list"] == "none"]
    if off:
        warnings.append(f"Paragraphs {off} read back without a list.")
    follow(doc, rng)
    result = {"success": True, "paragraphs": out, "levels": described}
    if warnings:
        result["warnings"] = warnings
    return result


class ListSpanInput(Input):
    start: int = Field(..., ge=1, description="First paragraph, numbered from 1.")
    end: int | None = Field(None, ge=1, description="Last paragraph. Leave out for one.")


class SetListLevelInput(ListSpanInput):
    level: int = Field(..., ge=1, le=9, description="The list level, 1 to 9.")


def _set_list_level_impl(start, end, level):
    doc = backend.word.document()
    start, end = check_span(doc, start, end)
    for _, p in paragraphs_in(doc, start, end):
        if p.Range.ListFormat.ListType == 0:
            raise ValueError(f"Paragraph {_} is not in a list. Make it one with "
                             "word_apply_list first.")
    for _, p in paragraphs_in(doc, start, end):
        p.Range.ListFormat.ListLevelNumber = level
    out = [describe_paragraph_list(i, p) for i, p in paragraphs_in(doc, start, end)]
    result = {"success": True, "paragraphs": out}
    off = [p["index"] for p in out if p.get("level") != level]
    if off:
        result["warnings"] = [f"Paragraphs {off} read back at another level."]
    follow(doc, span_range(doc, start, end))
    return result


def _remove_list_impl(start, end):
    doc = backend.word.document()
    start, end = check_span(doc, start, end)
    rng = span_range(doc, start, end)
    rng.ListFormat.RemoveNumbers()
    out = [describe_paragraph_list(i, p) for i, p in paragraphs_in(doc, start, end)]
    result = {"success": True, "paragraphs": out}
    left = [p["index"] for p in out if p["list"] != "none"]
    if left:
        result["warnings"] = [f"Paragraphs {left} are still in a list."]
    follow(doc, rng)
    return result


class GetListInput(Input):
    index: int = Field(..., ge=1, description="A paragraph of the list.")


def _get_list_impl(index):
    doc = backend.word.document()
    p = paragraph(doc, index)
    info = describe_paragraph_list(index, p)
    if info["list"] == "none":
        return {"success": True, "paragraph": info}
    template = p.Range.ListFormat.ListTemplate
    levels = template.ListLevels
    return {"success": True, "paragraph": info,
            "levels": [describe_level(levels.Item(i)) for i in range(1, levels.Count + 1)]}


def register(mcp) -> None:
    @mcp.tool(name="word_apply_list", annotations=annotations("Make a list", WRITES))
    async def word_apply_list(params: ApplyListInput) -> str:
        """Make paragraphs a bulleted or numbered list, with each level's bullet
        character or number format, the bullet's own font, color and size, and
        the positions of the number and the text.

        The list gets a template of its own, so no other list in the document
        changes. Reads back each paragraph's list text (ListString) and each
        level as Word holds it.
        """
        p = params
        return await call(_apply_list_impl, p.start, p.end,
                          [lv.model_dump() for lv in p.levels], p.level,
                          p.continue_previous, undo="Apply list")

    @mcp.tool(name="word_set_list_level", annotations=annotations("List level", WRITES))
    async def word_set_list_level(params: SetListLevelInput) -> str:
        """Move list paragraphs to another level (indent or outdent them)."""
        p = params
        return await call(_set_list_level_impl, p.start, p.end, p.level,
                          undo="List level")

    @mcp.tool(name="word_remove_list", annotations=annotations("Remove a list", DESTRUCTIVE))
    async def word_remove_list(params: ListSpanInput) -> str:
        """Take paragraphs out of their list: their bullets or numbers go, the
        text stays."""
        return await call(_remove_list_impl, params.start, params.end, undo="Remove list")

    @mcp.tool(name="word_get_list", annotations=annotations("Read a list", READ_ONLY))
    async def word_get_list(params: GetListInput) -> str:
        """The list a paragraph is in: its type, the paragraph's level and list
        text, and every level's bullet or number format, font and positions."""
        return await call(_get_list_impl, params.index)
