"""Small helpers for walking Word's object model over COM.

Paragraphs are addressed by their 1-based index in the document, the way
Word counts them. Every writing call reports the indices it changed, read
back from Word after the change.
"""

from __future__ import annotations

# Values Word uses, written out here so nothing depends on makepy constants.
WD_UNDEFINED = 9999999
WD_CHARACTER = 1
WD_COLLAPSE_END = 0
WD_FIND_STOP = 0
WD_REPLACE_ALL = 2
WD_OUTLINE_BODY = 10
WD_STATISTIC_WORDS = 0
WD_STATISTIC_PAGES = 2
WD_STATISTIC_CHARACTERS = 3

# Built-in styles by their English name, as WdBuiltinStyle. A localized Word
# names its styles in its own language (見出し 1, 標準), so a style is looked
# up by this constant first: "Heading 1" works on every Word. Every value is
# checked in docs/reference/word-object-model.md (B.1). "No Spacing" has no
# WdBuiltinStyle member (-158 is wdStyleNormalObject), so it is found by name.
BUILTIN_STYLES = {
    "normal": -1,
    "heading 1": -2, "heading 2": -3, "heading 3": -4, "heading 4": -5,
    "heading 5": -6, "heading 6": -7, "heading 7": -8, "heading 8": -9,
    "heading 9": -10,
    "caption": -35,
    "title": -63,
    "body text": -67,
    "subtitle": -75,
    "list paragraph": -180,
    "quote": -181,
    "intense quote": -182,
    "list bullet": -49,
    "list number": -50,
}

ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3, "distribute": 4}
ALIGNMENT_NAMES = {v: k for k, v in ALIGNMENTS.items()}


def clean_text(text: str) -> str:
    """A range's text without its paragraph mark or cell marker."""
    return (text or "").rstrip("\r\x07")


def style_name(style) -> str | None:
    """The local name of a Style object, or of whatever Word answered."""
    if style is None:
        return None
    name = getattr(style, "NameLocal", None)
    return str(name) if name is not None else str(style)


def flag(value):
    """A True/False/wdUndefined Long as true, false or None (mixed).

    Font.Bold and its kind are Longs that hold True (-1), False (0) or
    wdUndefined; read through pywin32 they come back as numbers.
    """
    value = undefined(value)
    return None if value is None else bool(value)


def undefined(value):
    """None where Word answers wdUndefined (a range holding mixed values)."""
    return None if value in (WD_UNDEFINED, -WD_UNDEFINED, "") else value


def paragraph(doc, index: int):
    count = doc.Paragraphs.Count
    if not 1 <= index <= count:
        raise ValueError(
            f"There is no paragraph {index}: the document has {count} paragraphs "
            "(numbered from 1).")
    return doc.Paragraphs.Item(index)


def check_span(doc, start: int, end: int | None) -> tuple[int, int]:
    """The paragraph span [start, end] checked against the document."""
    count = doc.Paragraphs.Count
    end = start if end is None else end
    if end < start:
        raise ValueError(f"end ({end}) comes before start ({start}).")
    if not 1 <= start <= count or end > count:
        raise ValueError(
            f"Paragraphs {start} to {end} are not all in the document: it has "
            f"{count} paragraphs (numbered from 1).")
    return start, end


def span_range(doc, start: int, end: int):
    """A Range from the start of paragraph `start` to the end of `end`."""
    return doc.Range(paragraph(doc, start).Range.Start, paragraph(doc, end).Range.End)


def paragraphs_in(doc, start: int, end: int):
    """Yield (index, Paragraph) from `start` to `end`, walking with Next."""
    p = paragraph(doc, start)
    for index in range(start, end + 1):
        if p is None:
            return
        yield index, p
        p = p.Next() if index < end else None


def index_of(doc, rng) -> int:
    """The index of the paragraph `rng` starts in.

    The range counted runs through the first character of `rng`: one that
    ended just before it would leave out a paragraph `rng` starts.
    """
    return doc.Range(0, rng.Start + 1).Paragraphs.Count


def resolve_style(doc, name: str):
    """The Style called `name`: a built-in English name, or the document's own."""
    key = " ".join(name.lower().split())
    if key in BUILTIN_STYLES:
        return doc.Styles.Item(BUILTIN_STYLES[key])
    try:
        return doc.Styles.Item(name)
    except Exception:  # noqa: BLE001 - answered below with what exists
        raise ValueError(
            f"No style called '{name}' in this document. Use a built-in English name "
            f"({', '.join(sorted(k.title() for k in BUILTIN_STYLES))}) or one that "
            "word_list_styles gives.") from None


def fonts_of(rng) -> dict:
    """The fonts of a range, one per script. None means the range is mixed."""
    font = rng.Font
    return {
        "latin": font.NameAscii or None,
        "east_asian": font.NameFarEast or None,
        "complex_script": font.NameBi or None,
        "size": undefined(font.Size),
        "size_complex_script": undefined(font.SizeBi),
    }


def paragraph_info(index: int, p, detail: bool = False) -> dict:
    rng = p.Range
    info = {"index": index, "text": clean_text(rng.Text), "style": style_name(p.Style)}
    level = p.OutlineLevel
    if level != WD_OUTLINE_BODY:
        info["outline_level"] = level
    if detail:
        fmt = p.Format
        info["alignment"] = ALIGNMENT_NAMES.get(fmt.Alignment, fmt.Alignment)
        info["fonts"] = fonts_of(rng)
        rule = fmt.LineSpacingRule
        info["line_spacing_rule"] = LINE_SPACING_NAMES.get(rule, rule)
        lines = spacing_in_lines(rule, fmt.LineSpacing)
        if lines is not None:
            info["line_spacing_lines"] = lines
        else:
            info["line_spacing_pt"] = fmt.LineSpacing
        info["space_before"] = fmt.SpaceBefore
        info["space_after"] = fmt.SpaceAfter
    return info


# WdLineSpacing
LINE_SPACING = {"single": 0, "one_and_half": 1, "double": 2, "at_least": 3,
                "exactly": 4, "multiple": 5}
LINE_SPACING_NAMES = {v: k for k, v in LINE_SPACING.items()}
_FIXED_LINES = {0: 1.0, 1: 1.5, 2: 2.0}


def spacing_in_lines(rule, points) -> float | None:
    """Line spacing in lines for the rules measured in lines, else None.

    Word stores a multiple in points, 12 to a line, and turns a multiple of
    exactly 1, 1.5 or 2 lines into the single, one-and-a-half or double rule
    (measured on Word 16 for Windows and Mac), so those read back as lines too.
    """
    if rule in _FIXED_LINES:
        return _FIXED_LINES[rule]
    if rule == LINE_SPACING["multiple"]:
        return round(points / 12.0, 3)
    return None
