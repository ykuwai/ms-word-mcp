"""Small helpers for walking Word's object model over COM.

Paragraphs are addressed by their 1-based index in the document, the way
Word counts them. Every writing call reports the indices it changed, read
back from Word after the change.
"""

from __future__ import annotations

import re
from collections import Counter

# Values Word uses, written out here so nothing depends on makepy constants.
WD_UNDEFINED = 9999999
WD_REVISION_DELETE = 2
WD_REVISION_MOVED_FROM = 14
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
# up by this constant first: "Heading 1" works on every Word. The table is
# every member of WdBuiltinStyle, named as Microsoft's page describes it
# (docs/reference/word-object-model.md, B.2). "No Spacing" has none, so it is
# found by name.
BUILTIN_STYLES = {
    'normal': -1,
    'heading 1': -2,
    'heading 2': -3,
    'heading 3': -4,
    'heading 4': -5,
    'heading 5': -6,
    'heading 6': -7,
    'heading 7': -8,
    'heading 8': -9,
    'heading 9': -10,
    'index 1': -11,
    'index 2': -12,
    'index 3': -13,
    'index 4': -14,
    'index 5': -15,
    'index 6': -16,
    'index 7': -17,
    'index 8': -18,
    'index 9': -19,
    'toc 1': -20,
    'toc 2': -21,
    'toc 3': -22,
    'toc 4': -23,
    'toc 5': -24,
    'toc 6': -25,
    'toc 7': -26,
    'toc 8': -27,
    'toc 9': -28,
    'normal indent': -29,
    'footnote text': -30,
    'comment text': -31,
    'header': -32,
    'footer': -33,
    'index heading': -34,
    'caption': -35,
    'table of figures': -36,
    'envelope address': -37,
    'envelope return': -38,
    'footnote reference': -39,
    'comment reference': -40,
    'line number': -41,
    'page number': -42,
    'endnote reference': -43,
    'endnote text': -44,
    'table of authorities': -45,
    'macro text': -46,
    'toa heading': -47,
    'list': -48,
    'list bullet': -49,
    'list number': -50,
    'list 2': -51,
    'list 3': -52,
    'list 4': -53,
    'list 5': -54,
    'list bullet 2': -55,
    'list bullet 3': -56,
    'list bullet 4': -57,
    'list bullet 5': -58,
    'list number 2': -59,
    'list number 3': -60,
    'list number 4': -61,
    'list number 5': -62,
    'title': -63,
    'closing': -64,
    'signature': -65,
    'default paragraph font': -66,
    'body text': -67,
    'body text indent': -68,
    'list continue': -69,
    'list continue 2': -70,
    'list continue 3': -71,
    'list continue 4': -72,
    'list continue 5': -73,
    'message header': -74,
    'subtitle': -75,
    'salutation': -76,
    'date': -77,
    'body text first indent': -78,
    'body text first indent 2': -79,
    'note heading': -80,
    'body text 2': -81,
    'body text 3': -82,
    'body text indent 2': -83,
    'body text indent 3': -84,
    'block text': -85,
    'hyperlink': -86,
    'followed hyperlink': -87,
    'strong': -88,
    'emphasis': -89,
    'document map': -90,
    'plain text': -91,
    'html acronym': -96,
    'html address': -97,
    'html cite': -98,
    'html code': -99,
    'html definition': -100,
    'html keyboard': -101,
    'html preformatted': -102,
    'html sample': -103,
    'html typewriter': -104,
    'html variable': -105,
    'light shading': -159,
    'light list': -160,
    'light grid': -161,
    'medium shading 1': -162,
    'medium shading 2': -163,
    'medium list 1': -164,
    'medium list 2': -165,
    'medium grid 1': -166,
    'medium grid 2': -167,
    'medium grid 3': -168,
    'dark list': -169,
    'colorful shading': -170,
    'colorful list': -171,
    'colorful grid': -172,
    'light shading accent 1': -173,
    'light list accent 1': -174,
    'light grid accent 1': -175,
    'medium shading 1 accent 1': -176,
    'medium shading 2 accent 1': -177,
    'medium list 1 accent 1': -178,
    'list paragraph': -180,
    'quote': -181,
    'intense quote': -182,
    'subtle emphasis': -261,
    'intense emphasis': -262,
    'subtle reference': -263,
    'intense reference': -264,
    'book title': -265,
}

ALIGNMENTS = {"left": 0, "center": 1, "right": 2, "justify": 3, "distribute": 4}
ALIGNMENT_NAMES = {v: k for k, v in ALIGNMENTS.items()}


_RUBY_CODE = re.compile(r"\\s\\up\s*\d+\((.*?)\),(.*)\)\s*$", re.S)


def ruby_of(code: str) -> tuple[str, str] | None:
    """(text, reading) from the EQ field Word stores ruby in, or None."""
    # The pattern backtracks, and a document can hold a field code of any length.
    m = _RUBY_CODE.search(code) if code and len(code) <= 4000 else None
    return (m.group(2), m.group(1)) if m else None


def with_ruby(text: str, codes: list[str]) -> tuple[str, list[dict]]:
    """A paragraph's text with its ruby fields' words in place of their \\x15 marks.

    Word answers a single \\x15 for text that has ruby (measured on both
    platforms). `codes` are the field codes of the paragraph, in order.
    """
    found = [r for r in (ruby_of(c) for c in codes) if r]
    rubies = [{"text": base, "ruby": reading} for base, reading in found]
    for base, _ in found:
        text = text.replace("\x15", base, 1)
    return text.replace("\x15", ""), rubies


def expand_ruby(info: dict, read_fields) -> None:
    """Put the ruby words in a paragraph's description in place of their \\x15 marks.

    `read_fields()` answers the paragraph's fields in order, each as (code,
    whether a tracked deletion covers it). A field marked deleted stays in
    Word's list (measured on both platforms), so its word belongs to
    `deleted_text`, and only the others to `text`.
    """
    if "\x15" not in info["text"] + info.get("deleted_text", ""):
        return
    fields = read_fields()
    info["text"], rubies = with_ruby(info["text"], [c for c, gone in fields if not gone])
    if rubies:
        info["ruby"] = rubies
    if "deleted_text" in info:
        info["deleted_text"], _ = with_ruby(info["deleted_text"],
                                            [c for c, gone in fields if gone])


def check_spans(text: str, spans: list[str]) -> None:
    """Refuse spans that are not all in `text`, one after the other."""
    cursor = 0
    for span in spans:
        at = text.find(span, cursor)
        if at < 0:
            raise ValueError(f"'{span}' is not in the paragraph after what comes before it. "
                             "Give the spans in the order they come. Nothing was changed.")
        cursor = at + len(span)


def ruby_to_drop(raw: str, codes: list[str], spans: list[str]) -> list[tuple[int, str]]:
    """Which ruby fields give way to `spans`: (place among `codes`, its word) of each.

    `raw` is the paragraph's text, with a \\x15 for each ruby word, and `codes`
    its field codes in order. The spans are checked against the text as it
    will be once those fields are plain text, so a request that cannot be
    carried out is refused before any reading is taken off.
    """
    wanted = Counter(spans)
    found = []
    for i, code in enumerate(codes):
        got = ruby_of(code)
        if got is None:
            continue
        if wanted[got[0]] > 0:
            wanted[got[0]] -= 1
            found.append((i, got[0]))
            raw = raw.replace("\x15", got[0], 1)
        else:
            raw = raw.replace("\x15", "\x00", 1)  # stays a field: no span can cross it
    check_spans(raw, spans)
    return found


def locate_spans(raw: str, spans: list[str], read, start: int, hidden: int) -> list[tuple[int, int]]:
    """Word's (start, end) of each span, in order, each read back from Word.

    `raw` is the paragraph's text; Word's positions run ahead of it by up to
    `hidden` after ruby or another field that hides its characters.
    """
    def u16(text):
        return len(text.encode("utf-16-le")) // 2

    check_spans(raw, spans)
    out, cursor, drift = [], 0, 0
    for span in spans:
        at = raw.find(span, cursor)
        base, n = start + u16(raw[:at]), u16(span)
        for shift in range(drift, hidden + 1):
            if read(base + shift, base + shift + n) == span:
                out.append((base + shift, base + shift + n))
                drift = shift
                break
        else:
            raise ValueError(f"'{span}' could not be found in Word at its place. Nothing "
                             "was changed.")
        cursor = at + len(span)
    return out


def clean_text(text: str) -> str:
    """A range's text without its paragraph mark or cell marker."""
    return (text or "").rstrip("\r\x07")


def tracked_text(read, start: int, end: int, spans: list[tuple[int, int]]) -> tuple[str, dict]:
    """The text of [start, end) as Word shows it, and what is struck out in it.

    A tracked deletion stays in a range's text while the markup is shown.
    `read(a, b)` answers the text between two positions; `spans` are the
    deletions. Answers the text left, and the fields to add to a paragraph's
    description: `deleted_text`, and `deleted` when all of it is marked.
    """
    kept, gone, at, mark_gone = [], [], start, False
    for a, b in sorted(spans):
        a, b = max(a, at), min(b, end)
        if a >= b:
            continue
        kept.append(read(at, a) or "" if a > at else "")
        gone.append(read(a, b) or "")
        at = b
        mark_gone = mark_gone or b == end
    if at < end:
        kept.append(read(at, end) or "")
    text = "".join(kept)
    if not gone:
        return text, {}
    extra = {"deleted_text": clean_text("".join(gone))}
    if mark_gone and not clean_text(text):
        extra["deleted"] = True
    return text, extra


def inside_spans(start: int, end: int, spans: list[tuple[int, int]]) -> bool:
    """Whether [start, end) lies within one of the spans."""
    return any(a <= start and end <= b for a, b in spans)


def deletion_spans(rng) -> list[tuple[int, int]]:
    """(start, end) of each tracked deletion in a range."""
    revisions = rng.Revisions
    count = revisions.Count
    if not count:
        return []
    out = []
    for i in range(1, count + 1):
        revision = revisions.Item(i)
        if revision.Type in (WD_REVISION_DELETE, WD_REVISION_MOVED_FROM):
            out.append((revision.Range.Start, revision.Range.End))
    return out


def visible_text(rng) -> tuple[str, dict]:
    """A range's text without its tracked deletions (see tracked_text)."""
    spans = deletion_spans(rng)
    if not spans:
        return rng.Text, {}
    return tracked_text(lambda a, b: rng.Document.Range(a, b).Text, rng.Start, rng.End, spans)


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
            f"No style called '{name}' in this document. Use a built-in style's English "
            "name (Normal, Heading 1, Title, Body Text, Caption, TOC 1, Header...: every "
            "Word built-in style), or a name word_list_styles gives.") from None


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


def _field_codes(rng) -> list[tuple[str, bool]]:
    """Each field of a range: its code, and whether a tracked deletion covers it."""
    spans = deletion_spans(rng)
    fields = rng.Fields
    out = []
    for i in range(1, fields.Count + 1):
        code = fields.Item(i).Code
        out.append((code.Text, inside_spans(code.Start, code.End, spans)))
    return out


def paragraph_info(index: int, p, detail: bool = False) -> dict:
    rng = p.Range
    text, tracked = visible_text(rng)
    info = {"index": index, "text": clean_text(text), "style": style_name(p.Style), **tracked}
    expand_ruby(info, lambda: _field_codes(rng))
    level = p.OutlineLevel
    if level != WD_OUTLINE_BODY:
        info["outline_level"] = level
    if detail:
        fmt = p.Format
        info["alignment"] = ALIGNMENT_NAMES.get(fmt.Alignment, fmt.Alignment)
        info["fonts"] = fonts_of(rng)
        font = rng.Font
        # Bold, italic and color too: a paragraph made to look like a heading
        # with direct formatting is found by these (asked for in a trial).
        color = undefined(font.Color)
        info["fonts"].update({
            "bold": flag(font.Bold), "italic": flag(font.Italic),
            "color": (None if color is None or color < 0 or color > 0xFFFFFF else
                      f"#{color & 0xFF:02X}{(color >> 8) & 0xFF:02X}{(color >> 16) & 0xFF:02X}")})
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
