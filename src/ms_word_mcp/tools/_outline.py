"""Finding a document's headings without asking Word about every paragraph.

Asking Word for one paragraph costs 5-20 ms through COM and Apple Events
alike, so reading every paragraph of a long document takes many seconds
(docs/PROBES.md, Reading). Two answers come back in one call each:

- the whole text, which split at its paragraph marks gives each paragraph's
  text and, in a plain document, its number; and
- Word's own list of headings, the one its cross-reference dialog shows
  (Document.GetCrossReferenceItems(wdRefTypeHeading)): each heading's text,
  indented two spaces a level, in order.

The split is a guess, never an answer: an equation's text holds paragraph
marks of its own and a section break ends a paragraph without one, so the
numbers drift (measured: 31 pieces for 19 paragraphs with two equations).
Each heading is therefore read back from Word at the number guessed, and
looked for nearby when it is not there; a heading not found means the caller
reads every paragraph instead. A number answered has always been read back.
"""

from __future__ import annotations

import re
from typing import Callable, Optional

#: How far below its guess a heading is looked for: equations add pieces to
#: the split, so the guess runs ahead of Word's numbering.
LOOK_BACK = 80
#: How far above: a section break ends a paragraph without a mark.
LOOK_AHEAD = 4

_FIELD_CODE = re.compile(r"\x13[^\x13\x14\x15]*\x14")
_SPACES = re.compile(r"[\s\x07\x0b\x0c\xa0]+")


def norm(text: str) -> str:
    """A paragraph's text as the heading list writes it.

    The list gives a tab, a line break or a run of spaces as one space and
    drops the spaces a heading begins with (measured on the Mac).
    """
    text = _FIELD_CODE.sub("", text or "").replace("\x13", "").replace("\x15", "")
    return _SPACES.sub(" ", text).strip()


def split_texts(whole: str) -> list[str]:
    """The whole text cut at its paragraph marks, normalized.

    A cell, and each row after its cells, ends in \\r\\x07; any other
    paragraph in \\r. The last paragraph's mark leaves nothing after it.
    """
    pieces = re.split(r"\r\x07?", whole or "")
    if pieces and pieces[-1] == "":
        pieces.pop()
    return [norm(p) for p in pieces]


def parse_item(item: str) -> tuple[int, str]:
    """(level, text) of one entry of Word's heading list."""
    item = item or ""
    text = item.lstrip(" ")
    return (len(item) - len(text)) // 2 + 1, norm(text)


def locate_headings(items: list[str], texts: list[str],
                    read_text: Callable[[int], Optional[str]],
                    read_level: Callable[[int], Optional[int]],
                    is_heading: Callable[[int], bool]) -> Optional[list[int]]:
    """The paragraph number of each heading in `items`, or None.

    `texts` is split_texts of the whole text. `read_text(i)`, `read_level(i)`
    and `is_heading(i)` ask Word about paragraph i (None when there is none).
    A number is kept only when Word's text starts with the heading's (the list
    shortens long headings), Word's level is the heading's, and the paragraph
    has a heading style: the same words at the same level can be a body
    paragraph given an outline level directly. None means a heading could not
    be found, and the caller reads every paragraph.
    """
    found: list[int] = []
    low, shift = 1, 0
    known: dict[int, Optional[str]] = {}

    def text_at(i: int) -> Optional[str]:
        if i not in known:
            raw = read_text(i)
            known[i] = None if raw is None else norm(raw)
        return known[i]

    def is_it(i: int, target: str, level: int) -> bool:
        got = text_at(i)
        return (got is not None and got.startswith(target) and read_level(i) == level
                and is_heading(i))

    for item in items:
        level, target = parse_item(item)
        if not target:
            continue
        guesses = [i for i in range(low, len(texts) + 1) if texts[i - 1].startswith(target)]
        hit = None
        # First where the split puts it, less how far it has run ahead so far.
        for g in guesses[:5]:
            i = g - shift
            if i >= low and is_it(i, target, level):
                hit, shift = i, g - i
                break
        # Then around each guess.
        if hit is None:
            for g in guesses[:3]:
                near = [g - shift - d for d in range(1, LOOK_BACK + 1)]
                near += [g - shift + d for d in range(1, LOOK_AHEAD + 1)]
                for i in near:
                    if i >= low and is_it(i, target, level):
                        hit, shift = i, g - i
                        break
                if hit is not None:
                    break
        if hit is None:
            return None
        found.append(hit)
        low = hit + 1
    return found
