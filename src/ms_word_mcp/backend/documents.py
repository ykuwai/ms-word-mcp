"""Picking the document a call works on, the same way on both platforms.

A call names its document with `document` (a full path, a file name, or a
file name without its extension). One that names none works on the session
target set by word_activate_document, and failing that on Word's active
document. A name that matches no open document, or several, is an error: it
never falls back to some other document.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from urllib.parse import unquote


@dataclass(frozen=True)
class OpenDocument:
    """What is needed to choose between open documents."""

    index: int  # 1-based, as Word counts them
    name: str
    full_name: str


def full_name_key(full_name: str) -> str:
    """A FullName compared without case, and with a URL's escapes undone."""
    text = full_name.strip()
    if text.lower().startswith(("http://", "https://")):
        text = unquote(text)
    return text.replace("\\", "/").lower()


def _stem(name: str) -> str:
    return os.path.splitext(name)[0]


def pick_document(candidates: list[OpenDocument], wanted: str) -> OpenDocument:
    """The one open document `wanted` names. Raises ValueError otherwise."""
    wanted = wanted.strip()
    key = full_name_key(wanted)
    exact = [d for d in candidates if full_name_key(d.full_name) == key]
    if len(exact) == 1:
        return exact[0]
    lowered = wanted.lower()
    by_name = [d for d in candidates if d.name.lower() == lowered]
    if not by_name:
        by_name = [d for d in candidates if _stem(d.name).lower() == lowered]
    if len(by_name) == 1:
        return by_name[0]
    open_names = ", ".join(d.name for d in candidates) or "none"
    if by_name:
        raise ValueError(
            f"More than one open document is called '{wanted}'. Pass its full path "
            f"instead: {', '.join(d.full_name for d in by_name)}.")
    raise ValueError(
        f"No open document is called '{wanted}'. Open documents: {open_names}. "
        "Open it with word_open_document first.")
