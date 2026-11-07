"""What every tool module shares: running a call on Word, and its answer.

A tool answers JSON. Success is ``{"success": true, ...}``; a failure is
``{"error": "..."}`` with a sentence that says what to do instead, which
server.py marks as a tool error. Nothing is swallowed, and nothing falls back
quietly to something close to what was asked.
"""

from __future__ import annotations

import json
import logging
import os
from typing import Any, Callable

from pydantic import BaseModel, ConfigDict

from .. import backend

logger = logging.getLogger(__name__)

# The Mac halves written so far, by the name of their COM implementation.
# A tool whose implementation is not here refuses on macOS rather than
# sending Windows calls to Word for Mac.
MAC_IMPLS: dict[str, Callable[..., Any]] = {}


class Input(BaseModel):
    """The base of every tool's arguments: unknown keys are refused."""

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")


def mac_impl(com_impl: Callable[..., Any]):
    """Register the decorated function as `com_impl`'s Mac half."""

    def register(func):
        MAC_IMPLS[com_impl.__name__] = func
        return func

    return register


def _setting(name: str, default: bool) -> bool:
    value = os.getenv(name, "").strip().lower()
    if value in ("", "default"):
        return default
    if value in ("1", "true", "on", "yes"):
        return True
    if value in ("0", "false", "off", "no"):
        return False
    logger.warning("%s=%r is not a setting; use on or off. Using %s.",
                   name, value, "on" if default else "off")
    return default


#: Whether a writing call scrolls the document's window to what it changed.
FOLLOW_EDITS = _setting("WORD_FOLLOW_EDITS", True)


def error_text(exc: BaseException) -> str:
    from ..backend.com import ComError, describe_com_error

    if isinstance(exc, ComError):
        return f"Word refused the call: {describe_com_error(exc)}"
    return str(exc) or exc.__class__.__name__


def run(impl: Callable[..., Any], *args: Any, undo: str | None = None,
        idempotent: bool = False) -> str:
    """Run `impl(*args)` on Word and answer its result as JSON.

    `undo` names the undo step a writing call's changes are gathered under.
    """
    if backend.IS_MACOS:
        mac = MAC_IMPLS.get(impl.__name__)
        if mac is None:
            return json.dumps({
                "error": "This tool does not work on macOS yet.",
                "reason": ("ms-word-mcp is new, and this tool's Word for Mac half is "
                           "not written yet. It works on Windows."),
                "platform": backend.PLATFORM_NAME,
            }, ensure_ascii=False)
        impl = mac
    try:
        result = backend.word.execute(impl, *args, undo=undo, idempotent=idempotent)
    except Exception as exc:  # noqa: BLE001 - answered, never swallowed
        logger.debug("%s failed", getattr(impl, "__name__", impl), exc_info=True)
        return json.dumps({"error": error_text(exc)}, ensure_ascii=False)
    return json.dumps(result, ensure_ascii=False)


async def call(impl: Callable[..., Any], *args: Any, undo: str | None = None,
               idempotent: bool = False) -> str:
    """`run` on a thread, for a tool handler to await."""
    from functools import partial

    from ..utils.offload import run_offloaded

    return await run_offloaded(partial(run, impl, *args, undo=undo, idempotent=idempotent))


READ_ONLY = {"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True,
             "openWorldHint": False}
WRITES = {"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False,
          "openWorldHint": False}
DESTRUCTIVE = {"readOnlyHint": False, "destructiveHint": True, "idempotentHint": False,
               "openWorldHint": False}


def annotations(title: str, kind: dict) -> dict:
    return {"title": title, **kind}


def follow(doc, rng) -> None:
    """Scroll the document's window to `rng`, if following edits.

    The user's selection is left where it is: Window.ScrollIntoView moves
    the view only. When the document has a second window (View > New
    Window), the last one follows, so the first can stay where the user is
    writing. Any failure is ignored: the edit has landed either way.
    """
    if not FOLLOW_EDITS:
        return
    try:
        windows = doc.Windows
        count = windows.Count
        if count:
            windows.Item(count).ScrollIntoView(rng, True)
    except Exception:  # noqa: BLE001
        logger.debug("Could not scroll to the edit", exc_info=True)
