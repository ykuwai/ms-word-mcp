"""Undo and redo.

Each writing tool call is one undo step in Word, named after the tool, so
word_undo with times=1 takes back exactly the last call.
"""

from __future__ import annotations

from pydantic import Field

from .. import backend
from ._common import WRITES, Input, annotations, call


class UndoInput(Input):
    times: int = Field(1, ge=1, le=100, description="How many steps to take back.")


def _undo_impl(times: int):
    doc = backend.word.document()
    done = 0
    for _ in range(times):
        if not doc.Undo(1):
            break
        done += 1
    result = {"success": True, "undone": done}
    if done < times:
        result["warnings"] = [f"Only {done} of {times} steps could be undone."]
    return result


def _redo_impl(times: int):
    doc = backend.word.document()
    done = 0
    for _ in range(times):
        if not doc.Redo(1):
            break
        done += 1
    result = {"success": True, "redone": done}
    if done < times:
        result["warnings"] = [f"Only {done} of {times} steps could be redone."]
    return result


def register(mcp) -> None:
    @mcp.tool(name="word_undo", annotations=annotations("Undo", WRITES))
    async def word_undo(params: UndoInput) -> str:
        """Undo the last steps in the document. Each writing tool call is one step."""
        return await call(_undo_impl, params.times)

    @mcp.tool(name="word_redo", annotations=annotations("Redo", WRITES))
    async def word_redo(params: UndoInput) -> str:
        """Redo steps that word_undo took back."""
        return await call(_redo_impl, params.times)
