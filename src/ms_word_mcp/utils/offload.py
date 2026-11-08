"""Run a blocking tool call on a thread without blocking the event loop.

Every tool handler is async, and everything under it waits on the Word
worker. Run inline, that would stop the server from answering a ping or a
cancellation while Word works. On cancellation the await is abandoned and
every call this tool queued but Word has not started is dropped, so nothing
lands after the client gave up. A call already inside Word cannot be stopped.
"""

from __future__ import annotations

from typing import Any, Callable, TypeVar

import anyio

from ..backend.worker import pending_calls

T = TypeVar("T")


async def run_offloaded(func: Callable[..., T], *args: Any) -> T:
    queued: list = []
    token = pending_calls.set(queued)
    try:
        return await anyio.to_thread.run_sync(func, *args, abandon_on_cancel=True)
    except anyio.get_cancelled_exc_class():
        for future in queued:
            future.cancel()
        raise
    finally:
        pending_calls.reset(token)
