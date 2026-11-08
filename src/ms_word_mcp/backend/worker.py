"""One thread that talks to Word, and the queue that feeds it.

Word answers one caller at a time. Calls made from several threads at once
come back with errors that are not true while the change still lands, so
every call goes through one worker thread, in the order it was asked for. On
Windows that thread is also the COM apartment Word's objects belong to.

The rules this keeps, all learned on ppt-mcp:

- Time spent waiting in the queue is not charged to the call. A call waits
  up to QUEUE_WAIT to start, then up to its own budget to finish.
- A call the caller gave up on before it started is never applied later.
- A call that failed part-way is never run again from the top, unless it is
  marked idempotent (connecting is the only such call).
- A call stuck inside Word past its budget is abandoned with its thread, and
  a fresh worker takes over the queue behind it (a "wedge").

The platform halves (backend/com.py, backend/mac.py) fill in how a thread is
set up, how to tell that Word is busy, and how to connect.
"""

from __future__ import annotations

import contextvars
import logging
import queue
import threading
import time
from concurrent.futures import CancelledError, Future
from concurrent.futures import TimeoutError as FutureTimeout
from dataclasses import dataclass, field
from typing import Any, Callable

logger = logging.getLogger(__name__)

# The document a tool call names with its `document` argument. server.py sets
# it around each call; execute() reads it and hands it to the worker thread,
# where Word.document() resolves it.
call_document: contextvars.ContextVar[str | None] = contextvars.ContextVar(
    "call_document", default=None)

# The futures a tool call has queued, so that a cancelled call (the client went
# away) drops what has not started yet. utils.offload sets it.
pending_calls: contextvars.ContextVar[list | None] = contextvars.ContextVar(
    "pending_calls", default=None)

QUEUE_WAIT = 60.0


class WordBusy(RuntimeError):
    """Word refused a call because a dialog or menu is open."""


@dataclass
class _Job:
    func: Callable[..., Any]
    args: tuple
    kwargs: dict
    idempotent: bool
    document: str | None
    undo: str | None
    future: Future = field(default_factory=Future)
    dequeued: threading.Event = field(default_factory=threading.Event)


class Worker:
    """The queue and the thread. Subclasses connect to Word."""

    #: How long one call may take once it has started, in seconds.
    call_budget = 45.0
    #: How long to wait for a busy Word before giving up on a call.
    busy_budget = 15.0
    busy_backoff = (0.5, 1.0, 2.0, 3.0)
    app_name = "Word"

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self._generation = 0
        self._running = False
        self._local = threading.local()
        self.app = None
        self.launched_by_us = False
        # The session target set by word_activate_document, as a FullName.
        self.target_full_name: str | None = None

    # -- lifecycle ------------------------------------------------------

    def start(self) -> None:
        with self._lock:
            if self._running:
                return
            self._running = True
            self._start_thread_locked()

    def stop(self) -> None:
        with self._lock:
            if not self._running:
                return
            self._running = False
            q, thread = self._queue, self._thread
        q.put(None)
        if thread and thread.is_alive():
            thread.join(timeout=5.0)

    def _start_thread_locked(self) -> None:
        generation = self._generation
        q = self._queue
        self._thread = threading.Thread(
            target=self._loop, args=(q, generation), name=f"word-worker-{generation}",
            daemon=True)
        self._thread.start()

    # -- the calling side -------------------------------------------------

    def execute(self, func: Callable[..., Any], *args: Any, idempotent: bool = False,
                undo: str | None = None, **kwargs: Any) -> Any:
        """Run `func(*args, **kwargs)` on the worker and return what it returns.

        `undo` names the undo step a writing call's changes are gathered
        under, so one Undo in Word takes back one tool call.
        """
        if not self._running:
            raise RuntimeError(
                "The Word worker is not running. This is a bug in the server; "
                "restart it.")
        job = _Job(func, args, kwargs, idempotent, call_document.get(), undo)
        with self._lock:
            self._queue.put(job)
            generation = self._generation
        pending = pending_calls.get()
        if pending is not None:
            pending.append(job.future)
        if not job.dequeued.wait(QUEUE_WAIT):
            job.future.cancel()
            raise RuntimeError(
                f"{self.app_name} did not get to this call within {QUEUE_WAIT:.0f} s "
                "because earlier calls were still running. Nothing was changed; "
                "try again.")
        try:
            return job.future.result(timeout=self.busy_budget + self.call_budget)
        except CancelledError:
            raise RuntimeError(
                "This call was dropped before it started because an earlier one "
                "got stuck. Nothing was changed; try again.") from None
        except FutureTimeout:
            if job.future.cancel():
                raise RuntimeError(
                    f"{self.app_name} did not answer in time. Nothing was changed; "
                    "try again.") from None
            self._recover_from_wedge(generation)
            raise RuntimeError(
                f"{self.app_name} stopped answering in the middle of this call "
                f"(a dialog may be open in {self.app_name}). It may have been "
                "partly applied: check the document before trying again.") from None

    # -- the worker side --------------------------------------------------

    @property
    def wanted_document(self) -> str | None:
        """The document the running call named, or None. Worker thread only."""
        return getattr(self._local, "document", None)

    def _loop(self, q: queue.Queue, generation: int) -> None:
        self.thread_setup()
        try:
            if generation > 0:
                self.after_wedge()
            while True:
                job = q.get()
                if job is None:
                    break
                job.dequeued.set()
                if generation != self._generation:
                    job.future.cancel()
                    break
                self._run(job)
        finally:
            if generation == self._generation:
                self.app = None
            self.thread_teardown()

    def _run(self, job: _Job) -> None:
        if job.future.cancelled():
            return
        deadline = time.monotonic() + self.busy_budget
        attempt = 0
        started = False
        while True:
            try:
                self.wait_until_responsive(deadline)
            except WordBusy as exc:
                if not started and not job.future.set_running_or_notify_cancel():
                    return
                job.future.set_exception(exc)
                return
            if not started:
                if not job.future.set_running_or_notify_cancel():
                    return
                started = True
            self._local.document = job.document
            try:
                job.future.set_result(self.run_call(job))
                return
            except BaseException as exc:  # noqa: BLE001 - every outcome reaches the caller
                if not self.is_busy_error(exc):
                    job.future.set_exception(
                        exc if isinstance(exc, Exception) else RuntimeError(repr(exc)))
                    return
                if not job.idempotent:
                    job.future.set_exception(WordBusy(
                        f"{self.app_name} became busy in the middle of this call (a "
                        "dialog or menu is open?). It may have been partly applied: "
                        "check the document before trying again."))
                    return
            finally:
                self._local.document = None
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                job.future.set_exception(WordBusy(self.busy_message()))
                return
            time.sleep(min(self.busy_backoff[min(attempt, len(self.busy_backoff) - 1)],
                           remaining))
            attempt += 1

    def run_call(self, job: _Job) -> Any:
        """Run the job's function. Subclasses wrap it in an undo step."""
        return job.func(*job.args, **job.kwargs)

    def _recover_from_wedge(self, generation: int) -> None:
        with self._lock:
            if not self._running or generation != self._generation:
                return
            logger.warning("Word stopped answering; starting a fresh worker")
            self._generation += 1
            old, self._queue = self._queue, queue.Queue()
            self._start_thread_locked()
        while True:
            try:
                job = old.get_nowait()
            except queue.Empty:
                break
            if job is not None:
                job.future.cancel()
                job.dequeued.set()
        old.put(None)

    # -- for the platform halves -------------------------------------------

    def thread_setup(self) -> None:
        """Run once on each worker thread before its first call."""

    def thread_teardown(self) -> None:
        """Run once on each worker thread after its last call."""

    def after_wedge(self) -> None:
        """Run on a fresh worker that replaced a stuck one."""
        self.app = None

    def wait_until_responsive(self, deadline: float) -> None:
        """Return once Word takes calls, or raise WordBusy at the deadline."""

    def is_busy_error(self, exc: BaseException) -> bool:
        return False

    def busy_message(self) -> str:
        return (f"{self.app_name} is busy: a dialog or menu is open in it. Close it "
                "and try again. Nothing was changed.")
