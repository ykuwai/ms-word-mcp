"""The worker's promises: one call at a time, nothing applied after the caller
gave up, nothing run twice after a partial failure, and recovery from a call
stuck inside Word."""

import contextvars
import threading
import time

import pytest

from ms_word_mcp.backend import worker as worker_mod
from ms_word_mcp.backend.worker import Worker, WordBusy, call_document


class BusyError(Exception):
    pass


class Probe(Worker):
    busy_budget = 0.5
    busy_backoff = (0.01,)

    def is_busy_error(self, exc):
        return isinstance(exc, BusyError)


@pytest.fixture
def w():
    worker = Probe()
    worker.start()
    yield worker
    worker.stop()


def test_calls_run_on_one_thread_in_order(w):
    seen = []
    threads = set()

    def note(i):
        seen.append(i)
        threads.add(threading.current_thread().name)
        return i

    assert [w.execute(note, i) for i in range(5)] == [0, 1, 2, 3, 4]
    assert seen == [0, 1, 2, 3, 4] and len(threads) == 1


def test_errors_reach_the_caller(w):
    def boom():
        raise ValueError("no such paragraph")

    with pytest.raises(ValueError, match="no such paragraph"):
        w.execute(boom)


def test_the_called_document_is_seen_on_the_worker(w):
    token = call_document.set("report.docx")
    try:
        assert w.execute(lambda: w.wanted_document) == "report.docx"
    finally:
        call_document.reset(token)
    assert w.execute(lambda: w.wanted_document) is None


def test_a_call_busy_part_way_is_not_run_again(w):
    runs = []

    def write():
        runs.append(1)
        raise BusyError()

    with pytest.raises(WordBusy, match="partly applied"):
        w.execute(write)
    assert len(runs) == 1


def test_an_idempotent_call_is_retried_while_busy(w):
    runs = []

    def connect():
        runs.append(1)
        if len(runs) < 3:
            raise BusyError()
        return "connected"

    assert w.execute(connect, idempotent=True) == "connected"
    assert len(runs) == 3


def test_a_cancelled_call_is_never_applied(w):
    release = threading.Event()
    applied = []
    queued = []
    token = worker_mod.pending_calls.set(queued)
    try:
        blocker = threading.Thread(target=w.execute, args=(release.wait,))
        blocker.start()
        time.sleep(0.05)
        # A tool call's thread sees the caller's context, as anyio gives it.
        context = contextvars.copy_context()
        late = threading.Thread(target=lambda: context.run(
            _swallow, w.execute, applied.append, 1))
        late.start()
        time.sleep(0.05)
        for future in queued:
            future.cancel()
        release.set()
        blocker.join(2)
        late.join(2)
    finally:
        worker_mod.pending_calls.reset(token)
    assert applied == []


def _swallow(func, *args):
    try:
        func(*args)
    except Exception:  # noqa: BLE001
        pass


def test_a_stuck_call_is_abandoned_and_the_next_one_runs(monkeypatch):
    worker = Probe()
    worker.call_budget = 0.2
    worker.start()
    try:
        stuck = threading.Event()
        with pytest.raises(RuntimeError, match="stopped answering"):
            worker.execute(stuck.wait, 5)
        assert worker.execute(lambda: "fresh") == "fresh"
        stuck.set()
    finally:
        worker.stop()
