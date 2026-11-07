"""Word on Windows, over COM.

Everything here runs on the worker thread, which is the COM apartment Word's
objects belong to. The module imports on any platform so the tests can load
it; only connecting needs pywin32.
"""

from __future__ import annotations

import logging
import time

from .documents import OpenDocument, pick_document
from .worker import Worker, WordBusy

logger = logging.getLogger(__name__)

try:  # pragma: no cover - Windows only
    import pythoncom
    import pywintypes
    import win32com.client

    ComError = pywintypes.com_error
except ImportError:  # the tests, and macOS
    pythoncom = None
    win32com = None

    class ComError(Exception):  # type: ignore[no-redef]
        """Stands in for pywintypes.com_error where pywin32 is missing."""

        def __init__(self, hresult=0, strerror="", excepinfo=None, argerror=None):
            super().__init__(hresult, strerror, excepinfo, argerror)
            self.hresult = hresult
            self.strerror = strerror
            self.excepinfo = excepinfo

# RPC_E_CALL_REJECTED and RPC_E_SERVERCALL_RETRYLATER: Word has a modal dialog
# or a menu open and refuses calls until it closes.
BUSY_HRESULTS = frozenset({-2147418111, -2147417846})


def describe_com_error(exc: BaseException) -> str:
    """The most useful sentence a com_error carries."""
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        return str(excepinfo[2]).strip()
    strerror = getattr(exc, "strerror", None)
    return str(strerror or exc)


class ComWord(Worker):
    """The worker, connected to Word over COM."""

    def thread_setup(self) -> None:
        if pythoncom is not None:
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)

    def thread_teardown(self) -> None:
        if pythoncom is not None:
            pythoncom.CoUninitialize()

    def is_busy_error(self, exc: BaseException) -> bool:
        return isinstance(exc, ComError) and getattr(exc, "hresult", None) in BUSY_HRESULTS

    def wait_until_responsive(self, deadline: float) -> None:
        if self.app is None:
            return
        attempt = 0
        while True:
            try:
                _ = self.app.Name
                return
            except ComError as exc:
                if not self.is_busy_error(exc):
                    return  # stale or gone: connecting deals with it
            except AttributeError:
                return
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise WordBusy(self.busy_message())
            time.sleep(min(self.busy_backoff[min(attempt, len(self.busy_backoff) - 1)],
                           remaining))
            attempt += 1

    def run_call(self, job):
        # One undo step per writing call. Word gathers every change made
        # through COM between StartCustomRecord and EndCustomRecord into one
        # entry named after the tool, so a single Undo takes the call back.
        # A record the user's own macro already started is left alone.
        record = None
        if job.undo:
            try:
                candidate = self.get_app().UndoRecord
                if not candidate.IsRecordingCustomRecord:
                    candidate.StartCustomRecord(job.undo[:64])
                    record = candidate
            except Exception:  # noqa: BLE001 - the call still runs, ungrouped
                logger.debug("Could not start an undo record", exc_info=True)
        try:
            return job.func(*job.args, **job.kwargs)
        finally:
            if record is not None:
                try:
                    record.EndCustomRecord()
                except Exception:  # noqa: BLE001
                    logger.debug("Could not end the undo record", exc_info=True)

    # -- connecting -------------------------------------------------------

    def connect(self, allow_launch: bool = False):
        """Word's Application object, attaching to a running Word first."""
        if win32com is None:
            raise RuntimeError("pywin32 is not installed, so Word cannot be reached.")
        try:
            app = win32com.client.GetActiveObject("Word.Application")
            self.launched_by_us = False
        except ComError:
            if not allow_launch:
                raise ConnectionError(
                    "Word is not running. Start it, or call word_create_document or "
                    "word_open_document, which start it.") from None
            app = win32com.client.Dispatch("Word.Application")
            self.launched_by_us = True
            app.Visible = True
        self.app = app
        return app

    def get_app(self, allow_launch: bool = False):
        """The connected Application, reconnecting if the old one went away."""
        if self.app is not None:
            try:
                _ = self.app.Name
                return self.app
            except ComError as exc:
                if self.is_busy_error(exc):
                    raise
                self.app = None
            except AttributeError:
                self.app = None
        return self.connect(allow_launch)

    # -- documents --------------------------------------------------------

    def open_documents(self) -> list[OpenDocument]:
        app = self.get_app()
        docs = app.Documents
        return [
            OpenDocument(i, docs.Item(i).Name, docs.Item(i).FullName)
            for i in range(1, docs.Count + 1)
        ]

    def document(self):
        """The Document this call works on. Worker thread only.

        Never activates a window: reading or writing a document through its
        ranges leaves the user's own window, and the cursor in it, alone.
        """
        app = self.get_app()
        wanted = self.wanted_document
        if wanted:
            found = pick_document(self.open_documents(), wanted)
            return app.Documents.Item(found.index)
        if self.target_full_name:
            try:
                found = pick_document(self.open_documents(), self.target_full_name)
                return app.Documents.Item(found.index)
            except ValueError:
                self.target_full_name = None  # it was closed
        if app.Documents.Count == 0:
            raise ValueError(
                "No document is open in Word. Open one with word_open_document or "
                "create one with word_create_document.")
        return app.ActiveDocument

    def window_of(self, doc):
        """The document's first window, or None if it has none."""
        try:
            windows = doc.Windows
            return windows.Item(1) if windows.Count else None
        except ComError:
            return None
