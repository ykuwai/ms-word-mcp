"""Word on macOS, over Apple Events.

Not yet verified against a live Word for Mac. What is here follows what
ppt-mcp measured for PowerPoint for Mac, which shares Office's scripting
layer: the dictionary must be loaded from the sdef, references Word hands
back are not trusted (count, then address by index), and one call at a time.
Tools whose Mac half is not written yet refuse on macOS and say so; see
tools/_platform.py.
"""

from __future__ import annotations

import logging
import subprocess

from .documents import OpenDocument, pick_document
from .worker import Worker

logger = logging.getLogger(__name__)

BUNDLE_ID = "com.microsoft.Word"

# Apple Event errors worth one more try for a call that may run twice: a
# timeout (-1712), a connection that went away (-609), an app not running (-600).
RETRYABLE = frozenset({-1712, -609, -600})

# A walk over Word's documents never goes further than this, whatever Word answers.
MAX_DOCUMENTS = 200


def _decode_wide_integers() -> None:
    """Let appscript read the 8-byte integers Word for Mac sends as 32-bit ones.

    Measured on Word 16.114: the answers to `count each paragraph` and
    `compute statistics` are typed as 32-bit integers but carry 8 bytes, and
    appscript fails on them with "unpack requires a buffer of 4 bytes".
    """
    import struct

    from aem import aemcodecs

    original = aemcodecs.Codecs.unpacksint32
    if getattr(original, "_word_wide", False):
        return

    def unpacksint32(self, desc):
        if len(desc.data) == 8:
            return struct.unpack("q", desc.data)[0]
        return original(self, desc)

    unpacksint32._word_wide = True
    aemcodecs.Codecs.unpacksint32 = unpacksint32


def _ae_number(exc: BaseException) -> int | None:
    number = getattr(exc, "errornumber", None)
    if number is None:
        number = getattr(exc, "number", None)
    return number if isinstance(number, int) else None


class MacWord(Worker):
    """The worker, connected to Word for Mac over Apple Events."""

    call_budget = 69.0

    def is_busy_error(self, exc: BaseException) -> bool:
        return _ae_number(exc) in RETRYABLE

    def _word_running(self) -> bool:
        try:
            return subprocess.run(["pgrep", "-x", "Microsoft Word"], capture_output=True,
                                  timeout=5).returncode == 0
        except (OSError, subprocess.SubprocessError):
            return False

    def connect(self, allow_launch: bool = False):
        import appscript

        _decode_wide_integers()
        running = self._word_running()
        if not running and not allow_launch:
            raise ConnectionError(
                "Word is not running. Start it, or call word_create_document or "
                "word_open_document, which start it.")
        # terms="sdef": Office's applications answer the terminology request
        # with a few terms only; the full dictionary comes from the sdef.
        app = appscript.app(id=BUNDLE_ID, terms="sdef")
        if not running:
            app.launch()
            self.launched_by_us = True
        else:
            self.launched_by_us = False
        self.app = app
        return app

    def get_app(self, allow_launch: bool = False):
        if self.app is not None and self._word_running():
            return self.app
        self.app = None
        return self.connect(allow_launch)

    def open_documents(self) -> list[OpenDocument]:
        """The open documents, found by asking for each index in turn.

        Measured on Word 16.114: `documents.get()` answers missing value when
        none is open, and asking for a document past the last one answers
        missing value for its name rather than an error, so the walk stops at
        the first missing name (and at MAX_DOCUMENTS whatever happens).
        """
        from appscript import k

        app = self.get_app()
        out = []
        for i in range(1, MAX_DOCUMENTS + 1):
            doc = app.documents[i]
            name = doc.name.get()
            if name == k.missing_value or not isinstance(name, str):
                break
            out.append(OpenDocument(i, name, doc.full_name.get()))
        return out

    def document(self):
        app = self.get_app()
        docs = self.open_documents()
        wanted = self.wanted_document or self.target_full_name
        if wanted:
            return app.documents[pick_document(docs, wanted).index]
        if not docs:
            raise ValueError(
                "No document is open in Word. Open one with word_open_document or "
                "create one with word_create_document.")
        return app.documents[1]
