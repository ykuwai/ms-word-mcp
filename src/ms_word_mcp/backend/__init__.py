"""The Word this server drives, on the platform it runs on.

Every tool module says ``from ms_word_mcp import backend`` and reaches Word
through ``backend.word`` at call time, so the tests can put a fake in its
place.
"""

from __future__ import annotations

import sys

from .worker import Worker, call_document, pending_calls  # noqa: F401

IS_WINDOWS = sys.platform == "win32"
IS_MACOS = sys.platform == "darwin"
PLATFORM_NAME = "Windows" if IS_WINDOWS else ("macOS" if IS_MACOS else sys.platform)

if IS_MACOS:
    from .mac import MacWord

    word: Worker = MacWord()
else:
    # Windows, and anywhere else the tests run: the COM half imports without
    # pywin32 and only fails when it is asked to connect.
    from .com import ComWord

    word = ComWord()
