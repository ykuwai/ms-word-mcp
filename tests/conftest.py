"""Keeps every test off the real Word.

The tool implementations reach Word through ``backend.word``; the `word`
fixture puts a FakeWord there, which runs each call inline against a
FakeDocument. Nothing here starts Word or sends it an Apple Event.
"""

import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

# Full mode, so the tests can call tools by name; test_catalog checks the default.
os.environ["WORD_TOOL_MODE"] = "full"
# The scrolls are recorded by the fake window; keep them on.
os.environ.setdefault("WORD_FOLLOW_EDITS", "on")

from fakeword import FakeDocument, FakeWord  # noqa: E402


@pytest.fixture
def word(monkeypatch):
    from ms_word_mcp import backend

    fake = FakeWord()
    monkeypatch.setattr(backend, "word", fake)
    # The COM implementations run on any platform against the fake; the
    # macOS refusal in tools._common is tested on its own.
    monkeypatch.setattr(backend, "IS_MACOS", False)
    return fake


@pytest.fixture
def make_doc(word):
    def make(*paragraphs, **kwargs):
        word.doc = FakeDocument(paragraphs or ("",), **kwargs)
        return word.doc

    return make
