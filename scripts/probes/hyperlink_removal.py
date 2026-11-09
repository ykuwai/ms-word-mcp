"""Whether taking a hyperlink off also takes off its look, on Windows.

    uv run python scripts/probes/hyperlink_removal.py

A scratch document, closed unsaved. Prints the character style of the linked
text before and after Hyperlink.Delete. On the Mac the same was measured with
the dictionary: the text goes from Hyperlink back to Normal.
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.tools import _common  # noqa: E402


def measure():
    doc = backend.word.get_app().Documents.Add()
    try:
        doc.Content.Text = "Alpha beta gamma."
        link = doc.Hyperlinks.Add(doc.Range(6, 10), "https://example.com", "")
        result = link.Range.Fields.Item(1).Result
        before = (str(result.CharacterStyle), result.Font.Underline, result.Font.Color)
        link.Delete()
        text = doc.Range(6, 10)
        after = (str(text.CharacterStyle), text.Font.Underline, text.Font.Color)
        return {"text": text.Text, "before": before, "after": after,
                "whole": doc.Content.Text, "links": doc.Hyperlinks.Count}
    finally:
        doc.Close(0)


if backend.IS_MACOS:
    print("Windows only.")
else:
    print(json.dumps(json.loads(_common.run(measure)), ensure_ascii=False))
