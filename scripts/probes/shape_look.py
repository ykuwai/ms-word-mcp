"""A page of shapes as Word lays it out, saved as a picture to look at.

    uv run --no-sync python scripts/probes/shape_look.py

Saves .word-test/shape-look.png. Look for: a pale rectangle behind the text
with the words readable over it; a dark rectangle hiding the text under it;
text going round the square-wrapped oval; true vertical text (upright kana,
reading top to bottom) in the box on the right; and two arrows, the second
drawn backwards, from right to left with its head at the left.
"""

import json
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.tools import _common, documents, screen, shapes, text  # noqa: E402


def run(impl, *args):
    out = json.loads(_common.run(impl, *args))
    if "error" in out:
        print("ERROR", impl.__name__, out["error"])
    return out


FOLDER = Path(__file__).resolve().parents[2] / ".word-test"
FOLDER.mkdir(exist_ok=True)
PAGE = {"horizontal_from": "page", "vertical_from": "page"}
BODY = "これは本文の文章です。The quick brown fox jumps over the lazy dog. " * 7

backend.word.start()
try:
    doc = run(documents._create_document_impl, None, None)
    token = backend.call_document.set(doc["full_name"])
    try:
        run(text._insert_paragraphs_impl, [{"text": BODY, "style": "Normal"}] * 4, None, None)
        run(shapes._insert_shape_impl, "rectangle", 1, 300.0, 50.0, {
            **PAGE, "left": 90.0, "top": 100.0, "wrap": "behind", "fill": "#FFE699",
            "line_color": "none"})
        run(shapes._insert_shape_impl, "rectangle", 2, 200.0, 30.0, {
            **PAGE, "left": 90.0, "top": 230.0, "wrap": "front", "fill": "#203864"})
        run(shapes._insert_shape_impl, "oval", 3, 100.0, 70.0, {
            "left": "center", "top": 10.0, "fill": "#C5E0B4", "text": "楕円"})
        run(shapes._insert_text_box_impl, 4, 50.0, 180.0, {
            **PAGE, "left": 510.0, "top": 420.0, "wrap": "front", "orientation": "vertical",
            "text": "縦書きのテキストABC"})
        run(shapes._insert_line_impl, 4, 90.0, 660.0, 400.0, 660.0, {
            **PAGE, "end_arrow": "triangle", "line_weight": 3.0, "line_color": "#C00000"})
        run(shapes._insert_line_impl, 4, 400.0, 720.0, 90.0, 690.0, {
            **PAGE, "end_arrow": "triangle", "line_weight": 3.0, "line_color": "#0070C0"})
        picture = run(screen._get_page_image_impl, 1, None)
        if "path" in picture:
            shutil.move(picture["path"], FOLDER / "shape-look.png")
            print("Saved", FOLDER / "shape-look.png")
        print(json.dumps(run(shapes._list_shapes_impl, 1, 50), ensure_ascii=False)[:1500])
    finally:
        run(documents._close_document_impl, False)
        backend.call_document.reset(token)
finally:
    backend.word.stop()
