"""Every shape kind the tools name, made and read back, on either platform.

    uv run --no-sync python scripts/probes/shape_kinds.py

One scratch document, closed unsaved. The table in tools/shapes.py gives each
kind one name for both platforms; this says whether Word agrees with it.
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.tools import _common, documents, shapes, text  # noqa: E402


def run(impl, *args):
    return json.loads(_common.run(impl, *args))


backend.word.start()
try:
    doc = run(documents._create_document_impl, None, None)
    token = backend.call_document.set(doc["full_name"])
    try:
        run(text._insert_paragraphs_impl, [{"text": "Scratch.", "style": "Normal"}], None, None)
        kinds = sorted(shapes.KINDS, key=shapes.KINDS.get)
        wrong = []
        for kind in kinds:
            out = run(shapes._insert_shape_impl, kind, 1, 40.0, 40.0, {"wrap": "front"})
            if out.get("kind") != kind:
                wrong.append((kind, out.get("kind"), out.get("error")))
            elif out["name"]:
                run(shapes._delete_shape_impl, out["name"])
        print(f"{len(kinds) - len(wrong)} of {len(kinds)} kinds read back as themselves.")
        for row in wrong:
            print("  ", row)
    finally:
        run(documents._close_document_impl, False)
        backend.call_document.reset(token)
finally:
    backend.word.stop()
