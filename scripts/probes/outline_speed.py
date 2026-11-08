"""How long word_get_outline takes on a long document, on either platform.

    uv run python scripts/probes/outline_speed.py

Makes a scratch document of 1,000 paragraphs (a Heading 1 every 50, a
Heading 2 every 25, a numbered equation every 100), reads the outline from
Word's heading list and from every paragraph, checks both give the same
headings, and closes the document unsaved.
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.backend import call_document  # noqa: E402
from ms_word_mcp.tools import _common, documents, equations, text  # noqa: E402


def run(impl, *args):
    out = json.loads(_common.run(impl, *args))
    if "error" in out:
        raise RuntimeError(f"{impl.__name__}: {out['error']}")
    return out


backend.word.start()
try:
    created = run(documents._create_document_impl, None, None)
    token = call_document.set(created["full_name"])
    try:
        paragraphs = []
        for i in range(1, 1001):
            if i % 50 == 1:
                paragraphs.append({"text": f"Chapter {i // 50 + 1}", "style": "Heading 1"})
            elif i % 25 == 1:
                paragraphs.append({"text": f"Section {i}", "style": "Heading 2"})
            else:
                paragraphs.append({"text": f"Body paragraph {i} with some words in it.",
                                   "style": None})
        t = time.perf_counter()
        run(text._insert_paragraphs_impl, paragraphs, None, None)
        print(f"insert 1,000 paragraphs: {time.perf_counter() - t:.1f} s")
        for at in range(90, 1000, 100):
            run(equations._insert_equation_impl, "E=mc^2", "latex", True, "Equation", at, None)
        timings = {}
        for label, every in (("from the heading list", False), ("every paragraph", True)):
            t = time.perf_counter()
            out = run(text._get_outline_impl, every)
            timings[label] = (time.perf_counter() - t, out)
            print(f"outline {label}: {timings[label][0]:.2f} s, {len(out['headings'])} "
                  f"headings of {out['paragraphs']} paragraphs")
        same = (timings["from the heading list"][1]["headings"]
                == timings["every paragraph"][1]["headings"])
        print("same headings:", same)
        run(documents._close_document_impl, False)
    finally:
        call_document.reset(token)
finally:
    backend.word.stop()
