"""How long word_find_replace takes with thousands of matches, on either platform.

    uv run python scripts/probes/find_replace_speed.py

Makes a scratch document of 1,000 paragraphs holding three matches each,
counts the 3,000 matches, replaces them, and closes the document unsaved. A
call stops itself at its time limit and says how far it got.
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.backend import call_document  # noqa: E402
from ms_word_mcp.tools import _common, documents, text  # noqa: E402


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
        try:
            for _ in range(5):
                run(text._insert_paragraphs_impl,
                    [{"text": f"{i}: zeta and zeta, then zeta.", "style": None}
                     for i in range(200)], None, None)
            for label, new in (("count", None), ("replace", "eta")):
                t = time.perf_counter()
                out = run(text._find_replace_impl, "zeta", new, True, False, False)
                print(f"{label}: {time.perf_counter() - t:.1f} s, matches={out['matches']} "
                      f"replaced={out.get('replaced')} complete={out.get('complete', True)} "
                      f"warnings={out.get('warnings')}")
        finally:
            run(documents._close_document_impl, False)
    finally:
        call_document.reset(token)
finally:
    backend.word.stop()
