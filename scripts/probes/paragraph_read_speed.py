"""How long word_get_paragraphs takes at its largest, on either platform.

    uv run python scripts/probes/paragraph_read_speed.py

Makes a scratch document of 1,000 paragraphs, reads
200 paragraphs from the top and from the end, with and without `detail`, and
closes the document unsaved. A call has 45 s on Windows and 69 s on the Mac
before the worker waits on it as busy.
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
            for batch in range(5):
                run(text._insert_paragraphs_impl,
                    [{"text": f"{batch}-{i}: a paragraph of ordinary length to read back.",
                      "style": "Heading 2" if i % 25 == 0 else None}
                     for i in range(200)], None, None)
            for start in (1, 801):
                for detail in (False, True):
                    t = time.perf_counter()
                    out = run(text._get_paragraphs_impl, start, start + 199, detail)
                    print(f"start={start} detail={detail}: "
                          f"{time.perf_counter() - t:.1f} s, read {len(out['paragraphs'])}, "
                          f"more={out.get('more')}")
        finally:
            run(documents._close_document_impl, False)
    finally:
        call_document.reset(token)
finally:
    backend.word.stop()
