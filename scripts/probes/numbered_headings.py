"""Whether word_apply_list with linked styles numbers a document's headings.

    uv run python scripts/probes/numbered_headings.py

Makes a scratch document with headings and body text, applies a two-level list
to the first heading only with its levels linked to Heading 1 and Heading 2,
adds one more heading afterwards, prints what each paragraph reads back, and
closes the document unsaved.
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.backend import call_document  # noqa: E402
from ms_word_mcp.tools import _common, documents, lists, text  # noqa: E402


def run(impl, *args):
    out = json.loads(_common.run(impl, *args))
    if "error" in out:
        raise RuntimeError(f"{impl.__name__}: {out['error']}")
    return out


H1, H2, BODY = "Heading 1", "Heading 2", None
OUTLINE = [("Introduction", H1), ("Body one", BODY), ("Background", H2), ("Body two", BODY),
           ("Aim", H2), ("Method", H1), ("Body three", BODY), ("Steps", H2)]

backend.word.start()
try:
    created = run(documents._create_document_impl, None, None)
    token = call_document.set(created["full_name"])
    try:
        try:
            run(text._insert_paragraphs_impl,
                [{"text": t, "style": s} for t, s in OUTLINE], None, None)
            levels = [lists.LevelSpec(level=1, number_style="arabic", number_format="%1",
                                      linked_style=H1).model_dump(exclude_none=True),
                      lists.LevelSpec(level=2, number_style="arabic", number_format="%1.%2",
                                      linked_style=H2).model_dump(exclude_none=True)]
            applied = run(lists._apply_list_impl, 1, 1, levels, None, False)
            print("applied:", json.dumps(applied, ensure_ascii=False))
            run(text._insert_paragraphs_impl,
                [{"text": "Added afterwards", "style": H2}], None, None)
            for i in range(1, len(OUTLINE) + 2):
                print(i, json.dumps(run(lists._get_list_impl, i), ensure_ascii=False)[:200])
        finally:
            run(documents._close_document_impl, False)
    finally:
        call_document.reset(token)
finally:
    backend.word.stop()
