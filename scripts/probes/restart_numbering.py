"""Whether restarting a list's numbering gives the same numbers every time.

    uv run python scripts/probes/restart_numbering.py

Eight times, in a fresh scratch document with numbered headings and a list of
four after them: restart at the third item, then read the four numbers at once
and again two seconds later. The right answer is 1. 2. 1. 2. every time. Word
for Mac's scripting did not give it (docs/PROBES.md), so there the tool refuses
and this prints the refusal.
"""

import json
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.backend import call_document  # noqa: E402
from ms_word_mcp.tools import _common, documents, lists, text  # noqa: E402

RUNS = 8
HEADS = [("Intro", "Heading 1"), ("Body", None), ("Scope", "Heading 2"), ("Aim", "Heading 2"),
         ("Method", "Heading 1"), ("Steps", "Heading 2"), ("Later", "Heading 2")]
HEADING_LEVELS = [
    {"level": 1, "number_style": "arabic", "number_format": "%1", "linked_style": "Heading 1"},
    {"level": 2, "number_style": "arabic", "number_format": "%1.%2", "linked_style": "Heading 2"}]
LIST_LEVEL = [{"level": 1, "number_style": "arabic", "number_format": "%1."}]


def run(impl, *args):
    out = json.loads(_common.run(impl, *args))
    if "error" in out:
        raise RuntimeError(f"{impl.__name__}: {out['error']}")
    return out


def numbers():
    return [run(lists._get_list_impl, i)["paragraph"].get("list_string") for i in range(8, 12)]


def once(number: int) -> bool:
    created = run(documents._create_document_impl, None, None)
    token = call_document.set(created["full_name"])
    try:
        try:
            run(text._insert_paragraphs_impl, [{"text": t, "style": s} for t, s in HEADS],
                None, None)
            run(lists._apply_list_impl, 1, 1, HEADING_LEVELS, None, False)
            run(text._insert_paragraphs_impl,
                [{"text": t, "style": "Normal"} for t in "abcd"], None, None)
            run(lists._apply_list_impl, 8, 11, LIST_LEVEL, None, False)
            before = numbers()
            answer = json.loads(_common.run(lists._restart_numbering_impl, 10))
            if "error" in answer:
                print(f"{number}: refused: {answer['error']}")
                return False
            first = numbers()
            time.sleep(2)
            later = numbers()
            right = first == later == ["1.", "2.", "1.", "2."]
            print(f"{number}: before={before} at once={first} later={later} "
                  f"{'right' if right else 'WRONG'}")
            return right
        finally:
            run(documents._close_document_impl, False)
    finally:
        call_document.reset(token)


backend.word.start()
try:
    right = sum(once(n) for n in range(1, RUNS + 1))
    print(f"{right} of {RUNS} runs gave the right numbers.")
finally:
    backend.word.stop()
