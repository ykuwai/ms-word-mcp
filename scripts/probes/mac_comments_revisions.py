"""Comments and tracked changes through Word for Mac's scripting.

    uv run python scripts/probes/mac_comments_revisions.py tracking
    uv run python scripts/probes/mac_comments_revisions.py tools
    uv run python scripts/probes/mac_comments_revisions.py comment-at-document

A scratch document, closed unsaved. Each way of doing it is tried and its
answer printed, since the dictionary names the classes but not how to make one
(a comment's text is read-only there), and does not say what the text of a
paragraph holds once part of it is marked as deleted.

`make new Word comment at <text range>` ended the Word process (16.x, measured
2026-11-09), so that form is not tried again, and a section that makes a
comment runs alone. Nothing runs while any other document is open: a crash
would take that document's unsaved work with it.
"""

import json
import subprocess
import sys
import time
from pathlib import Path

import appscript
from appscript import k

from ms_word_mcp.backend.mac import _decode_wide_integers

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")


def word_is_empty() -> bool:
    """Start Word if it is not running; True when no document is open."""
    subprocess.run(["open", "-g", "-b", "com.microsoft.Word"], check=True)
    for _ in range(30):
        try:
            return int(w.count(each=k.document, timeout=10)) == 0
        except Exception:  # noqa: BLE001 - still starting
            time.sleep(1)
    raise SystemExit("Word did not answer. Nothing was run.")


def attempt(label, f):
    try:
        out = f()
        print(f"{label} -> {out!r}")
        return out
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {' '.join(x.strip() for x in str(e).splitlines()[:3])[:200]}")
        return None


def bounds(d, i):
    t = d.paragraphs[i].text_object
    return t.start_of_content.get(), t.end_of_content.get()


def para(d, i):
    return d.paragraphs[i].text_object.content.get()


def comments(d):
    n = int(d.count(each=k.Word_comment))
    out = []
    for i in range(1, n + 1):
        c = d.Word_comments[i]
        out.append({
            "index": c.comment_index.get(), "author": bool(c.author.get()),
            "initials": bool(c.initials.get()), "text": c.comment_text.content.get(),
            "scope": c.scope.content.get(),
            "at": (c.scope.start_of_content.get(), c.scope.end_of_content.get()),
            "mark": c.note_reference.start_of_content.get(), "date": str(c.date_value.get())})
    return out


def revisions(d):
    n = int(d.count(each=k.revision))
    out = []
    for i in range(1, n + 1):
        r = d.revisions[i]
        t = r.text_object
        out.append((r.entry_index.get(), str(r.revision_type.get()), bool(r.author.get()),
                    t.content.get(), t.start_of_content.get(), t.end_of_content.get()))
    return out


def scratch():
    d = w.make(new=k.document)
    d.text_object.content.set("Alpha beta gamma\rSecond 段落 here\rThird paragraph")
    print("paragraphs", [para(d, i) for i in (1, 2, 3)], [bounds(d, i) for i in (1, 2, 3)])
    return d


def tracking():
    d = scratch()
    try:
        print("track", d.track_revisions.get())
        d.track_revisions.set(True)
        print("track", d.track_revisions.get(), "revisions", d.count(each=k.revision))
        d.create_range(start=6, end_=11).content.set("")  # "beta "
        print("after delete: p1", repr(para(d, 1)), bounds(d, 1), "paras",
              d.count(each=k.paragraph))
        attempt("revisions", lambda: revisions(d))
        a, b = bounds(d, 2)
        attempt("insert at end of p2",
                lambda: d.create_range(start=b - 1, end_=b - 1).content.set("X"))
        print("p2", repr(para(d, 2)), bounds(d, 2))
        attempt("revisions", lambda: revisions(d))
        a, b = bounds(d, 3)
        attempt("replace a word in p3",
                lambda: d.create_range(start=a, end_=a + 5).content.set("NEW"))
        print("p3", repr(para(d, 3)), bounds(d, 3))
        attempt("revisions", lambda: revisions(d))
        attempt("format bold", lambda: d.create_range(start=0, end_=5).font_object.bold.set(True))
        attempt("revisions", lambda: revisions(d))
        view = d.active_window.view
        attempt("revisions view", lambda: view.revisions_view.get())
        attempt("show markup", lambda: view.show_revisions_and_comments.get())

        print("\n== the same text with the markup hidden")
        attempt("markup off", lambda: view.show_revisions_and_comments.set(False))
        print("p1", repr(para(d, 1)), bounds(d, 1), "p3", repr(para(d, 3)))
        attempt("markup on", lambda: view.show_revisions_and_comments.set(True))

        print("\n== accept and reject")
        attempt("accept 1", lambda: d.revisions[1].accept())
        attempt("revisions", lambda: revisions(d))
        print("p1", repr(para(d, 1)))
        attempt("reject 1", lambda: d.revisions[1].reject())
        attempt("revisions", lambda: revisions(d))
        print("p2", repr(para(d, 2)))
        attempt("in a range", lambda: d.create_range(start=0, end_=5).count(each=k.revision))
        attempt("accept all in a range",
                lambda: d.create_range(start=0, end_=5).accept_all_revisions())
        attempt("revisions", lambda: revisions(d))
        attempt("reject all", lambda: d.reject_all_revisions())
        attempt("revisions", lambda: revisions(d))
        print("paragraphs", [para(d, i) for i in (1, 2, 3)])
    finally:
        d.track_revisions.set(False)
        d.close(saving=k.no)


def tools():
    """The existing tools, with tracking on: what they read back and report."""
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
    from ms_word_mcp import backend
    from ms_word_mcp.backend import call_document
    from ms_word_mcp.tools import _common, documents, text

    def run(label, impl, *args, undo=None):
        out = json.loads(_common.run(impl, *args, undo=undo))
        print(f"{label}\n   {json.dumps(out, ensure_ascii=False)[:900]}")
        return out

    backend.word.start()
    made = run("create", documents._create_document_impl, None, None)
    token = call_document.set(made["name"])
    try:
        run("insert", text._insert_paragraphs_impl,
            [{"text": "Alpha beta gamma"}, {"text": "Second 段落 here"}, {"text": "Third paragraph"},
             {"text": "Fourth cat and cat"}], None, None, undo="Insert")
        d = w.documents[made["name"]]
        d.track_revisions.set(True)
        run("rewrite paragraph 1", text._set_paragraph_text_impl, 1, "Rewritten", undo="Set")
        run("read 1-2", text._get_paragraphs_impl, 1, 2, False)
        run("delete paragraph 2", text._delete_paragraphs_impl, 2, None, undo="Delete")
        run("read all", text._get_paragraphs_impl, 1, None, False)
        run("insert after 1", text._insert_paragraphs_impl, [{"text": "New one"}], 1, None, undo="Insert")
        run("read all", text._get_paragraphs_impl, 1, None, False)
        run("find and replace", text._find_replace_impl, "cat", "dog", False, True, False,
            undo="Replace")
        run("read all", text._get_paragraphs_impl, 1, None, False)
        run("outline", text._get_outline_impl)
        print("revisions", int(d.count(each=k.revision)))
        d.track_revisions.set(False)
    finally:
        run("close", documents._close_document_impl, False)
        call_document.reset(token)
        backend.word.stop()


def comment_at_document():
    d = scratch()
    try:
        made = attempt("make at the document, the scope as a property", lambda: d.make(
            new=k.Word_comment, at=d,
            with_properties={k.scope: d.create_range(start=11, end_=16)}))
        print("count", d.count(each=k.Word_comment))
        attempt("list", lambda: comments(d))
        print("paragraphs", [para(d, i) for i in (1, 2, 3)], [bounds(d, i) for i in (1, 2, 3)])
        if int(d.count(each=k.Word_comment)):
            c = d.Word_comments[1]
            attempt("set the text", lambda: c.comment_text.content.set("書いた 𠮷 text"))
            attempt("list", lambda: comments(d))
            attempt("all properties", lambda: sorted(str(x) for x in c.properties.get()))
            attempt("delete", lambda: c.delete())
            attempt("list", lambda: comments(d))
        print("made", made)
    finally:
        d.close(saving=k.no)


SECTIONS = {"tracking": tracking, "tools": tools, "comment-at-document": comment_at_document}

if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in SECTIONS:
        raise SystemExit(f"Say which: {', '.join(SECTIONS)}")
    if not word_is_empty():
        raise SystemExit("Another document is open in Word. Nothing was run.")
    SECTIONS[sys.argv[1]]()
