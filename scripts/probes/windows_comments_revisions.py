"""Comments and tracked changes on Word for Windows, over COM.

    uv run --no-sync python scripts/probes/windows_comments_revisions.py comments
    uv run --no-sync python scripts/probes/windows_comments_revisions.py tracking
    uv run --no-sync python scripts/probes/windows_comments_revisions.py tools

A scratch document, closed unsaved. Nothing runs while any other document is
open: creating a comment ended the Word process on the Mac, and a crash would
take another document's unsaved work with it.

What it answers: what Range.Text holds once part of a paragraph is a tracked
deletion, whether taking the deletions' spans out gives the text Word shows,
how accepting and rejecting renumber the revisions, and what the existing
tools answer with tracking on.
"""

import json
import sys
from pathlib import Path

import pythoncom
import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
sys.stdout.reconfigure(errors="backslashreplace")


def attempt(label, f):
    try:
        out = f()
        print(f"{label} -> {out!r}")
        return out
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {str(e)[:300]}")
        return None


def word():
    try:
        return win32com.client.GetActiveObject("Word.Application")
    except pythoncom.com_error:
        app = win32com.client.Dispatch("Word.Application")
        app.Visible = True
        return app


def scratch(w):
    d = w.Documents.Add()
    d.Content.Text = "Alpha beta gamma\rSecond 段落 here\rThird paragraph"
    print("paragraphs", paragraphs(d))
    return d


def paragraphs(d):
    return [(i, p.Range.Text, p.Range.Start, p.Range.End)
            for i, p in enumerate(d.Paragraphs, 1)]


def revisions(d):
    return [(r.Index, r.Type, bool(r.Author), r.Range.Text, r.Range.Start, r.Range.End,
             str(r.Date), r.FormatDescription) for r in d.Revisions]


def comments(d):
    out = []
    for c in d.Comments:
        out.append({
            "index": c.Index, "author": bool(c.Author), "initial": bool(c.Initial),
            "text": c.Range.Text, "scope": c.Scope.Text,
            "at": (c.Scope.Start, c.Scope.End), "mark": c.Reference.Start,
            "date": str(c.Date), "done": c.Done, "replies": c.Replies.Count,
            "ancestor": c.Ancestor.Index if c.Ancestor is not None else None})
    return out


def visible(d, start, end):
    """The text between start and end without the tracked deletions in it."""
    pieces, at = [], start
    for r in d.Range(start, end).Revisions:
        if r.Type == 2:  # wdRevisionDelete
            pieces.append(d.Range(at, max(at, r.Range.Start)).Text or "")
            at = r.Range.End
    pieces.append(d.Range(at, end).Text or "")
    return "".join(pieces)


def run_comments(w):
    d = scratch(w)
    try:
        before = d.Paragraphs(1).Range.Text
        c1 = attempt("add over a word", lambda: d.Comments.Add(d.Range(6, 10), "書いた 𠮷 text"))
        print("p1 was", repr(before), "now", repr(d.Paragraphs(1).Range.Text),
              d.Paragraphs(1).Range.End)
        attempt("add at a point", lambda: d.Comments.Add(d.Range(11, 11), "at a point"))
        attempt("list", lambda: comments(d))
        attempt("reply", lambda: c1.Replies.Add(c1.Range, "a reply"))
        attempt("list", lambda: comments(d))
        attempt("set the text", lambda: setattr(c1.Range, "Text", "changed"))
        attempt("done", lambda: setattr(c1, "Done", True))
        attempt("list", lambda: comments(d))
        attempt("paragraphs", lambda: paragraphs(d))
        attempt("delete the point comment", lambda: d.Comments(d.Comments.Count).Delete())
        attempt("list", lambda: comments(d))
        attempt("delete a comment with a reply",
                lambda: d.Comments(1).DeleteRecursively())
        attempt("list", lambda: comments(d))
    finally:
        d.Close(0)


def run_tracking(w):
    d = scratch(w)
    try:
        print("track", d.TrackRevisions)
        d.TrackRevisions = True
        d.Range(6, 11).Delete()
        print("p1", repr(d.Paragraphs(1).Range.Text), d.Paragraphs(1).Range.End,
              "paras", d.Paragraphs.Count)
        attempt("visible p1", lambda: visible(d, d.Paragraphs(1).Range.Start,
                                              d.Paragraphs(1).Range.End))
        attempt("revisions", lambda: revisions(d))
        p2 = d.Paragraphs(2).Range
        d.Range(p2.End - 1, p2.End - 1).Text = "X"
        p3 = d.Paragraphs(3).Range
        d.Range(p3.Start, p3.Start + 5).Text = "NEW"
        print("p3", repr(d.Paragraphs(3).Range.Text))
        attempt("visible p3", lambda: visible(d, d.Paragraphs(3).Range.Start,
                                              d.Paragraphs(3).Range.End))
        d.Range(0, 5).Font.Bold = True
        d.Paragraphs(1).Format.Alignment = 1
        attempt("revisions", lambda: revisions(d))

        print("\n== the view")
        view = d.ActiveWindow.View
        attempt("ShowRevisionsAndComments", lambda: view.ShowRevisionsAndComments)
        attempt("RevisionsView", lambda: view.RevisionsView)
        attempt("RevisionsFilter.Markup", lambda: view.RevisionsFilter.Markup)
        view.ShowRevisionsAndComments = False
        print("markup off: p1", repr(d.Paragraphs(1).Range.Text), d.Paragraphs(1).Range.End,
              "p3", repr(d.Paragraphs(3).Range.Text))
        view.ShowRevisionsAndComments = True
        print("markup on: p1", repr(d.Paragraphs(1).Range.Text))

        print("\n== a whole paragraph deleted")
        before = d.Paragraphs.Count
        p2 = d.Paragraphs(2).Range
        d.Range(p2.Start, p2.End).Delete()
        print("count", before, "->", d.Paragraphs.Count)
        attempt("paragraphs", lambda: paragraphs(d))
        attempt("revisions", lambda: revisions(d))

        print("\n== accept and reject")
        n = d.Revisions.Count
        attempt("accept 1", lambda: d.Revisions(1).Accept())
        print("count", n, "->", d.Revisions.Count)
        attempt("revisions", lambda: revisions(d))
        attempt("reject 1", lambda: d.Revisions(1).Reject())
        attempt("revisions", lambda: revisions(d))
        attempt("range count", lambda: d.Paragraphs(1).Range.Revisions.Count)
        attempt("range accept all", lambda: d.Paragraphs(1).Range.Revisions.AcceptAll())
        attempt("revisions", lambda: revisions(d))
        attempt("reject all", lambda: d.Revisions.RejectAll())
        attempt("revisions", lambda: revisions(d))
        attempt("paragraphs", lambda: paragraphs(d))
    finally:
        d.TrackRevisions = False
        d.Close(0)


def run_tools(w):
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
        d = w.Documents(made["name"])
        d.TrackRevisions = True
        run("rewrite paragraph 1", text._set_paragraph_text_impl, 1, "Rewritten", undo="Set")
        run("read 1-2", text._get_paragraphs_impl, 1, 2, False)
        run("delete paragraph 2", text._delete_paragraphs_impl, 2, None, undo="Delete")
        run("read all", text._get_paragraphs_impl, 1, None, False)
        run("insert after 1", text._insert_paragraphs_impl, [{"text": "New one"}], 1, None,
            undo="Insert")
        run("read all", text._get_paragraphs_impl, 1, None, False)
        run("find and replace", text._find_replace_impl, "cat", "dog", False, True, False,
            undo="Replace")
        run("read all", text._get_paragraphs_impl, 1, None, False)
        run("outline", text._get_outline_impl)
        print("revisions", d.Revisions.Count)
        d.TrackRevisions = False
    finally:
        run("close", documents._close_document_impl, False)
        call_document.reset(token)
        backend.word.stop()


SECTIONS = {"comments": run_comments, "tracking": run_tracking, "tools": run_tools}

if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in SECTIONS:
        raise SystemExit(f"Say which: {', '.join(SECTIONS)}")
    app = word()
    if app.Documents.Count:
        raise SystemExit("Another document is open in Word. Nothing was run.")
    SECTIONS[sys.argv[1]](app)
