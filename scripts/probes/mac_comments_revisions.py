"""Comments and tracked changes through Word for Mac's scripting.

    uv run python scripts/probes/mac_comments_revisions.py tracking
    uv run python scripts/probes/mac_comments_revisions.py tools
    uv run python scripts/probes/mac_comments_revisions.py ruby
    uv run python scripts/probes/mac_comments_revisions.py revision-reads
    uv run python scripts/probes/mac_comments_revisions.py comment-at-document
    uv run python scripts/probes/mac_comments_revisions.py comment-read
    uv run python scripts/probes/mac_comments_revisions.py comment-edit
    uv run python scripts/probes/mac_comments_revisions.py comment-delete

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
import zipfile
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


SAMPLE = Path(__file__).resolve().parents[2] / ".word-test" / "comment-sample.docx"
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def write_sample() -> None:
    """A .docx that holds two comments, so Word for Mac need not make one."""
    def comment(i, text):
        return (f'<w:comment w:id="{i}" w:author="Reviewer" w:initials="R" '
                f'w:date="2026-11-09T10:00:00Z"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>')

    def anchored(i, text):
        return (f'<w:commentRangeStart w:id="{i}"/><w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
                f'<w:commentRangeEnd w:id="{i}"/><w:r><w:commentReference w:id="{i}"/></w:r>')

    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
            '</Types>'),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>'),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
            '</Relationships>'),
        "word/comments.xml": (
            f'<?xml version="1.0" encoding="UTF-8"?><w:comments {W}>'
            + comment(0, "最初の 𠮷 コメント") + comment(1, "Second comment") + '</w:comments>'),
        "word/document.xml": (
            f'<?xml version="1.0" encoding="UTF-8"?><w:document {W}><w:body><w:p>'
            '<w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' + anchored(0, "beta")
            + '<w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p><w:p>'
            + anchored(1, "二段落目") + '</w:p></w:body></w:document>'),
    }
    SAMPLE.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(SAMPLE, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)


def open_sample():
    write_sample()
    w.open(file_name=str(SAMPLE), add_to_recent_files=False)
    return w.documents[SAMPLE.name]


def comment_read():
    d = open_sample()
    try:
        print("count", d.count(each=k.Word_comment))
        attempt("list", lambda: comments(d))
        print("paragraphs", [para(d, i) for i in (1, 2)], [bounds(d, i) for i in (1, 2)])
        c = d.Word_comments[1]
        attempt("all properties", lambda: sorted(str(x) for x in c.properties.get()))
    finally:
        d.close(saving=k.no)


def comment_edit():
    d = open_sample()
    try:
        c = d.Word_comments[1]
        attempt("set the text", lambda: c.comment_text.content.set("書いた 𠮷 text"))
        attempt("list", lambda: comments(d))
    finally:
        d.close(saving=k.no)


def comment_delete():
    d = open_sample()
    try:
        attempt("delete the first", lambda: d.Word_comments[1].delete())
        attempt("list", lambda: comments(d))
        print("paragraphs", [para(d, i) for i in (1, 2)], [bounds(d, i) for i in (1, 2)])
    finally:
        d.close(saving=k.no)


def ruby():
    """A ruby word marked deleted: is its field still listed, and where is its code?"""
    sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
    from ms_word_mcp import backend
    from ms_word_mcp.backend import call_document
    from ms_word_mcp.tools import _common, documents, page, text

    def run(label, impl, *args, undo=None):
        out = json.loads(_common.run(impl, *args, undo=undo))
        print(f"{label}\n   {json.dumps(out, ensure_ascii=False)[:900]}")
        return out

    def fields(d):
        t = d.paragraphs[1].text_object
        out = []
        for i in range(1, int(t.count(each=k.field)) + 1):
            c = t.fields[i].field_code
            out.append((c.content.get(), c.start_of_content.get(), c.end_of_content.get()))
        return out

    backend.word.start()
    made = run("create", documents._create_document_impl, None, None)
    token = call_document.set(made["name"])
    try:
        run("insert", text._insert_paragraphs_impl, [{"text": "OLD and NEW and END"}], None, None,
            undo="Insert")
        run("ruby", page._add_ruby_impl, 1,
            [{"text": "OLD", "ruby": "o"}, {"text": "NEW", "ruby": "n"}, {"text": "END", "ruby": "e"}],
            "center", None, undo="Ruby")
        d = w.documents[made["name"]]
        print("before", repr(para(d, 1)), bounds(d, 1), fields(d))
        d.track_revisions.set(True)
        attempt("delete the first field", lambda: d.paragraphs[1].text_object.fields[1].delete())
        print("after", repr(para(d, 1)), bounds(d, 1))
        attempt("fields", lambda: fields(d))
        attempt("revisions", lambda: revisions(d))
        run("read", text._get_paragraphs_impl, 1, 1, False)
        d.track_revisions.set(False)
    finally:
        run("close", documents._close_document_impl, False)
        call_document.reset(token)
        backend.word.stop()


def revision_reads():
    """How to read the deletions of a part of a document, and what each way costs.

    Asking for a property of every revision at once (`d.revisions.revision_type`)
    ran past the 60 s timeout on 200 revisions and left Word busy for minutes
    (measured 2026-11-10), so that is not tried again. Counting the revisions
    of a range answers 0 even with deletions in it; its `revisions` list is right.
    """
    def timed(label, f):
        began = time.monotonic()
        out = attempt(label, f)
        print(f"   {time.monotonic() - began:.3f} s")
        return out

    def where(d, i):
        t = d.revisions[i].text_object
        return t.start_of_content.get(), t.end_of_content.get()

    for size in (20, 180):
        print(f"\n== {size} deletions")
        d = w.make(new=k.document)
        try:
            d.text_object.content.set(
                "\r".join(f"Paragraph {i:03} alpha beta" for i in range(1, size + 21)))
            d.track_revisions.set(True)
            for i in range(size, 0, -1):
                a, _ = bounds(d, i)
                d.create_range(start=a + 14, end_=a + 20).content.set("")
            a, _ = bounds(d, size + 5)
            _, b = bounds(d, size + 7)
            d.create_range(start=a + 5, end_=b - 5).content.set("")  # across three paragraphs
            n = int(d.count(each=k.revision))
            print("revisions", n)
            timed("count", lambda: d.count(each=k.revision))
            for i in (1, n // 2, n):
                timed(f"where is {i}", lambda: where(d, i))
                timed(f"type of {i}", lambda: d.revisions[i].revision_type.get())
            def in_range(r):
                listed = r.revisions.get()
                print("   listed", type(listed).__name__, str(listed)[:160])
                out = []
                for i in range(1, len(listed) + 1 if isinstance(listed, list) else 1):
                    t = r.revisions[i].text_object
                    out.append((_name_of(r.revisions[i].revision_type.get()),
                                t.start_of_content.get(), t.end_of_content.get()))
                return out

            para = lambda i: d.paragraphs[i].text_object  # noqa: E731
            span = lambda i, j: d.create_range(  # noqa: E731
                start=bounds(d, i)[0], end_=bounds(d, j)[1])
            print("paragraph 5", bounds(d, 5))
            timed("paragraph 5", lambda: in_range(para(5)))
            print("paragraphs 5-8", bounds(d, 5)[0], bounds(d, 8)[1])
            timed("paragraphs 5-8", lambda: in_range(span(5, 8)))
            last = size - 2
            print(f"paragraphs {last}-{size}", bounds(d, last)[0], bounds(d, size)[1])
            timed("near the end", lambda: in_range(span(last, size)))
            print("the middle of three", bounds(d, size + 6))
            timed("the middle of three", lambda: in_range(para(size + 6)))
            timed("paragraphs around the three", lambda: in_range(span(size + 4, size + 8)))
            timed("none", lambda: in_range(para(size + 12)))
            timed("a match inside a deletion", lambda: in_range(d.create_range(
                start=bounds(d, 3)[0] + 15, end_=bounds(d, 3)[0] + 18)))
            timed("text next to a deletion", lambda: in_range(d.create_range(
                start=bounds(d, 3)[0], end_=bounds(d, 3)[0] + 14)))
            timed("the whole document", lambda: len(in_range(d.text_object)) if size == 20 else None)
            d.track_revisions.set(False)
        finally:
            d.close(saving=k.no)


def _name_of(value):
    return getattr(value, "name", str(value))


SECTIONS = {"tracking": tracking, "tools": tools, "ruby": ruby, "revision-reads": revision_reads, "comment-at-document": comment_at_document,
            "comment-read": comment_read, "comment-edit": comment_edit,
            "comment-delete": comment_delete}

if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in SECTIONS:
        raise SystemExit(f"Say which: {', '.join(SECTIONS)}")
    if not word_is_empty():
        raise SystemExit("Another document is open in Word. Nothing was run.")
    SECTIONS[sys.argv[1]]()
