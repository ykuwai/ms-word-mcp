"""Footnotes, endnotes, hyperlinks and bookmarks through Word for Mac's scripting.

    uv run python scripts/probes/mac_notes_links.py

A scratch document, closed unsaved. Each way of doing it is tried and its
answer printed, since the dictionary names the classes but not how to make one.
"""

import appscript
from appscript import k

from ms_word_mcp.backend.mac import _decode_wide_integers

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")


def attempt(label, f):
    try:
        print(f"{label} -> {f()!r}")
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {' '.join(x.strip() for x in str(e).splitlines()[:3])[:200]}")


def bounds(d, i):
    t = d.paragraphs[i].text_object
    return t.start_of_content.get(), t.end_of_content.get()


def para(d, i):
    return d.paragraphs[i].text_object.content.get()


def notes(d, kind, label):
    n = int(d.count(each=kind))
    out = []
    plural = d.footnotes if kind == k.footnote else d.endnotes
    for i in range(1, n + 1):
        f = plural[i]
        out.append((f.entry_index.get(), f.note_reference.start_of_content.get(),
                    f.text_object.content.get()))
    print(f"{label}: {n} {out}")


d = w.create_new_document()
try:
    d.text_object.content.set("One two three.\rFour five six.\rSeven eight nine.\r")
    print("paragraphs", int(d.count(each=k.paragraph)), bounds(d, 1), bounds(d, 2))

    # -- footnotes --------------------------------------------------------------------
    s, e = bounds(d, 1)
    place = d.create_range(start=e - 1, end_=e - 1)
    attempt("footnote at document, text object", lambda: d.make(
        new=k.footnote, at=d, with_properties={k.text_object: place}))
    notes(d, k.footnote, "after a")
    attempt("footnote at range", lambda: d.make(
        new=k.footnote, at=d.create_range(start=e - 1, end_=e - 1)))
    notes(d, k.footnote, "after b")
    attempt("footnote at end of paragraph", lambda: d.make(
        new=k.footnote, at=d.paragraphs[2].text_object.end))
    notes(d, k.footnote, "after c")

    if int(d.count(each=k.footnote)):
        f = d.footnotes[1]
        attempt("set note text by content", lambda: f.text_object.content.set("First note"))
        attempt("read note text", lambda: f.text_object.content.get())
        attempt("note range bounds", lambda: (f.text_object.start_of_content.get(),
                                              f.text_object.end_of_content.get()))
        attempt("note story", lambda: f.text_object.story_type.get())
        attempt("note style", lambda: f.text_object.style.name_local.get())
        attempt("ref style", lambda: f.note_reference.style.name_local.get())
        attempt("ref content", lambda: f.note_reference.content.get())
        attempt("paragraph with the mark", lambda: para(d, 1))
        attempt("paragraph bounds now", lambda: bounds(d, 1))

    # Inserted before an existing note: does it renumber?
    s2, e2 = bounds(d, 1)
    attempt("a note earlier in the text", lambda: d.make(
        new=k.footnote, at=d, with_properties={k.text_object: d.create_range(start=s2 + 3,
                                                                             end_=s2 + 3)}))
    notes(d, k.footnote, "after an earlier one")

    # -- endnotes ---------------------------------------------------------------------
    s3, e3 = bounds(d, 2)
    attempt("endnote at document", lambda: d.make(
        new=k.endnote, at=d, with_properties={k.text_object: d.create_range(start=e3 - 1,
                                                                            end_=e3 - 1)}))
    notes(d, k.endnote, "endnotes")
    if int(d.count(each=k.endnote)):
        attempt("set endnote text", lambda: d.endnotes[1].text_object.content.set("An endnote"))
        notes(d, k.endnote, "endnotes, written")

    # -- deleting ---------------------------------------------------------------------
    n = int(d.count(each=k.footnote))
    if n:
        attempt("delete a footnote", lambda: d.footnotes[n].delete())
        notes(d, k.footnote, "after delete")
        attempt("paragraph after delete", lambda: para(d, 1))

    # -- hyperlinks -------------------------------------------------------------------
    s4, e4 = bounds(d, 3)
    word = d.create_range(start=s4, end_=s4 + 5)
    print("over:", word.content.get())
    attempt("hyperlink over text", lambda: d.make(
        new=k.hyperlink_object, at=d, with_properties={
            k.text_object: word, k.hyperlink_address: "https://example.com/a?b=1&c=2",
            k.screen_tip: "A tip"}))
    print("count", int(d.count(each=k.hyperlink_object)))
    s5, e5 = bounds(d, 3)
    attempt("hyperlink with display text at a point", lambda: d.make(
        new=k.hyperlink_object, at=d, with_properties={
            k.text_object: d.create_range(start=e5 - 1, end_=e5 - 1),
            k.hyperlink_address: "https://example.com/z",
            k.text_to_display: " shown text"}))
    n = int(d.count(each=k.hyperlink_object))
    print("count", n, "paragraph", para(d, 3), bounds(d, 3))
    for i in range(1, n + 1):
        h = d.hyperlink_objects[i]
        attempt(f"hyperlink {i}", lambda: (
            h.hyperlink_address.get(), h.sub_address.get(), h.screen_tip.get(),
            h.text_to_display.get(), h.text_object.start_of_content.get(),
            h.text_object.end_of_content.get(), h.hyperlink_type.get()))
    attempt("hyperlink style", lambda: d.hyperlink_objects[1].text_object.style.name_local.get())
    attempt("change the address", lambda: d.hyperlink_objects[1].hyperlink_address.set(
        "https://example.com/changed"))
    attempt("read it", lambda: d.hyperlink_objects[1].hyperlink_address.get())
    attempt("fields there", lambda: int(d.count(each=k.field)))

    # -- bookmarks --------------------------------------------------------------------
    s6, e6 = bounds(d, 2)
    attempt("bookmark", lambda: d.make(new=k.bookmark, at=d, with_properties={
        k.name: "Mine", k.text_object: d.create_range(start=s6, end_=s6 + 4)}))
    attempt("hidden bookmark", lambda: d.make(new=k.bookmark, at=d, with_properties={
        k.name: "_Hidden1", k.text_object: d.create_range(start=s6 + 5, end_=s6 + 8)}))
    attempt("a point bookmark", lambda: d.make(new=k.bookmark, at=d, with_properties={
        k.name: "Spot", k.text_object: d.create_range(start=s6 + 2, end_=s6 + 2)}))
    attempt("same name again", lambda: d.make(new=k.bookmark, at=d, with_properties={
        k.name: "Mine", k.text_object: d.create_range(start=s6 + 9, end_=s6 + 12)}))
    attempt("a bad name", lambda: d.make(new=k.bookmark, at=d, with_properties={
        k.name: "two words", k.text_object: d.create_range(start=s6, end_=s6 + 1)}))
    n = int(d.count(each=k.bookmark))
    print("bookmarks", n)
    for i in range(1, n + 1):
        b = d.bookmarks[i]
        attempt(f"bookmark {i}", lambda: (b.name.get(), b.start_of_bookmark.get(),
                                          b.end_of_bookmark.get(), b.empty.get(),
                                          b.text_object.content.get()))
    attempt("by name", lambda: d.bookmarks["Mine"].text_object.content.get())
    attempt("delete by name", lambda: d.bookmarks["Spot"].delete())
    attempt("left", lambda: int(d.count(each=k.bookmark)))
    attempt("link to a bookmark", lambda: d.make(
        new=k.hyperlink_object, at=d, with_properties={
            k.text_object: d.create_range(start=s6 + 9, end_=s6 + 12),
            k.hyperlink_address: "", k.sub_address: "Mine"}))
    n = int(d.count(each=k.hyperlink_object))
    attempt("that link", lambda: (d.hyperlink_objects[n].hyperlink_address.get(),
                                  d.hyperlink_objects[n].sub_address.get()))
    attempt("delete a hyperlink", lambda: d.hyperlink_objects[1].delete())
    attempt("links left", lambda: int(d.count(each=k.hyperlink_object)))
    attempt("text stays", lambda: para(d, 3))
finally:
    d.close(saving=k.no)
