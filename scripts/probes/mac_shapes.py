"""Shapes through Word for Mac's scripting: making them, and where they land.

    uv run python scripts/probes/mac_shapes.py

A scratch document, closed unsaved. Each way of doing it is tried and its answer
printed. The dictionary names the classes but not how to make one, and `shapes`
runs in z-order, so a shape is held by its name throughout.
"""

import os
import appscript
from appscript import k
from ms_word_mcp.backend.mac import _decode_wide_integers
from ms_word_mcp.tools import screen

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")
CONTAINER = os.path.expanduser("~/Library/Containers/com.microsoft.Word/Data/Documents")
print("dash names", [n for n in w.AS_appdata.referencebyname() if "dash" in n or "adjust" in n or "vertical_anchor" in n or "auto_size" in n])


def attempt(label, f):
    try:
        out = f()
        print(f"{label} -> {out!r}")
        return out
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {' '.join(x.strip() for x in str(e).splitlines()[:3])[:220]}")
        return None


TEXT = ("One two three four five six seven eight nine ten. " * 8 + "\r"
        "Second paragraph here.\rThird paragraph here.\r"
        "Fourth paragraph here.\rFifth paragraph here.\rSixth paragraph here.\r")


def starts(d):
    n = int(d.count(each=k.paragraph))
    return [d.paragraphs[i].text_object.start_of_content.get() for i in range(1, n + 1)]


def show(s):
    out = {}
    for name, get in [
            ("name", s.name), ("type", s.shape_type), ("left", s.left_position), ("top", s.top),
            ("w", s.width), ("h", s.height), ("relh", s.relative_horizontal_position),
            ("relv", s.relative_vertical_position), ("wrap", s.wrap_format.wrap_type),
            ("z", s.z_order_position), ("anchor", s.anchor.start_of_content)]:
        try:
            v = get.get()
            out[name] = str(getattr(v, "name", v)).replace("relative_", "").replace("_position", "")
        except Exception as e:  # noqa: BLE001
            out[name] = f"!{str(e).splitlines()[0][:40]}"
    return out


def names(d):
    n = int(d.count(each=k.shape))
    return [d.shapes[i].name.get() for i in range(1, n + 1)]


def counts(d):
    return {"shapes": int(d.count(each=k.shape)), "inline": int(d.count(each=k.inline_shape))}


def box(d, para, kind=k.autoshape_rectangle, **extra):
    return {k.auto_shape_type: kind, k.left_position: 100, k.top: 50, k.width: 120, k.height: 60,
            k.anchor: d.paragraphs[para].text_object, **extra}


def by(d, name):
    return d.shapes[name]


d = w.create_new_document()
picture = os.path.join(CONTAINER, "mac-shapes-probe.png")
try:
    d.text_object.content.set(TEXT)
    print("\n== anchors")
    print("starts before", starts(d))
    print("para 3", repr(d.paragraphs[3].text_object.content.get()))
    a = d.make(new=k.shape, at=d, with_properties=box(d, 3))
    an = a.name.get()
    print(an, show(by(d, an)))
    print("starts after", starts(d))
    print("para 3", repr(d.paragraphs[3].text_object.content.get()),
          [ord(c) for c in d.paragraphs[3].text_object.content.get()[:3]])
    p3 = d.paragraphs[3].text_object
    print("para 3 positions", p3.end_of_content.get() - p3.start_of_content.get(), "text len", len(p3.content.get()))
    b = d.make(new=k.shape, at=d, with_properties=box(d, 4, k.autoshape_oval))
    bn = b.name.get()
    print(bn, show(by(d, bn)), starts(d))
    # anchored mid paragraph by a collapsed range
    s4 = d.paragraphs[4].text_object.start_of_content.get()
    c = attempt("anchor: collapsed range", lambda: d.make(new=k.shape, at=d, with_properties={
        **box(d, 4), k.anchor: d.create_range(start=s4 + 5, end_=s4 + 5)}))
    if c:
        print(show(by(d, c.name.get())))
    tb = attempt("text box anchored", lambda: d.make(new=k.text_box, at=d, with_properties={
        k.left_position: 200, k.top: 200, k.width: 150, k.height: 80,
        k.anchor: d.paragraphs[2].text_object}))
    if tb:
        print(names(d))
    ln = attempt("line anchored", lambda: d.make(new=k.line_shape, at=d, with_properties={
        k.begin_line_X: 50, k.begin_line_Y: 400, k.end_line_X: 300, k.end_line_Y: 450,
        k.anchor: d.paragraphs[5].text_object}))
    ln2 = attempt("backwards line", lambda: d.make(new=k.line_shape, at=d, with_properties={
        k.begin_line_X: 300, k.begin_line_Y: 400, k.end_line_X: 50, k.end_line_Y: 350,
        k.anchor: d.paragraphs[5].text_object}))
    for n in names(d):
        print("  ", show(by(d, n)))
    print("starts", starts(d))
    # a line's flips
    last = names(d)[-1]
    attempt("flips", lambda: (by(d, last).horizontal_flip.get(), by(d, last).vertical_flip.get()))
    attempt("line ends", lambda: (by(d, last).begin_line_X.get(), by(d, last).begin_line_Y.get(),
                                  by(d, last).end_line_X.get(), by(d, last).end_line_Y.get()))

    print("\n== find after an anchor")
    r = d.create_range(start=starts(d)[3], end_=starts(d)[4])
    attempt("content", lambda: r.content.get())
    p = starts(d)[3]
    attempt("range over 'paragraph' if no anchor char", lambda: d.create_range(
        start=p + len("Fourth "), end_=p + len("Fourth paragraph")).content.get())
    attempt("characters count", lambda: int(r.count(each=k.character)))

    print("\n== look")
    s = by(d, an)
    attempt("set fill 16-bit", lambda: s.fill_format.fore_color.set([0xFFFF, 0x8080, 0x4040]))
    attempt("fill", lambda: s.fill_format.fore_color.get())
    attempt("set fill 8-bit", lambda: s.fill_format.fore_color.set([255, 128, 64]))
    attempt("fill", lambda: s.fill_format.fore_color.get())
    attempt("set line", lambda: s.line_format.fore_color.set([0x1010, 0x2020, 0xF0F0]))
    attempt("line", lambda: s.line_format.fore_color.get())
    attempt("dash_style_", lambda: s.line_format.dash_style_.get() if hasattr(s.line_format, "dash_style_") else None)
    attempt("line dash style", lambda: s.line_format.line_dash_style.get())
    attempt("set dash", lambda: (s.line_format.dash_style.set(k.line_dash_style_dash)))
    attempt("adjustments count", lambda: int(s.count(each=k.adjustment)))
    attempt("adjustment 1", lambda: s.adjustments[1].get())
    attempt("rounded adj", lambda: (s.auto_shape_type.set(k.autoshape_rounded_rectangle),
                                    int(s.count(each=k.adjustment)), s.adjustments[1].properties.get()))
    attempt("vertical anchor", lambda: s.text_frame.vertical_anchor.get())
    attempt("auto size", lambda: s.text_frame.auto_size.get())
    attempt("word wrap", lambda: s.text_frame.word_wrap.get())
    attempt("two paragraphs", lambda: (s.text_frame.text_range.content.set("日本語\r二段落目"),
                                       s.text_frame.text_range.content.get(),
                                       int(s.text_frame.text_range.count(each=k.paragraph))))
    attempt("bounds", lambda: (s.text_frame.text_range.start_of_content.get(),
                               s.text_frame.text_range.end_of_content.get()))
    attempt("doc paragraphs", lambda: int(d.count(each=k.paragraph)))

    print("\n== names")
    attempt("same name", lambda: (by(d, bn).name.set(an), names(d)))
    attempt("by that name", lambda: by(d, an).auto_shape_type.get())
    attempt("rename", lambda: (d.shapes[2].name.set(bn), names(d)))

    print("\n== wrap and order")
    print(names(d))
    for wt in (k.wrap_behind, k.wrap_square, k.wrap_front):
        attempt(str(wt), lambda wt=wt: (by(d, an).wrap_format.wrap_type.set(wt),
                                        by(d, an).wrap_format.wrap_type.get(),
                                        by(d, an).z_order_position.get()))
        print("   order", names(d))
    attempt("to back", lambda: by(d, bn).z_order(z_order_command=k.send_shape_to_back))
    print("   order", names(d), by(d, bn).z_order_position.get())
    attempt("to front", lambda: by(d, bn).z_order(z_order_command=k.bring_shape_to_front))
    print("   order", names(d), by(d, bn).z_order_position.get())
    attempt("forward", lambda: by(d, an).z_order(z_order_command=k.bring_shape_forward))
    print("   order", names(d), by(d, an).z_order_position.get())
    attempt("inline on autoshape", lambda: (by(d, an).wrap_format.wrap_type.set(k.wrap_inline),
                                            by(d, an).wrap_format.wrap_type.get()))
    print("   order", names(d), counts(d), show(by(d, an)), starts(d))
    attempt("back to square", lambda: (by(d, an).wrap_format.wrap_type.set(k.wrap_square),
                                       by(d, an).wrap_format.wrap_type.get()))
    print("   order", names(d), counts(d), show(by(d, an)), starts(d))

    print("\n== pictures")
    with open(picture, "wb") as f:
        f.write(screen.png(60, 30, [b"\x40\x80\xc0" * 60] * 30))
    direct = attempt("floating picture anchored", lambda: d.make(new=k.picture, at=d, with_properties={
        k.file_name: picture, k.link_to_file: False, k.save_with_document: True,
        k.left_position: 300, k.top: 500, k.width: 100, k.height: 50,
        k.anchor: d.paragraphs[3].text_object}))
    print(counts(d), names(d))
    pn = direct.name.get() if direct else None
    if pn:
        print(show(by(d, pn)))
    at = starts(d)[1]
    d.make(new=k.inline_picture, at=d.create_range(start=at, end_=at), with_properties={
        k.file_name: picture, k.link_to_file: False, k.save_with_document: True})
    attempt("alt on inline", lambda: (d.inline_shapes[1].alternative_text.set("A blue box"),
                                      d.inline_shapes[1].alternative_text.get()))
    print("inline", counts(d), repr(d.paragraphs[2].text_object.content.get()[:8]), starts(d)[:3])
    made = attempt("convert to shape", lambda: d.inline_shapes[1].convert_to_shape())
    print(counts(d), names(d), repr(d.paragraphs[2].text_object.content.get()[:8]), starts(d)[:3])
    fn = made.name.get() if made else None
    print("made ref", made, fn)
    if fn:
        f2 = by(d, fn)
        print(show(f2))
        attempt("alt kept?", lambda: f2.alternative_text.get())
        for wt in (k.wrap_behind, k.wrap_front, k.wrap_square):
            attempt(str(wt), lambda wt=wt: (by(d, fn).wrap_format.wrap_type.set(wt),
                                            by(d, fn).wrap_format.wrap_type.get()))
        attempt("crop", lambda: (by(d, fn).picture_format.crop_left.set(5),
                                 by(d, fn).picture_format.crop_left.get(), by(d, fn).width.get()))
        attempt("crop on class picture", lambda: d.pictures[fn].picture_format.crop_left.get())
        back = attempt("convert to inline", lambda: by(d, fn).convert_to_inline_shape())
        print(counts(d), names(d), repr(d.paragraphs[2].text_object.content.get()[:8]), starts(d)[:3])
        attempt("alt after coming back", lambda: d.inline_shapes[1].alternative_text.get())
    if pn:
        attempt("wrap inline on picture", lambda: (by(d, pn).wrap_format.wrap_type.set(k.wrap_inline),
                                                   by(d, pn).wrap_format.wrap_type.get()))
        print(counts(d), names(d))
    attempt("autoshape to inline", lambda: by(d, bn).convert_to_inline_shape())
    print(counts(d), names(d))

    print("\n== an inactive document")
    other = w.create_new_document()
    try:
        print("active", w.active_document.name.get(), "scratch", d.name.get())
        x = attempt("make", lambda: d.make(new=k.shape, at=d, with_properties=box(d, 5)))
        print("active still", w.active_document.name.get(), counts(d), "other", counts(other))
        if x:
            print(show(by(d, x.name.get())))
    finally:
        other.close(saving=k.no)

    print("\n== a table cell, undo")
    t = attempt("table", lambda: d.make(new=k.table, at=d.paragraphs[6].text_object.end_of_content if False else d.paragraphs[6].text_object, with_properties={k.number_of_rows: 2, k.number_of_columns: 2}))
    if t:
        c = attempt("anchored in a cell", lambda: d.make(new=k.shape, at=d, with_properties={
            k.auto_shape_type: k.autoshape_rectangle, k.left_position: 10, k.top: 10,
            k.width: 50, k.height: 30, k.anchor: d.tables[1].get_cell_from_table(row=2, column=2).text_object}))
        print(show(by(d, c.name.get())) if c else None)
    before = counts(d)
    x = d.make(new=k.shape, at=d, with_properties=box(d, 2))
    attempt("undo", lambda: d.undo(times=1))
    print("undo", before, counts(d))
finally:
    d.close(saving=k.no)
    if os.path.exists(picture):
        os.remove(picture)
