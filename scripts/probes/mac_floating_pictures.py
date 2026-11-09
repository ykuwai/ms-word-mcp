"""Pictures floating and in line, anchors in edited text, and undo, on Word for Mac.

    uv run python scripts/probes/mac_floating_pictures.py

A scratch document, closed unsaved, and one small picture in Word's container,
removed at the end.
"""

import os
import appscript
from appscript import k
from ms_word_mcp.backend.mac import _decode_wide_integers
from ms_word_mcp.tools import screen

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")
CONTAINER = os.path.expanduser("~/Library/Containers/com.microsoft.Word/Data/Documents")


def attempt(label, f):
    try:
        out = f()
        print(f"{label} -> {out!r}")
        return out
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {' '.join(x.strip() for x in str(e).splitlines()[:3])[:220]}")
        return None


TEXT = ("Alpha. " * 100 + "\rSecond paragraph here.\rThird paragraph here.\rFourth paragraph here.\r")


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


d = w.create_new_document()
picture = os.path.join(CONTAINER, "mac-shapes-probe.png")
with open(picture, "wb") as f:
    f.write(screen.png(60, 30, [b"\x40\x80\xc0" * 60] * 30))
try:
    d.text_object.content.set(TEXT)
    at = starts(d)[2]
    print(starts(d))
    print("\n== a picture from in line to floating and back")
    d.make(new=k.inline_picture, at=d.create_range(start=at, end_=at), with_properties={
        k.file_name: picture, k.link_to_file: False, k.save_with_document: True})
    d.inline_shapes[1].alternative_text.set("A blue box")
    print("in line", counts(d), starts(d), repr(d.paragraphs[3].text_object.content.get()[:8]))
    before = set(names(d))
    d.inline_shapes[1].convert_to_shape()
    fn = (set(names(d)) - before).pop()
    print("floating:", fn, counts(d), starts(d), repr(d.paragraphs[3].text_object.content.get()[:8]))
    print(show(d.shapes[fn]))
    attempt("class", lambda: d.shapes[fn].class_.get())
    p = d.shapes[fn]
    attempt("alt", lambda: p.alternative_text.get())
    attempt("crop by shapes", lambda: p.picture_format.crop_left.get())
    attempt("crop by shape props", lambda: p.properties.get())
    attempt("crop by pictures", lambda: d.pictures[fn].picture_format.crop_left.get())
    attempt("set crop by pictures", lambda: (d.pictures[fn].picture_format.crop_left.set(5),
                                             d.pictures[fn].picture_format.crop_left.get(),
                                             d.shapes[fn].width.get()))
    attempt("inline crop", lambda: 0)
    for wt in (k.wrap_behind, k.wrap_front, k.wrap_square, k.wrap_tight, k.wrap_top_bottom):
        attempt(str(wt), lambda wt=wt: (p.wrap_format.wrap_type.set(wt), p.wrap_format.wrap_type.get()))
    attempt("page relative, center", lambda: (
        p.relative_horizontal_position.set(k.relative_horizontal_position_page),
        p.relative_vertical_position.set(k.relative_vertical_position_page),
        p.left_position.set(-999995), p.top.set(200), p.left_position.get(), p.top.get()))
    print(show(p), starts(d))
    print("to inline by convert")
    p.convert_to_inline_shape()
    print(counts(d), names(d), starts(d), repr(d.paragraphs[3].text_object.content.get()[:8]))
    attempt("alt back in line", lambda: d.inline_shapes[1].alternative_text.get())
    attempt("size back in line", lambda: (d.inline_shapes[1].width.get(), d.inline_shapes[1].height.get()))
    attempt("inline start", lambda: d.inline_shapes[1].text_object.start_of_content.get())

    print("\n== by wrap type, floating then in line")
    before = set(names(d))
    d.inline_shapes[1].convert_to_shape()
    fn = (set(names(d)) - before).pop()
    print(fn, starts(d))
    d.shapes[fn].wrap_format.wrap_type.set(k.wrap_inline)
    print("wrap inline:", counts(d), names(d), starts(d))

    print("\n== a shape anchored mid paragraph, then editing that paragraph")
    p2s = starts(d)[1]
    s = d.make(new=k.shape, at=d, with_properties={
        k.auto_shape_type: k.autoshape_rectangle, k.left_position: 10, k.top: 10, k.width: 50,
        k.height: 30, k.anchor: d.paragraphs[2].text_object})
    name = s.name.get()
    print(show(d.shapes[name]), starts(d))
    tr = d.paragraphs[2].text_object
    print("content", repr(tr.content.get()), "span", tr.start_of_content.get(), tr.end_of_content.get())
    print("first chars", [d.paragraphs[2].characters[i].content.get() for i in range(1, 4)],
          [d.paragraphs[2].characters[i].start_of_content.get() for i in range(1, 4)])
    attempt("rewrite the paragraph", lambda: d.create_range(
        start=tr.start_of_content.get(), end_=tr.end_of_content.get() - 1).content.set("New words here."))
    print(counts(d), names(d), starts(d), repr(d.paragraphs[2].text_object.content.get()))

    print("\n== undo")
    before = counts(d)
    s = d.make(new=k.shape, at=d, with_properties={
        k.auto_shape_type: k.autoshape_oval, k.left_position: 10, k.top: 10, k.width: 50,
        k.height: 30, k.anchor: d.paragraphs[3].text_object})
    print(before, counts(d))
    attempt("can undo", lambda: d.can_undo.get() if hasattr(d, "can_undo") else None)
    attempt("undo", lambda: d.undo(times=1))
    print(counts(d))
    attempt("undo", lambda: d.undo(times=1))
    print(counts(d))

    print("\n== a rectangle's default look, an oval's text position")
    s = d.make(new=k.shape, at=d, with_properties={
        k.auto_shape_type: k.autoshape_rectangle, k.left_position: 10, k.top: 10, k.width: 150,
        k.height: 60, k.anchor: d.paragraphs[4].text_object})
    n = s.name.get()
    t = d.shapes[n].text_frame.text_range
    t.content.set("中央")
    attempt("default para alignment", lambda: t.paragraph_format.alignment.get())
    attempt("default color", lambda: t.font_object.color.get())
    attempt("default color index", lambda: t.font_object.color_index.get())
    attempt("set bold, name", lambda: (t.font_object.bold.set(True), t.font_object.east_asian_font_name.get(), t.font_object.ascii_font_name.get() if False else None))
    attempt("set font color", lambda: t.font_object.color.set([0, 0, 0]))
    attempt("read", lambda: t.font_object.color.get())
    attempt("story of a range made from positions", lambda: t.story_type.get())
    attempt("fill default", lambda: (d.shapes[n].fill_format.fore_color.get(), d.shapes[n].line_format.fore_color.get(), d.shapes[n].line_format.weight.get()))
finally:
    d.close(saving=k.no)
    if os.path.exists(picture):
        os.remove(picture)
