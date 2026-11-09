"""Shapes, text boxes, lines and floating pictures on Windows.

    uv run --no-sync python scripts/probes/shapes_windows.py

Scratch documents, closed unsaved, and one small picture written to the temp
folder and removed. Everything is printed: where a new shape is anchored,
whether an anchor takes a position in the text, how the Shapes collection is
ordered, what each wrap type does, and how a picture goes floating and back.
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.tools import _common, screen  # noqa: E402

RECT, OVAL, ROUNDED = 1, 9, 5
TEXT = ("One two three four five six seven eight nine ten. " * 8 + "\r"
        "Second paragraph here.\rThird paragraph here.\r"
        "Fourth paragraph here.\rFifth paragraph here.\rSixth paragraph here.")


def attempt(label, f):
    try:
        out = f()
        print(f"{label} -> {out!r}")
        return out
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {' '.join(str(e).split())[:240]}")
        return None


def new_doc():
    doc = backend.word.get_app().Documents.Add()
    doc.Content.Text = TEXT
    return doc


def starts(doc):
    return [doc.Paragraphs.Item(i).Range.Start for i in range(1, doc.Paragraphs.Count + 1)]


def para_range(doc, i):
    return doc.Paragraphs.Item(i).Range


def info(s):
    out = {}
    for name, get in [
            ("name", lambda: s.Name), ("id", lambda: s.ID), ("type", lambda: s.Type),
            ("auto", lambda: s.AutoShapeType), ("left", lambda: round(s.Left, 2)),
            ("top", lambda: round(s.Top, 2)), ("w", lambda: round(s.Width, 2)),
            ("h", lambda: round(s.Height, 2)), ("relh", lambda: s.RelativeHorizontalPosition),
            ("relv", lambda: s.RelativeVerticalPosition), ("wrap", lambda: s.WrapFormat.Type),
            ("side", lambda: s.WrapFormat.Side), ("z", lambda: s.ZOrderPosition),
            ("rot", lambda: s.Rotation), ("anchor", lambda: s.Anchor.Start)]:
        try:
            out[name] = get()
        except Exception as e:  # noqa: BLE001
            out[name] = f"!{' '.join(str(e).split())[:40]}"
    return out


def names(doc):
    return [doc.Shapes.Item(i).Name for i in range(1, doc.Shapes.Count + 1)]


def counts(doc):
    return {"shapes": doc.Shapes.Count, "inline": doc.InlineShapes.Count}


def anchors_and_text():
    print("\n== where a shape is anchored; whether an anchor takes a position")
    doc = new_doc()
    try:
        print("starts before", starts(doc))
        print("para 3 before", repr(para_range(doc, 3).Text))
        a = attempt("no anchor", lambda: doc.Shapes.AddShape(RECT, 100, 50, 120, 60))
        print(info(a))
        b = attempt("anchor: paragraph 3's range", lambda: doc.Shapes.AddShape(
            RECT, 100, 50, 120, 60, para_range(doc, 3)))
        print(info(b))
        p4 = para_range(doc, 4).Start
        c = attempt("anchor: collapsed in paragraph 4, 5 chars in", lambda: doc.Shapes.AddShape(
            OVAL, 100, 50, 120, 60, doc.Range(p4 + 5, p4 + 5)))
        print(info(c))
        print("starts after", starts(doc))
        print("para 3 after", repr(para_range(doc, 3).Text))
        print("para 3 chars", [ord(ch) for ch in para_range(doc, 3).Text[:4]])
        print("para 4 after", repr(para_range(doc, 4).Text))
        r = para_range(doc, 4)
        print("para 4 length: text", len(r.Text), "positions", r.End - r.Start)
        r3 = para_range(doc, 3)
        print("para 3 length: text", len(r3.Text), "positions", r3.End - r3.Start)
        print("'paragraph' by offset in 3:", repr(doc.Range(
            r3.Start + len("Third "), r3.Start + len("Third paragraph")).Text))
        f = doc.Range(para_range(doc, 4).Start, para_range(doc, 4).End)
        attempt("find 'paragraph' in 4", lambda: (f.Find.Execute("paragraph"), f.Start, f.End,
                                                  f.Text))
        tb = attempt("text box anchored in 2", lambda: doc.Shapes.AddTextbox(
            1, 200, 200, 150, 80, para_range(doc, 2)))
        print(info(tb))
        ln = attempt("line anchored in 5", lambda: doc.Shapes.AddLine(
            50, 400, 300, 450, para_range(doc, 5)))
        print(info(ln))
        print(names(doc), "anchors", [doc.Shapes.Item(i).Anchor.Start
                                      for i in range(1, doc.Shapes.Count + 1)])
        print("paragraph count", doc.Paragraphs.Count)
    finally:
        doc.Close(0)


def position_and_order():
    print("\n== position, wrap, order")
    doc = new_doc()
    try:
        s = doc.Shapes.AddShape(RECT, 100, 50, 120, 60, para_range(doc, 3))
        t = doc.Shapes.AddShape(OVAL, 300, 50, 80, 60, para_range(doc, 2))
        u = doc.Shapes.AddTextbox(1, 300, 300, 150, 80, para_range(doc, 4))
        print("order", names(doc))
        sn, tn, un = s.Name, t.Name, u.Name
        print("default", info(s))
        s.RelativeHorizontalPosition = 1
        s.RelativeVerticalPosition = 1
        print("relative page, position kept?", info(s))
        s.Left, s.Top = 72, 144
        print("moved", info(s))
        s.Left = -999995
        print("center", info(s))
        s.RelativeHorizontalPosition = 0
        print("center, margin", info(s))
        s.Left = 72
        print("order before z", names(doc))
        attempt("to back", lambda: doc.Shapes.Item(un).ZOrder(1))
        print("order", names(doc), [info(doc.Shapes.Item(n))["z"] for n in (sn, tn, un)])
        attempt("to front", lambda: doc.Shapes.Item(un).ZOrder(0))
        print("order", names(doc), [info(doc.Shapes.Item(n))["z"] for n in (sn, tn, un)])
        attempt("behind text (5)", lambda: doc.Shapes.Item(tn).ZOrder(5))
        print("order", names(doc), info(doc.Shapes.Item(tn)))
        attempt("in front of text (4)", lambda: doc.Shapes.Item(tn).ZOrder(4))
        print("order", names(doc), info(doc.Shapes.Item(tn)))
        for w in (0, 1, 2, 4, 5, 3):
            attempt(f"wrap {w}", lambda w=w: (setattr(s.WrapFormat, "Type", w), s.WrapFormat.Type,
                                              round(s.Left, 2), s.ZOrderPosition))
            print("  order", names(doc))
        attempt("wrap 7 on an autoshape", lambda: (setattr(s.WrapFormat, "Type", 7),
                                                   s.WrapFormat.Type))
        print("after inline", counts(doc), info(doc.Shapes.Item(sn)) if sn in names(doc) else None)
        attempt("back to square", lambda: setattr(doc.Shapes.Item(sn).WrapFormat, "Type", 0))
        print(counts(doc), info(doc.Shapes.Item(sn)))
        attempt("side, distance", lambda: (setattr(s.WrapFormat, "Side", 1),
                                           setattr(s.WrapFormat, "DistanceLeft", 9),
                                           s.WrapFormat.DistanceLeft))
        attempt("the same name twice", lambda: (setattr(t, "Name", sn), t.Name, names(doc)))
        attempt("item by that name", lambda: info(doc.Shapes.Item(sn)))
        attempt("japanese name", lambda: (setattr(t, "Name", "図形の名前"), t.Name))
        attempt("delete the anchor paragraph", lambda: para_range(doc, 3).Delete())
        print("after", names(doc), counts(doc))
    finally:
        doc.Close(0)


def look_and_text():
    print("\n== look and text")
    doc = new_doc()
    try:
        s = doc.Shapes.AddShape(ROUNDED, 100, 50, 120, 60, para_range(doc, 3))
        attempt("fill", lambda: (s.Fill.ForeColor.RGB, s.Fill.Visible, s.Fill.Transparency))
        attempt("set fill", lambda: (setattr(s.Fill.ForeColor, "RGB", 0x4080FF),
                                     s.Fill.ForeColor.RGB, s.Fill.Visible))
        attempt("transparency", lambda: (setattr(s.Fill, "Transparency", 0.25),
                                         round(s.Fill.Transparency, 2)))
        attempt("no fill", lambda: (setattr(s.Fill, "Visible", 0), s.Fill.Visible))
        attempt("fill again, color kept?", lambda: (setattr(s.Fill, "Visible", -1),
                                                    s.Fill.ForeColor.RGB))
        attempt("line", lambda: (s.Line.ForeColor.RGB, s.Line.Weight, s.Line.DashStyle,
                                 s.Line.Visible))
        attempt("set line", lambda: (setattr(s.Line.ForeColor, "RGB", 0xF02010),
                                     setattr(s.Line, "Weight", 2.5), setattr(s.Line, "DashStyle", 4),
                                     s.Line.ForeColor.RGB, s.Line.Weight, s.Line.DashStyle))
        attempt("no line", lambda: (setattr(s.Line, "Visible", 0), s.Line.Visible))
        attempt("rotation", lambda: (setattr(s, "Rotation", 30), s.Rotation))
        attempt("adjustments", lambda: (s.Adjustments.Count, s.Adjustments.Item(1)))
        attempt("set adjustment", lambda: (s.Adjustments.__setitem__(1, 0.5),
                                           s.Adjustments.Item(1)))
        attempt("alt text", lambda: (setattr(s, "AlternativeText", "代替"), s.AlternativeText))
        attempt("title", lambda: s.Title)
        attempt("change type", lambda: (setattr(s, "AutoShapeType", OVAL), s.AutoShapeType))
        attempt("lock aspect", lambda: s.LockAspectRatio)

        tf = s.TextFrame
        attempt("has text", lambda: tf.HasText)
        attempt("set text", lambda: (setattr(tf.TextRange, "Text", "日本語 text\r二段落目"),
                                     tf.TextRange.Text))
        attempt("has text", lambda: tf.HasText)
        attempt("paragraphs, story", lambda: (tf.TextRange.Paragraphs.Count, tf.TextRange.StoryType))
        attempt("start, end", lambda: (tf.TextRange.Start, tf.TextRange.End))
        attempt("font", lambda: (tf.TextRange.Font.Name, tf.TextRange.Font.NameFarEast,
                                 tf.TextRange.Font.Size, tf.TextRange.Font.Color))
        attempt("set size, alignment", lambda: (setattr(tf.TextRange.Font, "Size", 16),
                                                setattr(tf.TextRange.ParagraphFormat, "Alignment", 1),
                                                tf.TextRange.ParagraphFormat.Alignment))
        attempt("vertical anchor", lambda: tf.VerticalAnchor)
        attempt("set middle", lambda: (setattr(tf, "VerticalAnchor", 3), tf.VerticalAnchor))
        attempt("auto size", lambda: tf.AutoSize)
        attempt("word wrap", lambda: tf.WordWrap)
        attempt("orientation", lambda: tf.Orientation)
        attempt("set vertical", lambda: (setattr(tf, "Orientation", 6), tf.Orientation))
        attempt("back", lambda: (setattr(tf, "Orientation", 1), tf.Orientation))
        attempt("margins", lambda: (setattr(tf, "MarginLeft", 10), tf.MarginLeft))
        attempt("overflowing", lambda: tf.Overflowing)
        print("doc paragraphs", doc.Paragraphs.Count, repr(para_range(doc, 1).Text[:20]))

        tb = doc.Shapes.AddTextbox(6, 300, 300, 60, 150, para_range(doc, 4))
        print("vertical text box", info(tb))
        attempt("its orientation", lambda: tb.TextFrame.Orientation)

        ln = doc.Shapes.AddLine(50, 400, 300, 450, para_range(doc, 5))
        print("line", info(ln))
        attempt("arrow", lambda: (setattr(ln.Line, "EndArrowheadStyle", 2), ln.Line.EndArrowheadStyle))
        attempt("text on a line", lambda: ln.TextFrame.HasText)
        attempt("fill on a line", lambda: ln.Fill.Visible)
        attempt("flips", lambda: (ln.HorizontalFlip, ln.VerticalFlip))
        ln2 = doc.Shapes.AddLine(300, 400, 50, 350, para_range(doc, 5))
        print("backwards line", info(ln2))
        attempt("its flips", lambda: (ln2.HorizontalFlip, ln2.VerticalFlip))
    finally:
        doc.Close(0)


def pictures(path):
    print("\n== pictures")
    doc = new_doc()
    try:
        direct = attempt("floating picture directly", lambda: doc.Shapes.AddPicture(
            path, False, True, 300, 500, 100, 50, para_range(doc, 3)))
        print(info(direct), counts(doc))
        at = para_range(doc, 2).Start
        inline = doc.InlineShapes.AddPicture(path, False, True, doc.Range(at, at))
        inline.AlternativeText = "A blue box"
        print("inline", counts(doc), "start", inline.Range.Start, repr(para_range(doc, 2).Text[:8]))
        floated = attempt("convert to shape", lambda: inline.ConvertToShape())
        print(info(floated), counts(doc), "alt kept", floated.AlternativeText)
        print("para 2 after", repr(para_range(doc, 2).Text[:8]))
        name = floated.Name
        for w in (5, 3, 0):
            attempt(f"wrap {w}", lambda w=w: (setattr(doc.Shapes.Item(name).WrapFormat, "Type", w),
                                              doc.Shapes.Item(name).WrapFormat.Type))
        attempt("crop", lambda: (setattr(doc.Shapes.Item(name).PictureFormat, "CropLeft", 5),
                                 doc.Shapes.Item(name).PictureFormat.CropLeft,
                                 doc.Shapes.Item(name).Width))
        back = attempt("convert to inline", lambda: doc.Shapes.Item(name).ConvertToInlineShape())
        print(counts(doc), "alt kept", back.AlternativeText if back else None,
              "start", back.Range.Start if back else None, repr(para_range(doc, 2).Text[:8]))
        attempt("wrap 7 on the floating picture", lambda: (
            setattr(direct.WrapFormat, "Type", 7), direct.WrapFormat.Type))
        print("after", counts(doc), names(doc))
        s = doc.Shapes.AddShape(RECT, 10, 10, 50, 50, para_range(doc, 4))
        attempt("autoshape to inline", lambda: s.ConvertToInlineShape())
        print(counts(doc), names(doc))
    finally:
        doc.Close(0)


def inactive_and_table():
    print("\n== a document that is not active; a table cell; a later page")
    doc = new_doc()
    other = backend.word.get_app().Documents.Add()
    try:
        app = backend.word.get_app()
        print("active", app.ActiveDocument.Name, "scratch", doc.Name)
        s = attempt("add in the inactive document", lambda: doc.Shapes.AddShape(
            RECT, 10, 10, 50, 50, para_range(doc, 3)))
        print(info(s) if s else None, "active still", app.ActiveDocument.Name, "other",
              other.Shapes.Count)
    finally:
        other.Close(0)
    try:
        table = doc.Tables.Add(para_range(doc, 6), 2, 2)
        cell = table.Cell(2, 2).Range
        s = attempt("anchored in a table cell", lambda: doc.Shapes.AddShape(
            RECT, 10, 10, 50, 30, cell))
        if s:
            print(info(s))
            attempt("layout in cell", lambda: s.LayoutInCell)
        rng = doc.Content
        rng.InsertParagraphAfter()
        last = doc.Paragraphs.Count
        doc.Paragraphs.Item(last).Range.InsertBreak(7)
        last = doc.Paragraphs.Count
        s = attempt("anchored on page 2", lambda: doc.Shapes.AddShape(
            RECT, 72, 72, 50, 30, para_range(doc, last)))
        if s:
            s.RelativeHorizontalPosition = 1
            s.RelativeVerticalPosition = 1
            s.Left, s.Top = 72, 72
            print(info(s), "page", s.Anchor.Information(3))
        attempt("header shapes", lambda: doc.Sections.Item(1).Headers.Item(1).Shapes.Count)
    finally:
        doc.Close(0)


if backend.IS_MACOS:
    print("Windows only; scripts/probes/mac_shapes.py is the Mac half.")
else:
    picture = os.path.join(tempfile.gettempdir(), "shapes-probe.png")
    with open(picture, "wb") as f:
        f.write(screen.png(60, 30, [b"\x40\x80\xc0" * 60] * 30))
    backend.word.start()
    try:
        # A part that fails answers as an error rather than raising, so each
        # answer is printed.
        for part in (anchors_and_text, position_and_order, look_and_text, inactive_and_table):
            print(part.__name__, _common.run(part))
        print("pictures", _common.run(pictures, picture))
    finally:
        backend.word.stop()
        os.remove(picture)
