"""Measure, on Word for Windows, what the documentation leaves unstated.

    uv run python scripts/probes/windows_checks.py

Each question comes from docs/reference/word-object-model.md ("not stated").
Works in two scratch documents it creates and closes unsaved; touches no
other document. Prints one line per measurement.
"""

from __future__ import annotations

import re
import sys

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def show(label, value):
    print(f"{label}: {value!r}")


a = word.Documents.Add()
b = None
try:
    show("unsaved Document.Path", a.Path)
    a.Content.Text = "First paragraph\rSecond paragraph\rThird paragraph"

    # DisableLineHeightGrid: does True mean "snap to grid" or "do not"?
    p1 = a.Paragraphs(1)
    for value in (True, False):
        p1.Format.DisableLineHeightGrid = value
        xml = p1.Range.WordOpenXML
        snap = re.findall(r"<w:snapToGrid[^>]*/>", xml.split("<w:body>")[1])
        show(f"DisableLineHeightGrid={value} -> reads", p1.Format.DisableLineHeightGrid)
        show(f"DisableLineHeightGrid={value} -> body XML snapToGrid", snap)

    # Does Window.ScrollIntoView move the selection?
    word.Selection.SetRange(0, 0)
    a.Windows(1).ScrollIntoView(a.Paragraphs(3).Range, True)
    show("selection after ScrollIntoView(paragraph 3)",
         (word.Selection.Start, word.Selection.End))

    # Does Range.Paragraphs count a paragraph covered only in part?
    start3 = a.Paragraphs(3).Range.Start
    show("Range(0, start of p3).Paragraphs.Count", a.Range(0, start3).Paragraphs.Count)
    show("Range(0, start of p3 + 1).Paragraphs.Count",
         a.Range(0, start3 + 1).Paragraphs.Count)
    show("Range(5, 6).Paragraphs.Count", a.Range(5, 6).Paragraphs.Count)

    # Is Chr(11) a line break inside the paragraph?
    before = a.Paragraphs.Count
    a.Paragraphs(2).Range.InsertBefore("line one\x0bline two ")
    show("paragraphs before/after inserting Chr(11)", (before, a.Paragraphs.Count))
    show("paragraph 2 text", a.Paragraphs(2).Range.Text)

    # Can the document's last paragraph mark be deleted?
    end = a.Content.End
    a.Range(end - 1, end).Delete()
    show("characters after deleting the final mark", (end, a.Content.End))

    # Which slots does Font.Name write?
    rng = a.Paragraphs(1).Range
    rng.Font.NameAscii = "Arial"
    rng.Font.NameOther = "Arial"
    rng.Font.NameFarEast = "MS Mincho"
    rng.Font.NameBi = "Arial"
    rng.Font.Name = "Times New Roman"
    show("after Font.Name = Times New Roman (ascii, other, fareast, bi)",
         (rng.Font.NameAscii, rng.Font.NameOther, rng.Font.NameFarEast, rng.Font.NameBi))
    rng.Font.Name = "Yu Gothic"
    show("after Font.Name = Yu Gothic (ascii, other, fareast, bi)",
         (rng.Font.NameAscii, rng.Font.NameOther, rng.Font.NameFarEast, rng.Font.NameBi))

    # Mixed values on a range spanning two paragraphs.
    a.Paragraphs(1).Format.SpaceBefore = 6
    a.Paragraphs(2).Format.SpaceBefore = 12
    a.Paragraphs(1).Format.Alignment = 1
    a.Paragraphs(2).Format.Alignment = 2
    a.Paragraphs(1).Range.Font.SizeBi = 10
    a.Paragraphs(2).Range.Font.SizeBi = 14
    a.Paragraphs(1).Range.Font.Color = 255
    a.Paragraphs(2).Range.Font.Color = 16711680
    mixed = a.Range(a.Paragraphs(1).Range.Start, a.Paragraphs(2).Range.End)
    show("mixed SpaceBefore", mixed.ParagraphFormat.SpaceBefore)
    show("mixed Alignment", mixed.ParagraphFormat.Alignment)
    show("mixed SizeBi", mixed.Font.SizeBi)
    show("mixed Color", mixed.Font.Color)
    show("Bold as read through pywin32", (a.Paragraphs(1).Range.Font.Bold,
                                          type(a.Paragraphs(1).Range.Font.Bold).__name__))

    # Style.Type values and Style.Linked across the document's styles.
    types = {}
    for i in range(1, a.Styles.Count + 1):
        s = a.Styles(i)
        key = (s.Type, bool(s.Linked))
        types.setdefault(key, []).append(s.NameLocal)
    for key in sorted(types):
        show(f"Style.Type, Linked = {key}", f"{len(types[key])} e.g. {types[key][:3]}")

    # Undo(times) with fewer steps than asked.
    c = word.Documents.Add()
    c.Content.Text = "x"
    show("Undo(50) on a document with one edit", c.Undo(50))
    c.Close(0)

    # An UndoRecord while another document is the active one.
    b = word.Documents.Add()  # now the active document
    show("active document is b", word.ActiveDocument.Name == b.Name)
    a.Content.Text = "Reset\rA\rB"
    record = word.UndoRecord
    record.StartCustomRecord("probe")
    a.Paragraphs(1).Range.InsertAfter(" one")
    a.Paragraphs(2).Range.InsertAfter(" two")
    a.Paragraphs(3).Range.Font.Bold = True
    record.EndCustomRecord()
    show("before undo", [a.Paragraphs(i).Range.Text for i in (1, 2, 3)])
    a.Undo(1)
    show("after one Undo of a document that is not active",
         [a.Paragraphs(i).Range.Text for i in (1, 2, 3)])
finally:
    if b is not None:
        b.Close(0)
    a.Close(0)
    print("closed both scratch documents")
