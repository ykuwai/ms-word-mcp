"""Measure, on Word for Windows, what docs/reference/equations.md leaves unstated.

    uv run python scripts/probes/windows_equations.py

- which linear format BuildUp reads (UnicodeMath, and LaTeX control words),
- what Range.Text and Linearize give back for a built-up equation,
- whether "#(1)" built up from code numbers the equation,
- whether a SEQ field numbers equations in a three-cell table,
- whether Range.InsertXML takes OMML made from MathML with Word's own XSL.

Works in one scratch document and closes it unsaved. Prints one line per
measurement.
"""

from __future__ import annotations

import glob
import os

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def show(label, value):
    print(f"{label}: {value!r}")


def end_range(doc):
    end = doc.Content.End - 1
    return doc.Range(end, end)


def new_paragraph(doc):
    doc.Content.InsertParagraphAfter()
    return doc.Paragraphs(doc.Paragraphs.Count).Range


def build(doc, linear):
    """A new paragraph holding `linear`, made a math zone and built up."""
    rng = new_paragraph(doc)
    rng.Text = linear
    rng = doc.Paragraphs(doc.Paragraphs.Count).Range
    zone = doc.OMaths.Add(doc.Range(rng.Start, rng.End - 1))
    math = zone.OMaths(1)
    math.BuildUp()
    return math


doc = word.Documents.Add()
try:
    for label, linear in (
            ("unicodemath fraction", "a/b"),
            ("unicodemath power and root", "E=mc^2+√(x^2+1)"),
            ("unicodemath integral", "∫_0^∞ e^(-x^2) dx=√π/2"),
            ("tex words in unicodemath mode", "\\alpha+\\beta=\\gamma"),
            ("latex \\frac", "\\frac{a}{b}"),
            ("latex \\sqrt", "\\sqrt{x}"),
            ("numbered with #", "E=mc^2#(1)")):
        try:
            math = build(doc, linear)
            show(f"[{label}] Range.Text", math.Range.Text)
            show(f"[{label}] Type (0 display, 1 inline)", math.Type)
            show(f"[{label}] Functions.Count", math.Functions.Count)
            math.Linearize()
            show(f"[{label}] after Linearize, Range.Text", math.Range.Text)
            math.BuildUp()
        except Exception as exc:  # noqa: BLE001
            show(f"[{label}] raised", str(exc))

    # A numbered display equation in a three-cell table, numbered by SEQ.
    try:
        rng = new_paragraph(doc)
        table = doc.Tables.Add(rng, 1, 3)
        cell = table.Cell(1, 2).Range
        cell.Text = "a^2+b^2=c^2"
        cell = table.Cell(1, 2).Range
        zone = doc.OMaths.Add(doc.Range(cell.Start, cell.End - 1))
        zone.OMaths(1).BuildUp()
        number = table.Cell(1, 3).Range
        number.Text = "()"
        inner = doc.Range(table.Cell(1, 3).Range.Start + 1, table.Cell(1, 3).Range.Start + 1)
        field = doc.Fields.Add(inner, 12, "Equation", False)  # wdFieldSequence
        show("[table] field code", field.Code.Text)
        show("[table] field result", field.Result.Text)
        show("[table] number cell", table.Cell(1, 3).Range.Text)
        show("[table] equation type", doc.OMaths(doc.OMaths.Count).Type)
        # A second one, to see the count go on.
        rng = new_paragraph(doc)
        table2 = doc.Tables.Add(rng, 1, 3)
        number2 = table2.Cell(1, 3).Range
        number2.Text = "()"
        inner2 = doc.Range(number2.Start + 1, number2.Start + 1)
        field2 = doc.Fields.Add(inner2, 12, "Equation", False)
        show("[table] second number", table2.Cell(1, 3).Range.Text)
    except Exception as exc:  # noqa: BLE001
        show("[table] raised", str(exc))

    # MathML -> OMML with Word's own XSL, then InsertXML.
    folders = [os.path.dirname(word.Path + "\\WINWORD.EXE"), word.Path]
    found = []
    for folder in folders:
        found += glob.glob(os.path.join(folder, "*mml2omml*.xsl"), recursive=False)
        found += glob.glob(os.path.join(folder, "*MML2OMML*.XSL"), recursive=False)
    show("Word.Path", word.Path)
    show("MML2OMML found", sorted(set(found)))
    mathml = ('<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>a</mi>'
              '<mi>b</mi></mfrac></math>')
    if found:
        try:
            import subprocess  # noqa: F401 - only to note there is no XSLT here
            rng = new_paragraph(doc)
            rng = doc.Paragraphs(doc.Paragraphs.Count).Range
            before = doc.OMaths.Count
            doc.Range(rng.Start, rng.Start).InsertXML(mathml, found[0])
            show("[InsertXML with Transform=path] OMaths before/after",
                 (before, doc.OMaths.Count))
        except Exception as exc:  # noqa: BLE001
            show("[InsertXML with Transform=path] raised", str(exc))
    # MathML straight into InsertXML, no transform.
    try:
        rng = new_paragraph(doc)
        rng = doc.Paragraphs(doc.Paragraphs.Count).Range
        before = doc.OMaths.Count
        doc.Range(rng.Start, rng.Start).InsertXML(mathml)
        show("[InsertXML MathML, no transform] OMaths before/after",
             (before, doc.OMaths.Count))
    except Exception as exc:  # noqa: BLE001
        show("[InsertXML MathML, no transform] raised", str(exc))
    show("paragraph texts", [doc.Paragraphs(i).Range.Text for i in
                             range(1, doc.Paragraphs.Count + 1)][:20])
finally:
    doc.Close(0)
    print("closed the scratch document")
