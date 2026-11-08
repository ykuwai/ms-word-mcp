"""Measure tables on Word for Windows, in a scratch document closed unsaved.

    uv run python scripts/probes/windows_tables.py

The same questions as mac_tables.py: what style a new table gets, whether an
English table style name is taken, what is left around a table made in an
empty paragraph, and how rows, columns, merging and shading behave.
"""

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def attempt(label, f):
    try:
        print(f"{label} -> {f()!r}")
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {str(e)[:200]}")


def texts(doc):
    return [doc.Paragraphs(i).Range.Text for i in range(1, doc.Paragraphs.Count + 1)]


doc = word.Documents.Add()
try:
    doc.Content.Text = "Before\r\rAfter"
    attempt("default table style", lambda: str(doc.DefaultTableStyle))
    p2 = doc.Paragraphs(2).Range
    at = doc.Range(p2.Start, p2.Start)
    tb = doc.Tables.Add(at, 2, 3)
    attempt("style of new table (Word8 behavior)", lambda: tb.Style.NameLocal)
    attempt("borders enabled", lambda: tb.Borders.Enable)
    attempt("paragraphs", lambda: texts(doc))
    for r in (1, 2):
        for c in (1, 2, 3):
            tb.Cell(r, c).Range.Text = f"r{r}c{c}"
    attempt("paragraphs after fill", lambda: texts(doc))
    attempt("table range", lambda: (tb.Range.Start, tb.Range.End))
    n = doc.Paragraphs.Count
    attempt("paragraph after table", lambda: doc.Paragraphs(n - 1).Range.Text)
    attempt("delete empty paragraph after table",
            lambda: doc.Paragraphs(n - 1).Range.Delete())
    attempt("paragraphs after delete", lambda: texts(doc))
    attempt("style 'Table Grid' by English name", lambda: (
        setattr(tb, "Style", "Table Grid"), tb.Style.NameLocal))
    attempt("Styles.Item('Table Grid')", lambda: doc.Styles("Table Grid").NameLocal)
    attempt("style by constant -161 (Light Grid)", lambda: (
        setattr(tb, "Style", -161), tb.Style.NameLocal))
    attempt("apply style heading rows", lambda: tb.ApplyStyleHeadingRows)
    attempt("rows.add at end", lambda: tb.Rows.Add().Index)
    attempt("rows.add before row 1", lambda: tb.Rows.Add(tb.Rows(1)).Index)
    attempt("columns.add at end", lambda: tb.Columns.Add().Index)
    attempt("columns.add before 1", lambda: tb.Columns.Add(tb.Columns(1)).Index)
    attempt("size", lambda: (tb.Rows.Count, tb.Columns.Count))
    attempt("delete row 1", lambda: tb.Rows(1).Delete())
    attempt("delete column 1", lambda: tb.Columns(1).Delete())
    attempt("size", lambda: (tb.Rows.Count, tb.Columns.Count))
    attempt("merge 1,1 with 1,2", lambda: tb.Cell(1, 1).Merge(tb.Cell(1, 2)))
    attempt("uniform", lambda: tb.Uniform)
    attempt("cell 1,1 text", lambda: tb.Cell(1, 1).Range.Text)
    attempt("cell 1,2 text", lambda: tb.Cell(1, 2).Range.Text)
    attempt("row 1 cells", lambda: tb.Rows(1).Cells.Count)
    attempt("columns(1) after merge", lambda: tb.Columns(1).Width)
    c = tb.Cell(2, 1)
    attempt("shading", lambda: (setattr(c.Shading, "BackgroundPatternColor", 0x0000FF),
                                c.Shading.BackgroundPatternColor))
    attempt("vertical alignment", lambda: (setattr(c, "VerticalAlignment", 1),
                                           c.VerticalAlignment))
    attempt("autofit window", lambda: (tb.AutoFitBehavior(2), tb.PreferredWidthType,
                                       tb.PreferredWidth))
    attempt("rows alignment", lambda: (setattr(tb.Rows, "Alignment", 1), tb.Rows.Alignment))
    attempt("heading format", lambda: (setattr(tb.Rows(1), "HeadingFormat", True),
                                       tb.Rows(1).HeadingFormat))
    attempt("enable borders", lambda: (setattr(tb.Borders, "Enable", True), tb.Borders.Enable))
    attempt("paragraphs", lambda: texts(doc))
    attempt("delete table", lambda: tb.Delete())
    attempt("paragraphs after table delete", lambda: texts(doc))
    # A table at the very end, from an empty last paragraph.
    end = doc.Content.End
    doc.Range(end - 1, end - 1).InsertAfter("\r")
    last = doc.Paragraphs(doc.Paragraphs.Count).Range
    t2 = doc.Tables.Add(doc.Range(last.Start, last.Start), 1, 2)
    attempt("table at end: paragraphs", lambda: texts(doc))
    # Tables.Add with a non-active document.
    other = word.Documents.Add()
    try:
        attempt("active is other", lambda: word.ActiveDocument.Name == other.Name)
        end = doc.Content.End
        doc.Range(end - 1, end - 1).InsertAfter("\r")
        last = doc.Paragraphs(doc.Paragraphs.Count).Range
        attempt("table in inactive document",
                lambda: doc.Tables.Add(doc.Range(last.Start, last.Start), 1, 2).Rows.Count)
    finally:
        other.Close(0)
finally:
    doc.Close(0)
    print("closed")
