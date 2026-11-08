"""Pictures, captions, references, headers, vertical text and ruby on Word for Windows.

    uv run python scripts/probes/windows_layout.py

The same questions as mac_layout.py, in a scratch document closed unsaved.
The picture is written in the temp folder and deleted.
"""

import os
import tempfile

import pythoncom
import win32com.client

from ms_word_mcp.tools.screen import png

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def attempt(label, f):
    try:
        print(f"{label} -> {f()!r}")
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {str(e)[:220]}")


picture = os.path.join(tempfile.gettempdir(), "ms-word-mcp-probe-picture.png")
with open(picture, "wb") as f:
    f.write(png(40, 20, [b"\xc0\x40\x40" * 40] * 20))

doc = word.Documents.Add()
other = word.Documents.Add()
try:
    doc.Content.Text = "Intro\r\rFigure X: A red box\rSee here.\rHeading text\rLast"
    doc.Paragraphs(5).Style = -2  # wdStyleHeading1
    p2 = doc.Paragraphs(2).Range
    attempt("picture", lambda: doc.InlineShapes.AddPicture(picture, False, True,
                                                            doc.Range(p2.Start, p2.Start)).Type)
    shape = doc.InlineShapes(1)
    attempt("size", lambda: (shape.Width, shape.Height, shape.LockAspectRatio))
    attempt("set width only", lambda: (setattr(shape, "Width", 100), shape.Width, shape.Height))
    attempt("lock then width", lambda: (setattr(shape, "LockAspectRatio", -1),
                                        setattr(shape, "Width", 80), shape.Width, shape.Height))
    attempt("alt text", lambda: (setattr(shape, "AlternativeText", "A red box"),
                                 shape.AlternativeText))
    attempt("range", lambda: (shape.Range.Start, shape.Range.End, shape.Range.Text))
    attempt("para 2 text", lambda: doc.Paragraphs(2).Range.Text)
    attempt("active is other", lambda: word.ActiveDocument.Name == other.Name)
    p3 = doc.Paragraphs(3).Range
    field = doc.Fields.Add(doc.Range(p3.Start + 7, p3.Start + 8), -1, "SEQ Figure \\* ARABIC", False)
    attempt("field", lambda: (field.Code.Text, field.Result.Text, field.Result.Start, field.Result.End))
    attempt("para 3 text", lambda: doc.Paragraphs(3).Range.Text)
    attempt("items for 'Figure'", lambda: list(doc.GetCrossReferenceItems("Figure")))
    doc.Paragraphs(3).Style = -35  # wdStyleCaption
    attempt("caption style", lambda: doc.Paragraphs(3).Style.NameLocal)
    field = doc.Fields(1)
    p3 = doc.Paragraphs(3).Range
    attempt("bookmark", lambda: doc.Bookmarks.Add("_Ref100000001",
                                                   doc.Range(p3.Start, field.Result.End + 1)).Name)
    attempt("bookmark text", lambda: doc.Bookmarks("_Ref100000001").Range.Text)
    attempt("visible bookmarks", lambda: doc.Bookmarks.Count)
    doc.Bookmarks.ShowHidden = True
    attempt("with hidden", lambda: doc.Bookmarks.Count)
    doc.Bookmarks.ShowHidden = False
    p4 = doc.Paragraphs(4).Range
    attempt("REF", lambda: doc.Fields.Add(doc.Range(p4.Start + 4, p4.Start + 8), -1,
                                          "REF _Ref100000001 \\h", False).Result.Text)
    attempt("para 4 text", lambda: doc.Paragraphs(4).Range.Text)
    sec = doc.Sections(1)
    header = sec.Headers(1)
    footer = sec.Footers(1)
    attempt("header", lambda: (setattr(header.Range, "Text", "Report title"), header.Range.Text,
                               header.Exists))
    attempt("header alignment", lambda: (setattr(header.Range.ParagraphFormat, "Alignment", 2),
                                         header.Range.ParagraphFormat.Alignment))
    attempt("footer text", lambda: setattr(footer.Range, "Text", "Page X of Y"))
    attempt("field Y", lambda: doc.Fields.Add(footer.Range.Characters(11), -1, "NUMPAGES", False).Result.Text)
    attempt("field X", lambda: doc.Fields.Add(footer.Range.Characters(6), -1, "PAGE", False).Result.Text)
    attempt("footer", lambda: (footer.Range.Text, footer.Range.Fields.Count))
    pn = footer.PageNumbers
    attempt("page numbers", lambda: (pn.StartingNumber, pn.RestartNumberingAtSection, pn.NumberStyle,
                                     pn.ShowFirstPageNumber))
    attempt("roman, restart 5", lambda: (setattr(pn, "NumberStyle", 2),
                                         setattr(pn, "RestartNumberingAtSection", True),
                                         setattr(pn, "StartingNumber", 5), pn.NumberStyle,
                                         pn.StartingNumber))
    attempt("different first", lambda: (setattr(sec.PageSetup, "DifferentFirstPageHeaderFooter", True),
                                        sec.PageSetup.DifferentFirstPageHeaderFooter))
    attempt("first header", lambda: (setattr(sec.Headers(2).Range, "Text", "First"),
                                     sec.Headers(2).Range.Text))
    attempt("odd even", lambda: (setattr(sec.PageSetup, "OddAndEvenPagesHeaderFooter", True),
                                 sec.PageSetup.OddAndEvenPagesHeaderFooter))
    attempt("main text unchanged", lambda: doc.Paragraphs(1).Range.Text)
    ps = sec.PageSetup
    attempt("page", lambda: (ps.Orientation, ps.PageWidth, ps.PageHeight, ps.TextColumns.Count))
    attempt("section orientation", lambda: sec.Range.Orientation)
    attempt("set vertical east asian", lambda: (setattr(sec.Range, "Orientation", 1), sec.Range.Orientation))
    attempt("page after", lambda: (ps.Orientation, ps.PageWidth, ps.PageHeight))
    attempt("page setup text direction", lambda: ps.SectionDirection)
    attempt("back to portrait", lambda: (setattr(ps, "Orientation", 0), ps.Orientation, ps.PageWidth,
                                         ps.PageHeight, sec.Range.Orientation))
    attempt("grid", lambda: (ps.LayoutMode, ps.CharsLine, ps.LinesPage))
    attempt("set grid", lambda: (setattr(ps, "LayoutMode", 1), setattr(ps, "CharsLine", 40),
                                 setattr(ps, "LinesPage", 30), ps.LayoutMode, ps.CharsLine, ps.LinesPage))
    attempt("set horizontal", lambda: (setattr(sec.Range, "Orientation", 0), sec.Range.Orientation))
    first = doc.Paragraphs(1).Range
    doc.Range(first.End - 1, first.End - 1).InsertAfter(" 漢字です")
    first = doc.Paragraphs(1).Range
    at = first.End - 5
    attempt("target", lambda: doc.Range(at, at + 2).Text)
    attempt("ruby", lambda: doc.Range(at, at + 2).PhoneticGuide("かんじ"))
    attempt("para 1", lambda: doc.Paragraphs(1).Range.Text)
    attempt("fields", lambda: [doc.Fields(i).Code.Text for i in range(1, doc.Fields.Count + 1)])
    attempt("paragraph count", lambda: doc.Paragraphs.Count)
    attempt("ruby range text", lambda: doc.Range(at, at + 2).Text)
    attempt("active is still other", lambda: word.ActiveDocument.Name == other.Name)
finally:
    other.Close(0)
    doc.Close(0)
    os.remove(picture)
    print("closed")
