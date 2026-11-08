"""Does Styles.Add take the formatting of where the cursor is, and how is it undone?

    uv run python scripts/probes/windows_style_add.py

In the live check, a style made with Styles.Add while the cursor sat in a
Title paragraph came out with the Title's condensed spacing, kerning and
heading font. This makes a scratch document, puts the cursor in a formatted
paragraph, adds styles, and tries ways to make a new style plain (only what
its base style gives). Closes the document unsaved.
"""

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def show(label, style):
    f, p = style.Font, style.ParagraphFormat
    print(f"{label}: font=({f.NameAscii}, {f.NameFarEast}, {f.Size}, spacing {f.Spacing}, "
          f"kerning {f.Kerning}, bold {f.Bold}) para=(before {p.SpaceBefore}, after "
          f"{p.SpaceAfter}, align {p.Alignment})")
    print(f"    description: {style.Description!r}")


doc = word.Documents.Add()
try:
    doc.Content.Text = "Title text\rBody text"
    doc.Paragraphs(1).Style = -63  # wdStyleTitle
    doc.Paragraphs(1).Range.Font.Bold = True
    doc.Activate()
    word.Selection.SetRange(2, 2)  # the cursor in the Title paragraph
    show("Normal", doc.Styles(-1))

    plain = doc.Styles.Add("Probe Plain", 1)
    show("added with the cursor in Title", plain)

    reset = doc.Styles.Add("Probe Reset", 1)
    reset.BaseStyle = -1
    try:
        reset.Font.Reset()
        reset.ParagraphFormat.Reset()
        show("after Font.Reset and ParagraphFormat.Reset", reset)
    except Exception as exc:  # noqa: BLE001
        print(f"Reset raised: {exc}")

    copied = doc.Styles.Add("Probe Copied", 1)
    copied.BaseStyle = -1
    copied.Font = doc.Styles(-1).Font.Duplicate
    copied.ParagraphFormat = doc.Styles(-1).ParagraphFormat.Duplicate
    show("after copying Normal's Font and ParagraphFormat", copied)

    word.Selection.SetRange(doc.Paragraphs(2).Range.Start, doc.Paragraphs(2).Range.Start)
    from_body = doc.Styles.Add("Probe From Body", 1)
    show("added with the cursor in a Normal paragraph", from_body)

    # A collapsed range in a Normal paragraph, without moving the selection.
    word.Selection.SetRange(2, 2)
    char = doc.Styles.Add("Probe Char", 2)
    print(f"character style: bold {char.Font.Bold}, size {char.Font.Size}, "
          f"description {char.Description!r}")
finally:
    doc.Close(0)
    print("closed the scratch document")
