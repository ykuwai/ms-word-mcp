"""Where does a new style keep the formatting of where the cursor was?

    uv run python scripts/probes/windows_style_steps.py

windows_style_add.py showed that copying Normal's Font.Duplicate into a new
style cleans it, yet word_set_style still left a new style with the Title's
heading font, condensed spacing and kerning. This repeats what the tool does,
one step at a time, in a document set up as the live check sets it up, and
prints the style's font after each step. Closes the document unsaved.
"""

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def show(label, s):
    f = s.Font
    try:
        base = s.BaseStyle.NameLocal
    except Exception:  # noqa: BLE001
        base = s.BaseStyle
    print(f"{label}: base={base!r} ascii={f.NameAscii!r} spacing={f.Spacing} "
          f"kerning={f.Kerning} sizeBi={f.SizeBi} bold={f.Bold}")


doc = word.Documents.Add()
try:
    # As the live check: a Title paragraph first, the cursor left at the start.
    doc.Content.Text = "ms-word-mcp live check\rbody"
    doc.Paragraphs(1).Style = -63
    print("selection:", word.Selection.Start, word.Selection.End,
          "in", doc.Paragraphs(1).Style.NameLocal)

    for variant in ("font then paragraph", "paragraph then font", "assign twice"):
        name = f"Steps {variant}"
        s = doc.Styles.Add(name, 1)
        show(f"[{variant}] after Add", s)
        s.BaseStyle = -1
        show(f"[{variant}] after BaseStyle = -1", s)
        normal = doc.Styles(-1)
        if variant == "paragraph then font":
            s.ParagraphFormat = normal.ParagraphFormat.Duplicate
            show(f"[{variant}] after ParagraphFormat = Duplicate", s)
            s.Font = normal.Font.Duplicate
            show(f"[{variant}] after Font = Duplicate", s)
        else:
            s.Font = normal.Font.Duplicate
            show(f"[{variant}] after Font = Duplicate", s)
            s.ParagraphFormat = normal.ParagraphFormat.Duplicate
            show(f"[{variant}] after ParagraphFormat = Duplicate", s)
        if variant == "assign twice":
            s.Font = normal.Font.Duplicate
            show(f"[{variant}] after Font = Duplicate again", s)
        s.Font.NameFarEast = "BIZ UDPゴシック"
        s.Font.Size = 10.5
        show(f"[{variant}] after setting EA font and size", s)
        print(f"[{variant}] description: {s.Description!r}")
        try:
            linked = s.LinkStyle
            print(f"[{variant}] LinkStyle: {getattr(linked, 'NameLocal', linked)!r}")
        except Exception as exc:  # noqa: BLE001
            print(f"[{variant}] LinkStyle raised: {exc}")

    # Setting the members one by one, as an alternative.
    s = doc.Styles.Add("Steps members", 1)
    s.BaseStyle = -1
    nf = doc.Styles(-1).Font
    for member in ("NameAscii", "NameOther", "NameFarEast", "NameBi", "Size", "SizeBi",
                   "Bold", "Spacing", "Kerning"):
        setattr(s.Font, member, getattr(nf, member))
    show("[members] after copying members", s)
finally:
    doc.Close(0)
    print("closed the scratch document")
