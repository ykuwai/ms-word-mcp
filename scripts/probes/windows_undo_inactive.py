"""Does one UndoRecord group a call's edits when the document is not the active one?

    uv run python scripts/probes/windows_undo_inactive.py

The server writes to the document a call names, which is often not Word's
active document. Microsoft's UndoRecord page warns that a record ends when
code "begins to write to the second document". This makes two scratch
documents, keeps the second one active, makes three edits to the first inside
one record, and undoes once. It also runs the same with the first one active,
for comparison. Both documents are closed unsaved.
"""

import pythoncom
import win32com.client

pythoncom.CoInitialize()
word = win32com.client.Dispatch("Word.Application")
word.Visible = True


def texts(doc):
    return [doc.Paragraphs(i).Range.Text.rstrip("\r") for i in range(1, doc.Paragraphs.Count + 1)]


def run(label, target, active):
    target.Content.Text = "one\rtwo\rthree"
    target.Paragraphs(3).Range.Font.Bold = False
    active.Activate()
    print(f"[{label}] active is target: {word.ActiveDocument.Name == target.Name}")
    record = word.UndoRecord
    record.StartCustomRecord("probe")
    p1 = target.Paragraphs(1).Range
    target.Range(p1.End - 1, p1.End - 1).InsertAfter(" ONE")
    p2 = target.Paragraphs(2).Range
    target.Range(p2.End - 1, p2.End - 1).InsertAfter(" TWO")
    target.Paragraphs(3).Range.Font.Bold = True
    print(f"[{label}] recording while editing: {record.IsRecordingCustomRecord}")
    record.EndCustomRecord()
    print(f"[{label}] after edits: {texts(target)} bold3={target.Paragraphs(3).Range.Font.Bold}")
    result = target.Undo(1)
    print(f"[{label}] Undo(1) answered {result}: {texts(target)} "
          f"bold3={target.Paragraphs(3).Range.Font.Bold}")
    result = target.Undo(1)
    print(f"[{label}] second Undo(1) answered {result}: {texts(target)} "
          f"bold3={target.Paragraphs(3).Range.Font.Bold}")


a = word.Documents.Add()
b = word.Documents.Add()
try:
    run("target not active", a, b)
    run("target active", a, a)
finally:
    b.Close(0)
    a.Close(0)
    print("closed both scratch documents")
