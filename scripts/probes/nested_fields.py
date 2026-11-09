"""How a field inside another field's code looks, and what updating does with it.

    uv run --no-sync python scripts/probes/nested_fields.py

Scratch documents, closed unsaved. Prints the outer field's code as Word
answers it, whether the document's fields list the inner one, whether a field
mark typed into text makes a field, and what locking the outer field does.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.tools import _common  # noqa: E402

WD_FIELD_EMPTY = -1


def windows():
    app = backend.word.get_app()
    doc = app.Documents.Add()
    try:
        doc.Content.Text = "Words here.\r"
        outer = doc.Fields.Add(doc.Range(0, 0), WD_FIELD_EMPTY, 'QUOTE "base" ', False)
        print("before nesting", repr(outer.Code.Text), "fields", doc.Fields.Count)
        at = outer.Code.End - 1
        inner = doc.Fields.Add(doc.Range(at, at), WD_FIELD_EMPTY, 'QUOTE "inner"', False)
        print("inner", repr(inner.Code.Text), repr(inner.Result.Text))
        print("document fields", doc.Fields.Count,
              [repr(doc.Fields.Item(i).Code.Text) for i in range(1, doc.Fields.Count + 1)])
        first = doc.Fields.Item(1)
        print("first: code", repr(first.Code.Text), [ord(c) for c in first.Code.Text[:6]])
        print("first: Code.Fields", first.Code.Fields.Count)
        first.Update()
        print("outer result after update", repr(first.Result.Text))

        # A name built by a field, as far as Word lets one exist.
        doc.Content.Text = "Second.\r"
        outer = doc.Fields.Add(doc.Range(0, 0), WD_FIELD_EMPTY, ' "x"', False)
        at = outer.Code.Start
        inner = doc.Fields.Add(doc.Range(at, at), WD_FIELD_EMPTY, 'QUOTE "QUOTE"', False)
        first = doc.Fields.Item(1)
        print("name from a field: code", repr(first.Code.Text),
              [ord(c) for c in first.Code.Text[:4]], "fields", doc.Fields.Count)
        print("its result", repr(first.Result.Text), first.Update(), repr(first.Result.Text))

        doc.Content.Text = "Third.\r"
        doc.Range(0, 0).Text = 'A\x13 QUOTE "made" \x15B'
        print("marks typed as text: fields", doc.Fields.Count,
              repr(doc.Paragraphs.Item(1).Range.Text))

        doc.Content.Text = "Fourth.\r"
        locked = doc.Fields.Add(doc.Range(0, 0), WD_FIELD_EMPTY, "NUMPAGES", False)
        locked.Result.Text = "SEEN"
        locked.Locked = True
        print("locked, collection update returns", doc.Fields.Update(),
              repr(doc.Fields.Item(1).Result.Text))
        return {"done": True}
    finally:
        doc.Close(0)


def mac():
    import appscript
    from appscript import k

    from ms_word_mcp.backend.mac import _decode_wide_integers

    _decode_wide_integers()
    w = appscript.app(id="com.microsoft.Word", terms="sdef")
    d = w.create_new_document()

    def fields():
        n = int(d.count(each=k.field))
        return [d.fields[i].field_code.content.get() for i in range(1, n + 1)]

    try:
        d.text_object.content.set("Words here.\r")
        from ms_word_mcp.tools.fields import _mac_new_field

        _mac_new_field(w, d, 0, 0, 'QUOTE "base" ')
        print("before nesting", fields())
        f = d.fields[1]
        code = f.field_code
        at = code.end_of_content.get() - 1
        try:
            _mac_new_field(w, d, at, at, 'QUOTE "inner"')
        except Exception as e:  # noqa: BLE001
            print("nesting failed", str(e)[:200])
        got = fields()
        print("document fields", len(got), [repr(c) for c in got])
        first = got[0]
        print("first: code", repr(first), [ord(c) for c in first[:6]])
        try:
            print("first: fields in its code", int(d.fields[1].field_code.count(each=k.field)))
        except Exception as e:  # noqa: BLE001
            print("count failed", str(e)[:200])
        d.fields[1].update_field()
        print("outer result", repr(d.fields[1].result_range.content.get()))

        d.text_object.content.set("Third.\r")
        d.create_range(start=0, end_=0).content.set('A\x13 QUOTE "made" \x15B')
        print("marks typed as text: fields", int(d.count(each=k.field)),
              repr(d.paragraphs[1].text_object.content.get()))
        return {"done": True}
    finally:
        d.close(saving=k.no)


if backend.IS_MACOS:
    print(mac())
else:
    backend.word.start()
    try:
        print(_common.run(windows))
    finally:
        backend.word.stop()
