"""word_unlink_fields: fields turned into their text, chosen by type or code."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import fields


def test_which_fields_must_be_said():
    with pytest.raises(ValidationError, match="kind"):
        fields.UnlinkFieldsInput()
    assert fields.UnlinkFieldsInput(kind="QUOTE").notes is True
    assert fields.UnlinkFieldsInput(contains="CSL_CITATION", notes=False).contains == "CSL_CITATION"


def test_chosen_by_type_and_code_from_the_end_back(monkeypatch):
    found = [{"story": 1, "begin": 5, "start": 20, "end": 23, "index": 1, "code": " QUOTE \"one\" ", "result": "one"},
             {"story": 1, "begin": 40, "start": 50, "end": 53, "index": 2, "code": " PAGE ", "result": "1"},
             {"story": 1, "begin": 60, "start": 75, "end": 78, "index": 3, "code": "QUOTE \"two\"", "result": "two"},
             {"story": 2, "begin": 2, "start": 15, "end": 20, "index": 0, "code": " QUOTE \"note\" ", "result": "note"}]
    monkeypatch.setattr(fields, "field_list", lambda has="": [f for f in found if has in f["code"]])
    order = []
    monkeypatch.setattr(fields, "unlink", lambda story, begin: order.append((story, begin)))
    out = fields._unlink_fields_impl("quote", None, False)
    assert order == [(1, 60), (1, 5)] and out["unlinked"] == 2
    assert out["fields"] == [{"in": "main text", "kind": "QUOTE", "text": "one", "index": 1},
                             {"in": "main text", "kind": "QUOTE", "text": "two", "index": 3}]
    order.clear()
    out = fields._unlink_fields_impl(None, "note", True)
    assert order == [(2, 2)] and out["fields"] == [{"in": "footnotes", "kind": "QUOTE", "text": "note"}]
    with pytest.raises(ValueError, match="No field"):
        fields._unlink_fields_impl("TOC", None, True)
