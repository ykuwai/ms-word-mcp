import json
from types import SimpleNamespace

import pytest

from ms_word_mcp.tools import _common, fields


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


class FakeField:
    def __init__(self, code, locked=False):
        self.Code = SimpleNamespace(Text=code)
        self.Locked = locked
        self.updated = False


class FakeFields:
    def __init__(self, *items):
        self.items = list(items)
        self.Count = len(self.items)

    def Item(self, i):
        return self.items[i - 1]

    def Update(self):
        for f in self.items:
            f.updated = not f.Locked
        return 0


@pytest.mark.parametrize("code", [
    'DDEAUTO cmd "/c calc"', "ddeauto x", '"DDE" x', '  INCLUDETEXT "C:\\a.txt"',
    'IncludePicture "http://example.com/a.png"', "LINK Word.Document a", "MACROBUTTON Run x",
    "DATABASE x", "FILLIN x", "ASK a b", "RD a.docx", '\\INCLUDE "a"'])
def test_fields_that_reach_outside_are_refused(make_doc, code):
    make_doc("One.")
    assert "are refused" in answer(fields._insert_field_impl, 1, code, "end", None)["error"]


@pytest.mark.parametrize("code", ["PAGE", "NUMPAGES", 'DATE \\@ "yyyy"', "SEQ Figure", "REF a \\h",
                                  "HYPERLINK x", "INCLUDES", "PRINTDATE"])
def test_ordinary_fields_are_not_taken_for_them(code):
    assert fields.active_kind(code) is None
    fields.check_field_code(code)


def test_a_field_cannot_be_hidden_inside_a_code():
    with pytest.raises(ValueError, match="field marks"):
        fields.check_field_code("PAGE \x13 DDEAUTO x \x15")


@pytest.mark.parametrize("code", [
    ' REF x \x13 INCLUDETEXT "a" \x15 ', 'IF 1 = 1 \x13"DDE" x\x15', '\x13 QUOTE "DDE" \x15 x'])
def test_a_field_inside_another_counts(code):
    assert fields.active_kind(code) is not None


def test_a_harmless_field_inside_another_does_not(make_doc):
    assert fields.active_kind(' QUOTE "base" \x13 QUOTE "inner" \x15 ') is None


def test_numbers_and_references_leave_a_field_hidden_inside(make_doc):
    doc = make_doc("One.")
    hidden, plain = FakeField('REF x \x13 INCLUDETEXT "a" \x15'), FakeField("REF y")
    hidden.Update = plain.Update = lambda: True
    calls = []
    plain.Update = lambda: calls.append("plain") or True
    hidden.Update = lambda: calls.append("hidden") or True
    doc.Fields = FakeFields(hidden, plain)
    out = answer(fields._update_fields_impl, "references", None, None)
    assert calls == ["plain"] and "1 field(s)" in out["warnings"][0]


def test_the_setting_allows_them(monkeypatch):
    monkeypatch.setattr(_common, "ALLOW_ACTIVE_FIELDS", True)
    fields.check_field_code('INCLUDETEXT "a"')


def test_updating_all_leaves_them_as_they_are_and_says_so(make_doc):
    doc = make_doc("One.")
    active, already, plain = FakeField(' INCLUDETEXT "a" '), FakeField("LINK b", True), FakeField("PAGE")
    doc.Fields = FakeFields(active, already, plain)
    out = answer(fields._update_fields_impl, "all", None, None)
    assert out["updated"] == 1
    assert "2 field(s)" in out["warnings"][0] and "INCLUDETEXT" in out["warnings"][0]
    assert (active.updated, already.updated, plain.updated) == (False, False, True)
    assert (active.Locked, already.Locked) == (False, True)  # as they were


def test_updating_all_updates_them_when_allowed(make_doc, monkeypatch):
    monkeypatch.setattr(_common, "ALLOW_ACTIVE_FIELDS", True)
    doc = make_doc("One.")
    active = FakeField('INCLUDETEXT "a"')
    doc.Fields = FakeFields(active)
    assert answer(fields._update_fields_impl, "all", None, None)["updated"] == 1
    assert active.updated
