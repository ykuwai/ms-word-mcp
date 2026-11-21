"""The page picture on the Mac of a document with comments or tracked changes:
drawn from a copy without them, since Word for Mac's PDF shrinks the page
beside a markup pane (tools/screen.py)."""

import sys
from types import SimpleNamespace

import pytest

from ms_word_mcp.tools import screen

k = SimpleNamespace(missing_value="missing", text_column="text column", section="section",
                    revision="revision", Word_comment="Word comment", document="document", no="no",
                    header_footer_primary="primary", header_footer_first_page="first",
                    header_footer_even_pages="even")


@pytest.fixture(autouse=True)
def keys(monkeypatch):
    monkeypatch.setitem(sys.modules, "appscript", SimpleNamespace(k=k))


class Prop:
    def __init__(self, log, name, value=None):
        self.log, self.name, self.value = log, name, value

    def get(self):
        return self.value

    def set(self, value):
        self.log.append((self.name, value))
        self.value = value


class Part:
    """A header or footer."""

    def __init__(self, log, label, linked=False):
        self.label = label
        self.link_to_previous = Prop(log, f"{label} linked", linked)
        self.text_object = SimpleNamespace(formatted_text=Prop(log, f"{label} text"))


class Section:
    def __init__(self, log, setup, columns=1, linked=False):
        self.page_setup = SimpleNamespace(
            **{name: Prop(log, name, value) for name, value in setup.items()},
            count=lambda each: columns,
            set_number_of_text_columns=lambda number_of_columns: log.append(("columns", number_of_columns)))
        self.text_object = SimpleNamespace(orientation=Prop(log, "orientation of text", "horizontal"))
        self.parts = {(get, kind): Part(log, f"{get} {kind}", linked)
                      for get in ("get_header", "get_footer") for kind in ("primary", "first", "even")}

    def get_header(self, index):
        return self.parts[("get_header", index)]

    def get_footer(self, index):
        return self.parts[("get_footer", index)]


SETUP = {name: f"value of {name}" for name in screen.CARRIED_SETUP}


def test_the_last_sections_setup_is_carried_with_the_grid_after_its_lines():
    log = []
    doc = SimpleNamespace(count=lambda each: 2, sections={2: Section([], SETUP, columns=2)})
    doc.sections[2].page_setup.gutter.value = "missing"  # Word answers a missing value: left as is
    copy = SimpleNamespace(count=lambda each: 2,
                           sections={2: Section(log, {n: None for n in SETUP}, columns=1)})
    screen._carry_last_section(doc, copy)
    names = [name for name, _ in log]
    assert names.index("lines_page") < names.index("chars_line") < names.index("layout_mode")
    assert names.index("orientation") < names.index("page_width")
    assert "gutter" not in names and ("columns", 2) in log


def test_a_header_of_its_own_is_copied_and_a_linked_one_stays_linked():
    log = []
    original = Section([], SETUP)
    original.parts[("get_header", "primary")].link_to_previous.value = False
    original.parts[("get_footer", "primary")].link_to_previous.value = True
    doc = SimpleNamespace(count=lambda each: 2, sections={2: original})
    copy = SimpleNamespace(count=lambda each: 2, sections={2: Section(log, {n: None for n in SETUP})})
    screen._carry_last_section(doc, copy)
    assert ("get_header primary linked", False) in log
    assert ("get_header primary text", original.parts[("get_header", "primary")].text_object) in log
    assert ("get_footer primary linked", True) in log
    assert not any(name == "get_footer primary text" for name, _ in log)


def test_a_copy_with_another_number_of_sections_is_refused():
    doc = SimpleNamespace(count=lambda each: 2)
    copy = SimpleNamespace(count=lambda each: 1)
    with pytest.raises(RuntimeError, match="another number of sections"):
        screen._carry_last_section(doc, copy)


class Comment:
    def __init__(self, log, n):
        self.log, self.n = log, n

    def delete(self):
        self.log.append(("delete comment", self.n))


def _app(log, comments=2, revisions=1, fail=False):
    """An app whose `make` gives a copy holding `comments` and `revisions`."""
    named = []

    def set_text(value):
        if fail:
            raise RuntimeError("Word refused")
        log.append(("formatted text", value))

    copy = SimpleNamespace(
        text_object=SimpleNamespace(formatted_text=SimpleNamespace(set=set_text)),
        count=lambda each: {"revision": revisions, "Word comment": comments}[each],
        accept_all_revisions=lambda: log.append(("accept all",)),
        Word_comments={n: Comment(log, n) for n in range(1, comments + 1)},
        close=lambda saving: log.append(("close copy", saving)))

    def documents(name):
        named.append(name)
        return SimpleNamespace(text_object=f"the text of {name}")

    app = SimpleNamespace(make=lambda new: log.append(("make", new)) or copy,
                          documents=SimpleNamespace(__getitem__=documents))
    app.documents = type("Docs", (), {"__getitem__": lambda self, name: documents(name)})()
    return app, copy, named


def test_the_copy_is_made_from_the_original_named_and_its_markup_taken_out(monkeypatch):
    monkeypatch.setattr(screen, "_carry_last_section", lambda doc, copy: None)
    log = []
    app, copy, named = _app(log)
    doc = SimpleNamespace(name=SimpleNamespace(get=lambda: "Report.docx"))
    assert screen._mac_without_markup(app, doc) is copy
    assert named == ["Report.docx"]  # by name: a new document goes first in Word's list
    assert log == [("make", "document"), ("formatted text", "the text of Report.docx"),
                   ("accept all",), ("delete comment", 2), ("delete comment", 1)]


def test_a_copy_that_fails_is_closed(monkeypatch):
    monkeypatch.setattr(screen, "_carry_last_section", lambda doc, copy: None)
    log = []
    app, _, _ = _app(log, fail=True)
    doc = SimpleNamespace(name=SimpleNamespace(get=lambda: "Report.docx"))
    with pytest.raises(RuntimeError, match="Word refused"):
        screen._mac_without_markup(app, doc)
    assert log[-1] == ("close copy", "no")


def test_closing_the_copy_gives_the_users_document_back():
    log = []
    app = SimpleNamespace(documents={"Mine.docx": SimpleNamespace(
        activate_object=lambda: log.append("activated Mine.docx"))})
    copy = SimpleNamespace(close=lambda saving: log.append(f"closed copy, saving {saving}"))
    screen._mac_close_copy(app, copy, "Mine.docx")
    assert log == ["closed copy, saving no", "activated Mine.docx"]
