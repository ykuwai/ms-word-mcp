"""The VBA bridge (backend/bridge.py) and the paragraphs it reads."""

import hashlib
import os
import re
import zipfile
from types import SimpleNamespace

import pytest

from ms_word_mcp.backend import bridge
from ms_word_mcp.tools import mac, text

VBA = os.path.join(os.path.dirname(bridge.TEMPLATE), "WordMcpBridge.bas")


# -- the template ------------------------------------------------------------------

def test_the_template_was_built_from_the_source_as_it_is():
    """scripts/build_vba_bridge.py writes the source's hash beside the template."""
    with open(VBA, "rb") as f:
        digest = hashlib.sha256(f.read().replace(b"\r\n", b"\n")).hexdigest()
    with open(bridge.TEMPLATE + ".source", encoding="ascii") as f:
        assert f.read().strip() == digest, "Rebuild the template: scripts/build_vba_bridge.py"


def test_the_template_holds_macros():
    with zipfile.ZipFile(bridge.TEMPLATE) as z:
        assert "word/vbaProject.bin" in z.namelist()


def test_python_and_vba_agree_on_the_version():
    with open(VBA, encoding="utf-8") as f:
        assert re.search(r'BRIDGE_VERSION As String = "(\d+)"', f.read()).group(1) == bridge.VERSION


# -- the wire ----------------------------------------------------------------------

@pytest.mark.parametrize("value", ["", "Hi", "日本語 𠮷 ✓", "a|b,c\rd"])
def test_text_travels_as_utf16_units(value):
    assert bridge.decode(bridge.encode(value)) == value
    assert "|" not in bridge.encode(value) and "\n" not in bridge.encode(value)


def test_an_answer_is_read_or_its_error_raised():
    assert bridge.parse("ok\n1\n2\n") == ["1", "2"]
    assert bridge.parse("ok\r\n7\r\n") == ["7"]
    with pytest.raises(bridge.BridgeError, match="No open document is called x"):
        bridge.parse("error\n5\n" + bridge.encode("No open document is called x"))
    with pytest.raises(RuntimeError, match="not understood"):
        bridge.parse("")


@pytest.mark.parametrize("value, on", [("", True), ("on", True), ("off", False), ("0", False),
                                       ("maybe", True)])
def test_the_setting(monkeypatch, value, on):
    monkeypatch.setenv(bridge.SETTING, value)
    assert bridge.enabled() is on


# -- loading, on Windows -----------------------------------------------------------

class _WinApp:
    """Word over COM: AddIns.Add makes the macro there; Run answers through it."""

    def __init__(self, macros_allowed=True, version="1"):
        self.loaded, self.runs, self.allowed, self.version = False, [], macros_allowed, version
        self.items = []
        app = self

        class AddIns:
            @property
            def Count(self):
                return len(app.items)

            def Item(self, i):
                return app.items[i - 1]

            def Add(self, FileName, Install):
                app.loaded = True
                item = SimpleNamespace(Name=os.path.basename(FileName),
                                       Delete=lambda: app.items.remove(item))
                app.items.append(item)
                return item
        self.AddIns = AddIns()

    def Run(self, macro, what, request):
        self.runs.append((macro, what, request))
        if not (self.loaded and self.allowed):
            raise RuntimeError("Unable to run the specified macro")
        return f"ok\n{self.version}" if what == "Ping" else "ok\n" + request.upper()


def test_on_windows_the_template_is_added_the_first_time_and_then_used():
    app, vba = _WinApp(), bridge.Bridge("windows")
    assert vba.call(app, "Echo", "a", "b") == ["A", "B"]
    assert app.items[0].Name == bridge.NAME and vba.state() == "on"
    assert vba.call(app, "Echo", "c") == ["C"]
    assert [w for _, w, _ in app.runs] == ["Ping", "Ping", "Echo", "Echo"]
    vba.remove()
    assert app.items == []


def test_macros_turned_off_are_tried_once_and_then_read_as_before():
    app, vba = _WinApp(macros_allowed=False), bridge.Bridge("windows")
    with pytest.raises(bridge.Unavailable, match="Unable to run"):
        vba.call(app, "Echo", "a")
    runs = len(app.runs)
    with pytest.raises(bridge.Unavailable):
        vba.call(app, "Echo", "a")
    assert len(app.runs) == runs and vba.state().startswith("off: ")


def test_another_version_of_the_template_is_not_used():
    with pytest.raises(bridge.Unavailable, match="version"):
        bridge.Bridge("windows").call(_WinApp(version="0"), "Echo", "a")


def _preloaded(version):
    """A Word holding the template already, left by a server that stopped badly."""
    app = _WinApp(version=version)
    app.AddIns.Add(bridge.TEMPLATE, True)
    return app


def test_a_bridge_left_loaded_is_used_and_taken_out_at_the_end():
    app, vba = _preloaded("1"), bridge.Bridge("windows")
    assert vba.call(app, "Echo", "a") == ["A"]
    assert len(app.items) == 1  # not added a second time
    vba.remove()
    assert app.items == []


def test_another_version_left_loaded_is_replaced():
    app = _preloaded("0")
    added = []
    real_add = app.AddIns.Add

    def add(FileName, Install):
        added.append(FileName)
        app.version = "1"
        return real_add(FileName, Install)
    app.AddIns.Add = add
    vba = bridge.Bridge("windows")
    assert vba.call(app, "Echo", "a") == ["A"]
    assert added == [bridge.TEMPLATE] and len(app.items) == 1


def test_a_new_word_is_tried_again():
    vba = bridge.Bridge("windows")
    with pytest.raises(bridge.Unavailable):
        vba.call(_WinApp(macros_allowed=False), "Echo", "a")
    assert vba.call(_WinApp(), "Echo", "a") == ["A"]


def test_switched_off_it_never_asks_word(monkeypatch):
    monkeypatch.setenv(bridge.SETTING, "off")
    app = _WinApp()
    with pytest.raises(bridge.Unavailable, match="off"):
        bridge.Bridge("windows").call(app, "Echo", "a")
    assert app.runs == []


# -- loading, on the Mac -----------------------------------------------------------

class _Prop:
    def __init__(self, log, name, value):
        self.log, self.name, self.value = log, name, value

    def get(self):
        return self.value

    def set(self, value):
        self.log.append((self.name, value))
        self.value = value


class _MacApp:
    """Word for Mac: opening the template makes a new document; WordMcpInstall,
    run while it is open, makes the macros global."""

    def __init__(self, log, folder):
        self.log, self.folder, self.global_ = log, folder, False
        self.names = ["Mine.docx"]
        self.automation_security = _Prop(log, "automation security", "user's")
        app = self

        class Docs:
            def __getitem__(self, key):
                name = app.names[key - 1] if isinstance(key, int) else key
                return SimpleNamespace(
                    name=_Prop(log, "name", name),
                    close=lambda saving: (log.append(("close", name, saving)), app.names.remove(name)),
                    activate_object=lambda: log.append(("activate", name)))
        self.documents = Docs()

    def count(self, each):
        return len(self.names)

    def open(self, path):
        self.log.append(("open", os.path.basename(path), self.automation_security.value))
        self.names.insert(0, "Document2")

    def run_VB_macro(self, macro_name):
        with open(bridge.MAC_REQUEST, encoding="ascii") as f:
            lines = f.read().split("\n")
        self.log.append(("run", macro_name, lines[0]))
        answer = os.path.join(self.folder, "ms-word-mcp-bridge.answer")
        if macro_name == "WordMcpInstall":
            assert "Document2" in self.names, "run while the document made from the template is open"
            self.global_ = True
            reply = "ok\n1"
        elif not self.global_:
            raise RuntimeError("-1708")
        else:
            reply = "ok\n1" if lines[0] == "Ping" else "ok\n" + ",".join(lines[1:-1])
        with open(answer, "w") as f:
            f.write(reply + "\n")


@pytest.fixture
def mac_folder(monkeypatch, tmp_path):
    for name in ("MAC_FOLDER", "MAC_TEMPLATE", "MAC_REQUEST", "MAC_ANSWER", "MAC_LOCK"):
        default = getattr(bridge, name)
        monkeypatch.setattr(bridge, name, str(tmp_path / os.path.basename(default))
                            if name != "MAC_FOLDER" else str(tmp_path))
    monkeypatch.setitem(__import__("sys").modules, "appscript", SimpleNamespace(k=SimpleNamespace(
        msoAutomationSecurityLow="low", no="no", document="document", add_in="add in")))
    return tmp_path


@pytest.mark.skipif(os.name == "nt", reason="fcntl")
def test_on_the_mac_a_running_word_is_given_the_template_through_a_document_made_from_it(mac_folder):
    log = []
    app, vba = _MacApp(log, str(mac_folder)), bridge.Bridge("mac")
    assert vba.call(app, "Echo", "a", "b") == ["a,b"]
    assert (mac_folder / bridge.NAME).read_bytes() == open(bridge.TEMPLATE, "rb").read()
    steps = [entry for entry in log if entry[0] != "name"]
    assert steps == [
        ("run", "WordMcpCall", "Ping"),  # not there yet
        ("automation security", "low"), ("open", bridge.NAME, "low"),
        ("automation security", "user's"),
        ("run", "WordMcpInstall", ""),
        ("close", "Document2", "no"), ("activate", "Mine.docx"),
        ("run", "WordMcpCall", "Ping"), ("run", "WordMcpCall", "Echo")]
    assert app.names == ["Mine.docx"]
    assert not os.path.exists(bridge.MAC_REQUEST) and not os.path.exists(bridge.MAC_ANSWER)


@pytest.mark.skipif(os.name == "nt", reason="fcntl")
def test_on_the_mac_the_users_setting_is_put_back_when_opening_fails(mac_folder):
    log = []
    app = _MacApp(log, str(mac_folder))
    app.open = lambda path: (_ for _ in ()).throw(RuntimeError("-43"))
    with pytest.raises(bridge.Unavailable, match="-43"):
        bridge.Bridge("mac").call(app, "Echo")
    assert app.automation_security.value == "user's"


# -- the paragraphs it reads --------------------------------------------------------

def _row(index, text, style="Normal", level=10, changes=0, detail=None):
    fields = [str(index), str(changes), str(level), bridge.encode(style), bridge.encode(text)]
    return "|".join(fields + (detail or []))


DETAIL = ["0", bridge.encode("Calibri ()"), bridge.encode("游明朝"), "", "10.5", "9999999",
          "-1", "0", "-16777216", "5", "13.8", "0", "8"]


def test_a_row_reads_as_word_get_paragraphs_answers_it():
    info, slow = text.paragraph_from_row(_row(3, "Heading\r", "Heading 1", 1).split("|"), False)
    assert (info, slow) == ({"index": 3, "text": "Heading", "style": "Heading 1", "outline_level": 1},
                            False)


def test_a_rows_detail_goes_through_the_same_conversions_as_com():
    info, _ = text.paragraph_from_row(_row(1, "x\r", detail=DETAIL).split("|"), True)
    assert info["alignment"] == "left" and info["line_spacing_rule"] == "multiple"
    assert info["fonts"] == {"latin": "Calibri", "east_asian": "游明朝", "complex_script": None,
                             "size": 10.5, "size_complex_script": None, "bold": True,
                             "italic": False, "color": None}
    assert (info["line_spacing_lines"], info["space_before"], info["space_after"]) == (1.15, 0.0, 8.0)


@pytest.mark.parametrize("row", [_row(1, "struck\r", changes=2), _row(1, "a\x15b\r")])
def test_tracked_changes_and_ruby_are_read_the_slow_way(row):
    assert text.paragraph_from_row(row.split("|"), False)[1] is True


class _Vba:
    def __init__(self, lines=None, error=None):
        self.lines, self.error, self.asked = lines, error, []

    def call(self, app, what, *fields):
        self.asked.append((what, *fields))
        if self.error:
            raise self.error
        return self.lines


def _word(monkeypatch, vba):
    monkeypatch.setattr(text.backend, "word", SimpleNamespace(vba=vba, get_app=lambda: None))


def test_paragraphs_are_read_in_one_call_and_the_slow_ones_get_their_text(monkeypatch):
    vba = _Vba(["3", _row(1, "One\r", detail=DETAIL), _row(2, "Struck\r", changes=1, detail=DETAIL)])
    _word(monkeypatch, vba)
    slow = []

    def read_text(i):
        slow.append(i)
        return {"index": i, "text": "", "style": "Normal", "deleted_text": "Struck", "deleted": True}
    out = text.bridge_paragraphs(lambda: "Report.docx", 1, 2, True, True, text.time.monotonic(),
                                 read_text)
    assert vba.asked[0][:6] == ("Paragraphs", bridge.encode("Report.docx"), "1", "2", "1", "1")
    assert 1.0 <= float(vba.asked[0][6]) <= text.READ_SECONDS  # the time left
    assert slow == [2]
    assert out[1]["deleted_text"] == "Struck" and out[1]["text"] == ""
    assert out[1]["fonts"]["bold"] is True  # the bridge's detail stays


@pytest.mark.parametrize("error", [bridge.Unavailable("off"), bridge.BridgeError("not found")])
def test_a_bridge_that_cannot_read_leaves_it_to_the_old_way(monkeypatch, error):
    _word(monkeypatch, _Vba(error=error))
    assert text.bridge_paragraphs(lambda: "x", 1, 2, False, False, text.time.monotonic(),
                                  None) is None


def test_with_no_bridge_nothing_is_asked(monkeypatch):
    monkeypatch.setattr(text.backend, "word", SimpleNamespace())
    assert text.bridge_paragraphs(None, 1, 2, False, False, 0.0, None) is None


def test_on_the_mac_get_paragraphs_uses_the_bridge(monkeypatch):
    vba = _Vba(["2", _row(1, "One\r"), _row(2, "Two\r")])
    doc = SimpleNamespace(count=lambda each: 2)
    monkeypatch.setattr(text.backend, "word", SimpleNamespace(vba=vba, get_app=lambda: None,
                                                              document=lambda: doc))
    monkeypatch.setattr(mac, "_k", lambda: SimpleNamespace(revision="revision", paragraph="paragraph",
                                                           table="table"))
    monkeypatch.setattr(mac, "_count", lambda d: 2)
    monkeypatch.setattr(mac, "_full", lambda d: "/d/Report.docx")
    monkeypatch.setattr(mac, "_has_revisions", lambda d: False)
    import ms_word_mcp.tools.tables as tables
    monkeypatch.setattr(tables, "mac_mark_tables", lambda d, read: None)
    out = mac.get_paragraphs(1, None, False)
    assert [p["text"] for p in out["paragraphs"]] == ["One", "Two"] and out["count"] == 2
    assert vba.asked[0][1] == bridge.encode("/d/Report.docx")


def test_app_info_says_whether_the_bridge_works():
    from ms_word_mcp.tools import documents

    assert documents.vba_bridge_state(SimpleNamespace()) == "off (this server has none)"
    vba = bridge.Bridge("windows")
    assert documents.vba_bridge_state(SimpleNamespace(vba=vba)) == "not loaded yet"
    vba.call(_WinApp(), "Echo")
    assert documents.vba_bridge_state(SimpleNamespace(vba=vba)) == "on"
