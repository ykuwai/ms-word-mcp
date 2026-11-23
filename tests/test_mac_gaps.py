"""What Word for Mac's scripting lacks, through the VBA bridge (tools/_vba.py)."""

from types import SimpleNamespace

import pytest

from ms_word_mcp.backend import bridge
from ms_word_mcp.tools import _props as P
from ms_word_mcp.tools import _vba, mac, shapes, tables


class _Model:
    """Word's object model as the bridge's Member request walks it: objects are
    dicts of members, a method is a callable, and every request is noted."""

    def __init__(self, root: dict):
        self.root, self.requests = root, []

    def _arg(self, code):
        kind, rest = code[0], code[1:]
        if kind == "n":
            return float(rest) if "." in rest or "e" in rest else int(rest)
        if kind == "s":
            return bridge.decode(rest)
        if kind == "b":
            return rest == "1"
        return self._resolve(rest)

    def _resolve(self, path):
        here = self.root
        for step in filter(None, path.split("/")):
            if "(" in step:
                name, inside = step[:-1].split("(", 1)
                here = here[name](*[self._arg(c) for c in inside.split(";")])
            else:
                here = here[step]
        return here

    @staticmethod
    def _typed(value):
        if isinstance(value, dict):
            return "o"
        if value is None:
            return "e"
        if isinstance(value, bool):
            return "b1" if value else "b0"
        if isinstance(value, (int, float)):
            return f"n{value}"
        return "s" + bridge.encode(value)

    def call(self, app, what, *fields):
        assert what == "Member"
        self.requests.append(fields[1:])
        out = []
        for line in fields[1:]:
            how, path, name, *args = line.split("|")
            target, args = self._resolve(path), [self._arg(a) for a in args]
            if how == "let":
                target[name] = args[0]
                out.append("e")
            elif how == "call":
                out.append(self._typed(target[name](*args)))
            else:
                out.append(self._typed(target[name]))
        return out


@pytest.fixture
def word(monkeypatch):
    """A bridge answering from a model, and a way to switch it off."""
    state = {"why": None}

    def make(root):
        model = _Model(root)
        vba = SimpleNamespace(call=lambda app, what, *f: [] if what == "Ping" else model.call(app, what, *f))
        monkeypatch.setattr(_vba.backend, "word", SimpleNamespace(vba=vba, get_app=lambda: None))
        monkeypatch.setattr(_vba, "usable", lambda: state["why"])
        monkeypatch.setattr(mac, "_full", lambda d: "/d/Report.docx")
        return model
    make.off = lambda why="macros are off": state.update(why=why)
    return make


# -- the wire ----------------------------------------------------------------------

def test_arguments_and_paths_as_the_bridge_reads_them():
    assert [_vba.code(v) for v in (True, 3, 2.5, "Hi")] == ["b1", "n3", "n2.5", "s72,105"]
    assert _vba.path_of(("Range", 4, 9), "Font") == "Range(n4;n9)/Font"
    assert _vba.path_of("Styles", ("Item", "見出し")) == "Styles/Item(s" + bridge.encode("見出し") + ")"
    assert _vba.code(_vba.Remote("d", "Tables/Item(n1)")) == "pTables/Item(n1)"


def test_a_remote_object_reads_sets_and_calls(word):
    model = word({"Range": lambda a, b: {"Font": {"SizeBi": 11}}, "Stops": {"Add": lambda *a: a[0] * 2}})
    font = _vba.Remote("/d/Report.docx", _vba.path_of(("Range", 0, 4), "Font"))
    assert font.SizeBi == 11
    font.SizeBi = 14.5
    assert model.requests[-1] == ("let|Range(n0;n4)/Font|SizeBi|n14.5",)
    assert _vba.Remote("/d/Report.docx", "Stops").call("Add", 21) == 42
    assert isinstance(_vba.Remote("/d/Report.docx").at("Stops"), _vba.Remote)


# -- fonts and paragraphs ----------------------------------------------------------

def test_what_the_scripting_lacks_is_set_and_read_back_through_the_bridge(word):
    font = {"SizeBi": 10, "UnderlineColor": -16777216}
    word({"Range": lambda a, b: {"Font": font}})
    spec = {"size": 12, "size_complex_script": 14, "underline_color": "#FF0000"}
    refused = {"size_complex_script": "no such property", "underline_color": "no such property"}
    got = {"size": 12}
    mac._bridged(None, "Range(n0;n4)/Font", P.FONT_PROPS, spec, refused, got)
    assert refused == {}
    assert font == {"SizeBi": 14, "UnderlineColor": P.bgr("#FF0000")}
    assert got == {"size": 12, "size_complex_script": 14, "underline_color": "#FF0000"}


def test_without_the_bridge_they_stay_refused_and_say_why(word):
    word({})
    word.off("macros are off")
    refused = {"direction": "no such property"}
    mac._bridged(None, "x", P.PARAGRAPH_PROPS, {"direction": "right_to_left"}, refused, {})
    assert "cannot be used with this Word (macros are off)" in refused["direction"]


def test_tab_stops_are_replaced_and_read_back(word):
    stops = []

    def add(position, alignment, leader):
        stops.append({"CustomTab": True, "Position": position, "Alignment": alignment,
                      "Leader": leader})
    tabs = {"ClearAll": lambda: stops.clear(), "Add": add,
            "Item": lambda i: stops[i - 1] if i <= len(stops) else
            {"CustomTab": False, "Position": 108.0, "Alignment": 0, "Leader": 0}}
    word({"Range": lambda a, b: {"ParagraphFormat": {"TabStops": _Live(tabs, lambda: len(stops) + 1)}}})
    stops.append({"CustomTab": True, "Position": 9.0, "Alignment": 0, "Leader": 0})
    read, warnings = mac._bridged_tabs(None, "Range(n0;n9)/ParagraphFormat",
                                       [{"position": 72.0, "alignment": "right", "leader": "dots"}])
    assert warnings == []
    assert read == [{"position": 72.0, "alignment": "right", "leader": "dots"}]


class _Live(dict):
    """TabStops: its Count follows the stops; Word's default stop comes last."""

    def __init__(self, members, count):
        super().__init__(members)
        self._count = count

    def __getitem__(self, name):
        return self._count() if name == "Count" else super().__getitem__(name)


def test_without_the_bridge_tab_stops_say_why(word):
    word({})
    word.off("this server has no bridge")
    read, warnings = mac._bridged_tabs(None, "x", [{"position": 1, "alignment": "left",
                                                    "leader": "none"}])
    assert read is None and "cannot be used with this Word (this server has no bridge)" in warnings[0]


# -- tables and shapes -------------------------------------------------------------

def test_on_the_mac_a_column_is_added_before_the_one_after_it(word, monkeypatch):
    added = []
    columns = {"Item": lambda i: {"n": i}, "Add": lambda before=None: added.append(before)}
    word({"Tables": {"Item": lambda n: {"Columns": columns}}})
    size = {"cols": 2}
    tb = SimpleNamespace(text_object=SimpleNamespace(start_of_content=SimpleNamespace(get=lambda: 0),
                                                     end_of_content=SimpleNamespace(get=lambda: 9)))
    written = []
    monkeypatch.setattr(tables, "_mac_table", lambda d, n: tb)
    monkeypatch.setattr(tables, "_mac_size", lambda t: (2, size["cols"] + len(added)))
    monkeypatch.setattr(tables, "_mac_cell", lambda t, r, c: SimpleNamespace(text_object=SimpleNamespace(
        content=SimpleNamespace(set=lambda v: written.append((r, c, v))))))
    monkeypatch.setattr(tables, "_mac_grid", lambda t: [])
    monkeypatch.setattr(tables, "_mac_span", lambda d, t: {"first": 1, "last": 6})
    monkeypatch.setattr(mac, "_scroll", lambda *a: None)
    monkeypatch.setattr(tables.backend, "word", SimpleNamespace(document=lambda: None,
                                                                vba=_vba.backend.word.vba,
                                                                get_app=lambda: None))
    out = tables._mac_insert_columns(1, 1, None, [["x", "y"]])
    assert added == [{"n": 2}]  # before column 2: after column 1
    assert written == [(1, 2, "x"), (2, 2, "y")]
    assert (out["inserted_columns"], out["columns"]) == ([2, 2], 3)


def test_without_the_bridge_a_column_is_refused_saying_why(word, monkeypatch):
    word({})
    word.off("macros are off")
    monkeypatch.setattr(tables, "_mac_table", lambda d, n: None)
    monkeypatch.setattr(tables.backend, "word", SimpleNamespace(document=lambda: None))
    with pytest.raises(ValueError, match=r"cannot be used with this Word \(macros are off\)"):
        tables._mac_insert_columns(1, None, 1, None)


def test_on_the_mac_alt_text_and_the_vertical_anchor_go_through_the_bridge(word, monkeypatch):
    shape = {"AlternativeText": "", "TextFrame": {"VerticalAnchor": 1}}
    word({"Shapes": {"Item": lambda name: shape}})
    monkeypatch.setattr(shapes.backend, "word", SimpleNamespace(
        document=lambda: None, vba=_vba.backend.word.vba, get_app=lambda: None))
    monkeypatch.setattr(shapes, "_mac_where", lambda doc, name: None)  # in the body
    s = SimpleNamespace(name=SimpleNamespace(get=lambda: "Box 1"))
    remote, why = shapes._mac_shape_remote(s)
    remote.AlternativeText = "A blue box"
    frame, _ = shapes._mac_shape_remote(s, "TextFrame")
    frame.VerticalAnchor = shapes.VERTICAL_ANCHORS["middle"]
    assert why is None and shape == {"AlternativeText": "A blue box", "TextFrame": {"VerticalAnchor": 3}}


def test_without_the_bridge_a_shape_says_why(word):
    word({})
    word.off("macros are off")
    s = SimpleNamespace(name=SimpleNamespace(get=lambda: "Box 1"))
    assert shapes._mac_shape_remote(s) == (None, "macros are off")
    assert "(macros are off)" in shapes.NO_ALT_ON_MAC.format(why="macros are off")


def test_on_the_mac_columns_of_other_widths_are_deleted_through_the_bridge(word, monkeypatch):
    deleted = []
    columns = {"Item": lambda i: {"Delete": lambda: deleted.append(i)}}
    word({"Tables": {"Item": lambda n: {"Columns": columns}}})
    tb = SimpleNamespace(uniform=SimpleNamespace(get=lambda: False))
    monkeypatch.setattr(tables, "_mac_table", lambda d, n: tb)
    monkeypatch.setattr(tables, "_mac_size", lambda t: (2, 4 - len(deleted)))
    monkeypatch.setattr(tables, "_mac_grid", lambda t: [])
    monkeypatch.setattr(tables, "_mac_span", lambda d, t: {})
    monkeypatch.setattr(tables.backend, "word", SimpleNamespace(document=lambda: None,
                                                                vba=_vba.backend.word.vba,
                                                                get_app=lambda: None))
    out = tables._mac_delete_columns(1, 2, 3)
    assert deleted == [3, 2] and out["columns"] == 2


def test_without_the_bridge_columns_of_a_table_not_uniform_are_refused(word, monkeypatch):
    word({})
    word.off()
    tb = SimpleNamespace(uniform=SimpleNamespace(get=lambda: False))
    monkeypatch.setattr(tables, "_mac_table", lambda d, n: tb)
    monkeypatch.setattr(tables, "_mac_size", lambda t: (2, 4))
    monkeypatch.setattr(tables.backend, "word", SimpleNamespace(document=lambda: None))
    with pytest.raises(ValueError, match="merged cells"):
        tables._mac_delete_columns(1, 2, 3)


# -- text that overflows a shape on the Mac -----------------------------------------

class _Frame:
    """A shape whose text overflows by `need` points; the last paragraph's space
    after counts, as Word for Mac counts it."""

    def __init__(self, need, space):
        self.need, self.log = need, []
        self.space = space
        frame = self

        class Format:
            space_after = SimpleNamespace(get=lambda: frame.space,
                                          set=lambda v: (frame.log.append(v), setattr(frame, "space", v)))
        self.text_frame = SimpleNamespace(
            overflowing=SimpleNamespace(get=lambda: frame.need + frame.space > 0),
            text_range=SimpleNamespace(count=lambda each: 2,
                                       paragraphs={2: SimpleNamespace(paragraph_format=Format)}))


@pytest.fixture
def keys(monkeypatch):
    import sys
    monkeypatch.setitem(sys.modules, "appscript", SimpleNamespace(k=SimpleNamespace(paragraph="paragraph")))


def test_text_overflowing_only_by_its_last_space_after_is_not_cut_off(keys):
    s = _Frame(need=-3, space=8.0)  # fits with 3 pt to spare, but for the 8 pt after it
    assert shapes._mac_cut_off(s) is False
    assert s.space == 8.0 and s.log == [0, 8.0]  # put back


def test_text_that_is_cut_off_stays_cut_off(keys):
    s = _Frame(need=12, space=8.0)
    assert shapes._mac_cut_off(s) is True and s.space == 8.0


def test_with_no_space_after_word_is_taken_at_its_word(keys):
    s = _Frame(need=5, space=0.0)
    assert shapes._mac_cut_off(s) is True and s.log == []


def test_on_the_mac_a_header_shape_is_reached_through_its_header(word, monkeypatch):
    shape = {"AlternativeText": ""}
    header = {"Shapes": {"Item": lambda name: shape}}
    word({"Sections": {"Item": lambda n: {"Headers": {"Item": lambda i: header}}}})
    monkeypatch.setattr(shapes.backend, "word", SimpleNamespace(
        document=lambda: None, vba=_vba.backend.word.vba, get_app=lambda: None))
    monkeypatch.setattr(shapes, "_mac_where", lambda doc, name: ("header", "primary", 1))
    s = SimpleNamespace(name=SimpleNamespace(get=lambda: "Logo"))
    remote, why = shapes._mac_shape_remote(s)
    assert why is None and "Sections/Item(n1)/Headers/Item(n1)/Shapes" in remote._path
    remote.AlternativeText = "Logo"
    assert shape == {"AlternativeText": "Logo"}

