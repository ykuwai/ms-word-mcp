"""Members of Word's object model reached through the VBA bridge, for what
Word for Mac's scripting lacks: a font's complex-script size and underline
color, a paragraph's reading order and tab stops, adding a style or a table
column, a shape's alternative text and vertical anchor.

`Remote(document, path)` stands for an object of the document, by its path
from the document ("Range(n0;n40)/Font", "Styles/Item(s...)"). Its
properties are read and set as attributes, so the COM helpers of
tools/_props.py work on it as they do on a COM object; a method is called
with `call`. Each read, write or call is one request to Word (backend/vba/
WordMcpBridge.bas, Member); `batch` sends several in one.
"""

from __future__ import annotations

from .. import backend
from ..backend.bridge import BridgeError, Unavailable, decode, encode

#: Why a member Word for Mac lacks was not set, when the bridge cannot be used.
NEEDS_BRIDGE = ("Word for Mac's scripting has no such property, and ms-word-mcp's VBA bridge, "
                "which sets it, cannot be used with this Word ({why}). It works on Windows.")


def code(value) -> str:
    """An argument as the bridge reads it."""
    if isinstance(value, Remote):
        return "p" + value._path
    if isinstance(value, bool):
        return "b1" if value else "b0"
    if isinstance(value, (int, float)):
        return f"n{value!r}" if isinstance(value, float) else f"n{value}"
    return "s" + encode(str(value))


def path_of(*steps) -> str:
    """A path from steps, each a member name or (name, argument, ...)."""
    out = []
    for step in steps:
        if isinstance(step, tuple):
            name, *args = step
            out.append(f"{name}({';'.join(code(a) for a in args)})")
        else:
            out.append(step)
    return "/".join(out)


def _value(answer: str, document: str, path: str, name: str):
    kind, rest = answer[:1], answer[1:]
    if kind == "n":
        number = float(rest)
        return int(number) if number.is_integer() and "." not in rest else number
    if kind == "s":
        return decode(rest)
    if kind == "b":
        return rest == "1"
    if kind == "o":
        return Remote(document, f"{path}/{name}" if path else name)
    return None


def batch(document: str, operations: list[tuple]) -> list:
    """Send operations (how, path, member, *arguments) in one request; their
    values, an object answering as a Remote. Raises Unavailable when the
    bridge cannot be used, BridgeError when Word refused one."""
    word = backend.word
    vba = getattr(word, "vba", None)
    if vba is None:
        raise Unavailable("this server has no bridge")
    lines = [f"{how}|{path}|{name}" + "".join("|" + code(a) for a in args)
             for how, path, name, *args in operations]
    answers = vba.call(word.get_app(), "Member", encode(document), *lines)
    return [_value(a, document, path, name) for a, (_, path, name, *_) in zip(answers, operations)]


class Remote:
    """An object of the document, reached through the bridge."""

    def __init__(self, document: str, path: str = "") -> None:
        object.__setattr__(self, "_document", document)
        object.__setattr__(self, "_path", path)

    def __getattr__(self, name: str):
        if name.startswith("_"):
            raise AttributeError(name)
        return batch(self._document, [("get", self._path, name)])[0]

    def __setattr__(self, name: str, value) -> None:
        batch(self._document, [("let", self._path, name, value)])

    def call(self, name: str, *args):
        return batch(self._document, [("call", self._path, name, *args)])[0]

    def at(self, *steps) -> "Remote":
        """An object further along, without asking Word."""
        more = path_of(*steps)
        return Remote(self._document, f"{self._path}/{more}" if self._path else more)


def usable() -> str | None:
    """None when the bridge works in this Word, else why not."""
    word = backend.word
    vba = getattr(word, "vba", None)
    if vba is None:
        return "this server has no bridge"
    try:
        vba.call(word.get_app(), "Ping")
    except Unavailable as exc:
        return str(exc)
    except BridgeError:
        pass
    return None


def apply_tab_stops(fmt: Remote, tab_stops) -> None:
    """tools/_props.py com_apply_tab_stops, in one request."""
    from . import _props as P

    stops = fmt.at("TabStops")
    batch(fmt._document, [("call", stops._path, "ClearAll")] + [
        ("call", stops._path, "Add", float(t["position"]), P.TAB_ALIGNMENTS[t["alignment"]],
         P.TAB_LEADERS[t["leader"]]) for t in tab_stops])


def read_tab_stops(fmt: Remote) -> list:
    """tools/_props.py com_read_tab_stops, in two requests."""
    from . import _props as P

    stops = fmt.at("TabStops")
    count = stops.Count
    members = ("CustomTab", "Position", "Alignment", "Leader")
    values = batch(fmt._document, [("get", stops.at(("Item", i))._path, m)
                                   for i in range(1, count + 1) for m in members])
    out = []
    for i in range(count):
        custom, position, alignment, leader = values[4 * i:4 * i + 4]
        if not custom:
            continue
        out.append({"position": round(position, 2),
                    "alignment": next((k for k, v in P.TAB_ALIGNMENTS.items() if v == alignment),
                                      alignment),
                    "leader": next((k for k, v in P.TAB_LEADERS.items() if v == leader), leader)})
    return out
