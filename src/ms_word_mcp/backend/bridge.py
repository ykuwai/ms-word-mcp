"""The VBA bridge: a few macros that run inside Word, for what is slow or
missing over COM and Apple Events.

Reading a paragraph over COM or Apple Events asks Word once a property, a few
milliseconds each: 200 paragraphs took 8 to 55 s (docs/PROBES.md). A macro
reads them inside Word in a fraction of a second. Word for Mac's AppleScript
also has no comment replies or done state, which VBA has.

The macros are in ``vba/ms-word-mcp-bridge.dotm``, built by Word for Windows
from ``vba/WordMcpBridge.bas`` (scripts/build_mac_bridge.py). The template is
made a global template of the Word the server talks to the first time a call
needs it, and taken out again when the server stops, so the user's Word is
left as it was; Word does not load it again by itself at its next start
(measured on both platforms).

- **Windows:** ``AddIns.Add`` loads the template from the package, and
  ``Application.Run("WordMcpRun", what, request)`` answers as its value.
- **Mac:** Word for Mac's scripting cannot add an add-in, and reads its
  Startup folder only as it starts. A template it opens makes a new document
  attached to it, whose macros can run; ``WordMcpInstall`` then makes the
  template global, and the new document is closed. The template is copied to
  Word's own folder in its container, which Word reaches without asking for
  access. ``run VB macro`` passes nothing and answers nothing, so the request
  and the answer go through files beside it.

Strings travel as their UTF-16 code units in decimal ("72,105" is "Hi"):
VBA's file I/O on the Mac does not keep Unicode. When the bridge cannot be
used (macros turned off by policy, ``WORD_VBA_BRIDGE=off``, a template that
will not load), `call` raises Unavailable once and the tools read as before.
"""

from __future__ import annotations

import logging
import os
import shutil
import threading
from contextlib import contextmanager

logger = logging.getLogger(__name__)

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = "ms-word-mcp-bridge.dotm"
TEMPLATE = os.path.join(HERE, "vba", NAME)
VERSION = "4"  # BRIDGE_VERSION in WordMcpBridge.bas; a test keeps them together
SETTING = "WORD_VBA_BRIDGE"

#: Word for Mac's own folder, which it reaches without asking for access.
MAC_FOLDER = os.path.expanduser("~/Library/Containers/com.microsoft.Word/Data/Documents")
MAC_TEMPLATE = os.path.join(MAC_FOLDER, NAME)
MAC_REQUEST = os.path.join(MAC_FOLDER, "ms-word-mcp-bridge.request")
MAC_ANSWER = os.path.join(MAC_FOLDER, "ms-word-mcp-bridge.answer")
MAC_LOCK = os.path.join(MAC_FOLDER, "ms-word-mcp-bridge.lock")


def enabled() -> bool:
    value = os.getenv(SETTING, "").strip().lower()
    if value in ("", "default", "1", "true", "on", "yes"):
        return True
    if value in ("0", "false", "off", "no"):
        return False
    logger.warning("%s=%r is not a setting; use on or off. Using on.", SETTING, value)
    return True


def encode(text: str) -> str:
    units = (text or "").encode("utf-16-le")
    return ",".join(str(int.from_bytes(units[i:i + 2], "little")) for i in range(0, len(units), 2))


def decode(field: str) -> str:
    if not field:
        return ""
    return b"".join(int(x).to_bytes(2, "little") for x in field.split(",")).decode(
        "utf-16-le", "replace")


class BridgeError(RuntimeError):
    """The macro ran and Word refused what it asked: Word's own message."""


class Unavailable(RuntimeError):
    """The bridge cannot be used with this Word; read the way it was read before."""


def parse(answer: str) -> list[str]:
    """The lines of an answer after its "ok"."""
    lines = (answer or "").replace("\r\n", "\n").rstrip("\n").split("\n")
    if lines[0] == "error":
        raise BridgeError(decode(lines[2]) if len(lines) > 2 else "Word refused the request.")
    if lines[0] != "ok":
        raise RuntimeError(f"The bridge's answer was not understood: {lines[:2]!r}")
    return lines[1:]


class Bridge:
    """The bridge in one Word. Worker thread only, as every call to Word."""

    def __init__(self, platform: str) -> None:
        self.platform = platform  # "windows" or "mac"
        self.app_id = None  # the Word it was tried in: a new Word is tried again
        self.ready = False
        self.why_not: str | None = None
        self.installed_in = None  # the Word we made it a global template of

    def state(self) -> str:
        """For word_get_app_info."""
        if not enabled():
            return f"off ({SETTING}=off)"
        if self.why_not:
            return f"off: {self.why_not}"
        return "on" if self.ready else "not loaded yet"

    def call(self, app, what: str, *fields: str) -> list[str]:
        """Ask the bridge; the answer's lines after "ok". Raises Unavailable
        when it cannot be used, BridgeError when Word refused the request."""
        if not enabled():
            raise Unavailable(f"{SETTING}=off")
        if self.app_id != id(app):
            self.app_id, self.ready, self.why_not = id(app), False, None
        if self.why_not:
            raise Unavailable(self.why_not)
        if not self.ready:
            try:
                self._load(app)
                self.ready = True
            except Exception as exc:  # noqa: BLE001 - any failure: read as before
                self.why_not = str(exc) or type(exc).__name__
                logger.warning("The VBA bridge cannot be used; reading as before: %s", self.why_not)
                raise Unavailable(self.why_not) from exc
        return parse(self._run(app, what, list(fields)))

    # -- loading -------------------------------------------------------------------

    def _load(self, app) -> None:
        try:
            found = parse(self._run(app, "Ping", []))
        except Exception:  # noqa: BLE001 - not there yet
            found = None
        # Loaded already, by a server that stopped without taking it out: this
        # one takes it out when it stops. Another version is replaced.
        self.installed_in = app if found is not None else None
        if found == [VERSION]:
            return
        if found is not None:
            self._take_out(app)
        if self.platform == "mac":
            self._mac_install(app)
        else:
            app.AddIns.Add(FileName=TEMPLATE, Install=True)
        self.installed_in = app
        version = parse(self._run(app, "Ping", []))
        if version != [VERSION]:
            raise RuntimeError(f"Word answered version {version!r} of the bridge, not {VERSION}.")

    def _mac_install(self, app) -> None:
        from appscript import k

        if not os.path.isdir(MAC_FOLDER):
            raise RuntimeError(f"Word's folder {MAC_FOLDER} is not there.")
        shutil.copyfile(TEMPLATE, MAC_TEMPLATE)
        before = _mac_documents(app)
        active = before[0] if before else None
        was = app.automation_security.get()
        # Ours, and opened to run its macro: macros on for this one file only.
        app.automation_security.set(k.msoAutomationSecurityLow)
        try:
            app.open(MAC_TEMPLATE)
        finally:
            app.automation_security.set(was)
        made = [d for d in _mac_documents(app) if d not in before]
        try:
            if not made:
                raise RuntimeError("Word opened the bridge's template but made no document from it.")
            with _mac_locked():
                parse(_mac_send(app, "WordMcpInstall", []))
        finally:
            for name in made:
                _mac_close(app, name)
            if active is not None:
                try:
                    app.documents[active].activate_object()
                except Exception:  # noqa: BLE001 - the user's document is there either way
                    pass

    def remove(self) -> None:
        """Take the template out of the Word it was added to. On the way out of
        the server: never raises."""
        app, self.installed_in, self.ready, self.app_id = self.installed_in, None, False, None
        if app is None:
            return
        try:
            self._take_out(app)
        except Exception:  # noqa: BLE001 - Word may have quit already
            logger.debug("Could not take the VBA bridge out of Word", exc_info=True)
        if self.platform == "mac":
            for path in (MAC_REQUEST, MAC_ANSWER):
                try:
                    os.remove(path)
                except OSError:
                    pass

    def _take_out(self, app) -> None:
        """Delete the template from Word's add-ins, which unloads it."""
        if self.platform == "mac":
            from appscript import k

            for i in range(int(app.count(each=k.add_in)), 0, -1):
                if app.add_ins[i].name.get() == NAME:
                    app.delete(app.add_ins[i])
        else:
            addins = app.AddIns
            for i in range(addins.Count, 0, -1):
                if addins.Item(i).Name.lower() == NAME:
                    addins.Item(i).Delete()

    # -- one request -----------------------------------------------------------------

    def _run(self, app, what: str, fields: list[str]) -> str:
        if self.platform == "mac":
            with _mac_locked():
                return _mac_send(app, "WordMcpCall", [what, *fields])
        return str(app.Run("WordMcpRun", what, "\n".join(fields)))


def _mac_documents(app) -> list[str]:
    from appscript import k

    return [app.documents[i].name.get() for i in range(1, int(app.count(each=k.document)) + 1)]


def _mac_close(app, name: str) -> None:
    from appscript import k

    try:
        app.documents[name].close(saving=k.no)
    except Exception:  # noqa: BLE001 - reported by the install that follows
        logger.warning("Could not close %s, made from the bridge's template", name)


_thread_lock = threading.Lock()


@contextmanager
def _mac_locked():
    """One request at a time, among threads and among servers (two conversations)."""
    import fcntl  # here: the module is imported on Windows too

    with _thread_lock, open(MAC_LOCK, "a") as handle:
        fcntl.flock(handle, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(handle, fcntl.LOCK_UN)


def _mac_send(app, macro: str, lines: list[str]) -> str:
    with open(MAC_REQUEST, "w", encoding="ascii") as f:
        f.write("\n".join(lines) + "\n")
    try:
        os.remove(MAC_ANSWER)
    except FileNotFoundError:
        pass
    try:
        app.run_VB_macro(macro_name=macro)
        try:
            with open(MAC_ANSWER, encoding="ascii", errors="replace") as f:
                return f.read()
        except FileNotFoundError:
            raise RuntimeError(f"Word ran {macro} but it wrote no answer.") from None
    finally:
        for path in (MAC_REQUEST, MAC_ANSWER):
            try:
                os.remove(path)
            except FileNotFoundError:
                pass
