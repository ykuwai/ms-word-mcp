"""Build ms-word-mcp-bridge.dotm, the VBA bridge (backend/bridge.py), on Windows.

    uv run python scripts/build_vba_bridge.py [upload_url]

Word for Mac cannot import VBA from a script, so the template both platforms
load is built by Word for Windows from src/ms_word_mcp/backend/vba/
WordMcpBridge.bas. The source's SHA-256 is written beside it
(ms-word-mcp-bridge.dotm.source), and a test checks that the template was built
from the source as it is now. Importing VBA needs "Trust
access to the VBA project object model" (Word's AccessVBOM): it is turned on
for the build only and put back exactly as it was once the Word this script
started has quit, since Word writes its security settings back on its way out.
"""

import os
import sys
import time
import urllib.request
import winreg

import pythoncom
import win32com.client

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "..", "src", "ms_word_mcp", "backend", "vba", "WordMcpBridge.bas")
OUT = os.path.join(HERE, "..", "src", "ms_word_mcp", "backend", "vba", "ms-word-mcp-bridge.dotm")
UPLOAD = sys.argv[1] if len(sys.argv) > 1 else None
KEY = r"Software\Microsoft\Office\16.0\Word\Security"


def read_setting():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY) as k:
            return winreg.QueryValueEx(k, "AccessVBOM")[0]
    except FileNotFoundError:
        return None


def write_setting(value):
    if value is None:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY, 0, winreg.KEY_SET_VALUE) as k:
                winreg.DeleteValue(k, "AccessVBOM")
        except FileNotFoundError:
            pass
        return
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, KEY) as k:
        winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, value)


before = read_setting()
print("AccessVBOM before:", before)
write_setting(1)
pythoncom.CoInitialize()
word = win32com.client.DispatchEx("Word.Application")
try:
    doc = word.Documents.Add()
    # VBA's importer wants Windows line ends; the source is kept with LF.
    import tempfile

    crlf = os.path.join(tempfile.mkdtemp(), "WordMcpBridge.bas")
    with open(SOURCE, encoding="utf-8") as src, open(crlf, "w", encoding="cp1252", newline="\r\n") as dst:
        dst.write(src.read())
    doc.VBProject.VBComponents.Import(crlf)
    names = [c.Name for c in doc.VBProject.VBComponents]
    print("modules:", names)
    if "WordMcpBridge" not in names:
        raise SystemExit("the module did not import")
    out = os.path.abspath(OUT)
    if os.path.exists(out):
        os.remove(out)
    doc.SaveAs2(out, 15)  # wdFormatXMLTemplateMacroEnabled
    doc.Close(0)
finally:
    word.Quit()
    del word
    pythoncom.CoUninitialize()
    time.sleep(3)
    write_setting(before)
    print("AccessVBOM after:", read_setting())
import hashlib

with open(SOURCE, "rb") as f:
    digest = hashlib.sha256(f.read().replace(b"\r\n", b"\n")).hexdigest()
with open(OUT + ".source", "w", encoding="ascii") as f:
    f.write(digest + "\n")
print("built", os.path.abspath(OUT), os.path.getsize(os.path.abspath(OUT)), "bytes, source", digest[:12])
if UPLOAD:
    with open(OUT, "rb") as f:
        urllib.request.urlopen(urllib.request.Request(f"{UPLOAD}/ms-word-mcp-bridge.dotm", data=f.read(),
                                                      method="PUT")).read()
    with open(OUT + ".source", "rb") as f:
        urllib.request.urlopen(urllib.request.Request(f"{UPLOAD}/ms-word-mcp-bridge.dotm.source",
                                                      data=f.read(), method="PUT")).read()
    print("uploaded")
