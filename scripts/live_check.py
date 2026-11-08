"""Run the tools against a real Word, in a scratch document, and report.

    uv run python scripts/live_check.py

Creates a new document saved in the temp folder, runs each tool on it, checks
what Word reads back, exports a PDF next to it, and closes the document. It
touches no other document. Each step prints PASS or FAIL with the tool's
answer, and the run ends with the counts.
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.backend import call_document  # noqa: E402
from ms_word_mcp.tools import _common, documents, edit, formatting, text  # noqa: E402

results: list[tuple[str, bool]] = []


def step(label, impl, *args, undo=None, check=lambda out: "error" not in out):
    out = json.loads(_common.run(impl, *args, undo=undo))
    try:
        ok = bool(check(out))
    except (KeyError, IndexError, TypeError):
        ok = False
    results.append((label, ok))
    print(f"{'PASS' if ok else 'FAIL'}  {label}")
    print("      " + json.dumps(out, ensure_ascii=False)[:600])
    return out


def main() -> int:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    # Word for Mac is sandboxed: it writes to its own container without asking,
    # and anywhere else only after the user grants access in a dialog.
    folder = (Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents"
              if sys.platform == "darwin" else Path(tempfile.gettempdir()))
    path = str(folder / f"ms-word-mcp-live-{stamp}.docx")
    backend.word.start()
    try:
        created = step("create and save", documents._create_document_impl, None, path,
                       check=lambda o: o.get("full_name", "").lower().endswith(".docx"))
        if "error" in created:
            return 1
        token = call_document.set(created["full_name"])
        try:
            run_steps(path)
        finally:
            step("close", documents._close_document_impl, True)
            call_document.reset(token)
    finally:
        backend.word.stop()
    passed = sum(ok for _, ok in results)
    print(f"\n{passed} of {len(results)} steps passed. Document: {path}")
    return 0 if passed == len(results) else 1


def run_steps(path: str) -> None:
    step("insert paragraphs", text._insert_paragraphs_impl, [
        {"text": "ms-word-mcp live check", "style": "Title"},
        {"text": "見出し 1 の段落", "style": "Heading 1"},
        {"text": "本文です。English text and 日本語 in one paragraph.", "style": "Normal"},
        {"text": "Second heading", "style": "Heading 2"},
        {"text": "Another body paragraph with the word cat.", "style": None},
    ], None, None, undo="Insert paragraphs",
        check=lambda o: [p["index"] for p in o.get("inserted", [])] == [1, 2, 3, 4, 5])
    step("outline", text._get_outline_impl,
         check=lambda o: [h["level"] for h in o.get("headings", [])][-2:] == [1, 2])
    step("paragraphs with detail", text._get_paragraphs_impl, 1, 5, True,
         check=lambda o: len(o.get("paragraphs", [])) == 5)
    step("three fonts and a color", formatting._format_text_impl, 3, None, None,
         {"latin_font": "Segoe UI", "east_asian_font": "BIZ UDPゴシック",
          "complex_script_font": "Arial", "size": 11, "color": "#1F4E79"},
         undo="Format text", check=lambda o: "warnings" not in o
         and o["read_back"][0]["color"] == "#1F4E79")
    step("latin-only East Asian font warns", formatting._format_text_impl, 3, None, None,
         {"east_asian_font": "Segoe UI"}, undo="Format text",
         check=lambda o: any("east_asian_font" in w for w in o.get("warnings", [])))
    step("bold a match", formatting._format_text_impl, 3, None, "English",
         {"bold": True}, undo="Format text", check=lambda o: o.get("ranges") == 1)
    step("emphasis dots, spacing, scaling", formatting._format_text_impl, 3, None, "日本語",
         {"emphasis_mark": "dot_above", "character_spacing": 1, "scaling": 90},
         undo="Format text", check=lambda o: "warnings" not in o)
    step("line spacing 1.5 lines", formatting._set_paragraph_format_impl, 3, 5,
         {"alignment": "justify", "line_spacing_rule": "multiple", "line_spacing": 1.5,
          "space_before": 6, "space_after": 6, "snap_to_grid": False},
         undo="Paragraph format",
         check=lambda o: o["paragraphs"][0].get("line_spacing_lines") == 1.5
         and "warnings" not in o)
    step("line spacing exactly 18 pt", formatting._set_paragraph_format_impl, 5, None,
         {"line_spacing_rule": "exactly", "line_spacing": 18}, undo="Paragraph format",
         check=lambda o: o["paragraphs"][0].get("line_spacing_pt") == 18)
    step("Japanese spacing in lines and characters", formatting._set_paragraph_format_impl,
         3, None, {"space_before_lines": 0.5, "first_line_indent_chars": 1,
                   "east_asian_line_breaking": True, "hanging_punctuation": True},
         undo="Paragraph format", check=lambda o: "warnings" not in o)
    # Word for Mac cannot create a style through its scripting; there the
    # check is that it says so, and the steps after it use Body Text.
    style = "Body Text" if sys.platform == "darwin" else "MCP 本文"
    if sys.platform == "darwin":
        step("creating a style is refused on the Mac", formatting._set_style_impl,
             "MCP 本文", "paragraph", None, None, None, None, None, None, None, None,
             check=lambda o: "Styles pane" in o.get("error", ""))
    step("define a paragraph style", formatting._set_style_impl, style,
         None if sys.platform == "darwin" else "paragraph", "Normal", None, {"east_asian_font": "BIZ UDPゴシック", "size": 10.5,
                          "color": "#333333"},
         {"first_line_indent_chars": 1, "space_after_lines": 0.5,
          "line_spacing_rule": "exactly", "line_spacing": 18},
         None, None, None, None, undo="Define style",
         check=lambda o: "warnings" not in o and "error" not in o)
    step("read the style back", formatting._get_style_impl, style,
         check=lambda o: o["style"]["paragraph"]["line_spacing_pt"] == 18
         and o["style"]["font"]["size"] == 10.5)
    step("change Heading 1 by its English name", formatting._set_style_impl, "Heading 1",
         None, None, None, {"color": "#1F4E79"},
         {"space_before": 12, "keep_with_next": True}, None, None, None, None,
         undo="Define style", check=lambda o: "warnings" not in o)
    step("apply the style", formatting._apply_style_impl, 3, None, style, False,
         undo="Apply style", check=lambda o: "error" not in o and "warnings" not in o)
    step("apply Heading 2", formatting._apply_style_impl, 5, None, "Heading 2", True,
         undo="Apply style", check=lambda o: "warnings" not in o)
    step("rewrite a paragraph", text._set_paragraph_text_impl, 4, "Rewritten heading",
         undo="Rewrite paragraph",
         check=lambda o: o.get("paragraph", {}).get("text") == "Rewritten heading")
    step("undo the rewrite in one step", edit._undo_impl, 1,
         check=lambda o: o.get("undone") == 1)
    step("the rewrite is gone", text._get_paragraphs_impl, 4, 4, False,
         check=lambda o: o["paragraphs"][0]["text"] == "Second heading")
    step("find and replace", text._find_replace_impl, "cat", "dog", False, True, False,
         undo="Replace", check=lambda o: o.get("replaced") == 1)
    step("delete a paragraph", text._delete_paragraphs_impl, 5, None,
         undo="Delete paragraphs", check=lambda o: o.get("paragraphs") == 4)
    step("styles in use", formatting._list_styles_impl, True,
         check=lambda o: len(o.get("styles", [])) >= 3)
    step("document info", text._get_document_info_impl,
         check=lambda o: o.get("paragraphs") == 4)
    step("save", documents._save_document_impl)
    pdf = os.path.splitext(path)[0] + ".pdf"
    step("export PDF", documents._save_document_as_impl, pdf, True,
         check=lambda o: o.get("bytes", 0) > 0)


if __name__ == "__main__":
    raise SystemExit(main())
