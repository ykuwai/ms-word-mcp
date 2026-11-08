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
from ms_word_mcp.tools import _common, documents, edit, equations, fields, formatting, lists, screen, text  # noqa: E402

results: list[tuple[str, bool]] = []


def step(label, impl, *args, undo=None, check=lambda out: "error" not in out):
    out = json.loads(_common.run(impl, *args, undo=undo))
    try:
        ok = bool(check(out))
    except (KeyError, IndexError, TypeError):
        ok = False
    results.append((label, ok))
    print(f"{'PASS' if ok else 'FAIL'}  {label}")
    text = json.dumps(out, ensure_ascii=False)
    # A failure is printed whole, its warnings first, so nothing that
    # explains it is cut off.
    if not ok and "warnings" in out:
        print("      warnings: " + json.dumps(out["warnings"], ensure_ascii=False))
    print("      " + (text if not ok else text[:600]))
    return out


def main() -> int:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    # Inside the repository (.word-test/, ignored by git), so nothing outside
    # the project is read or written. Word for Mac, being sandboxed, asks once
    # to be allowed into the folder. The pictures of pages go there too.
    folder = Path(__file__).resolve().parent.parent / ".word-test"
    folder.mkdir(exist_ok=True)
    os.environ.setdefault("WORD_MAC_FILES_DIR", str(folder))
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
    step("right-to-left paragraph and tab stops", formatting._set_paragraph_format_impl,
         5, None, {"direction": "right_to_left",
                   "tab_stops": [{"position": 72, "alignment": "left", "leader": "none"},
                                 {"position": 200, "alignment": "right",
                                  "leader": "dots"}]},
         undo="Paragraph format",
         check=lambda o: (o["paragraphs"][0].get("direction") == "right_to_left"
                          and len(o["paragraphs"][0].get("tab_stops", [])) == 2)
         if sys.platform != "darwin" else len(o.get("warnings", [])) == 2)
    step("back to left-to-right, no tab stops", formatting._set_paragraph_format_impl,
         5, None, {"direction": "left_to_right", "tab_stops": []}, undo="Paragraph format",
         check=lambda o: "error" not in o)
    step("define a paragraph style", formatting._set_style_impl, style,
         None if sys.platform == "darwin" else "paragraph", "Normal", None, {"east_asian_font": "BIZ UDPゴシック", "size": 10.5,
                          "color": "#333333"},
         {"first_line_indent_chars": 1, "space_after_lines": 0.5,
          "line_spacing_rule": "exactly", "line_spacing": 18},
         None, None, None, None, undo="Define style",
         check=lambda o: "warnings" not in o and "error" not in o)
    if sys.platform != "darwin":
        step("the new style takes nothing from where the cursor was",
             formatting._get_style_impl, style,
             check=lambda o: o["style"]["font"].get("character_spacing") == 0
             and o["style"]["font"].get("latin_font") == "+本文のフォント")
    step("read the style back", formatting._get_style_impl, style,
         check=lambda o: o["style"]["paragraph"]["line_spacing_pt"] == 18
         and o["style"]["font"]["size"] == 10.5)
    step("change Heading 1 by its English name", formatting._set_style_impl, "Heading 1",
         None, None, None, {"color": "#1F4E79"},
         {"space_before": 12, "keep_with_next": True}, None, None, None, None,
         undo="Define style", check=lambda o: "warnings" not in o)
    step("apply the style", formatting._apply_style_impl, 3, None, style, False,
         undo="Apply style", check=lambda o: "error" not in o and "warnings" not in o)
    step("insert list items", text._insert_paragraphs_impl, [
        {"text": "First point", "style": None}, {"text": "Second point", "style": None},
        {"text": "A sub-point", "style": None}], None, None, undo="Insert paragraphs",
        check=lambda o: len(o.get("inserted", [])) == 3)
    step("red square bullets, own positions", lists._apply_list_impl, 6, 8, [
        {"level": 1, "bullet": "■", "font": {"color": "#C00000", "size": 9},
         "number_position": 0, "text_position": 14},
        {"level": 2, "bullet": "・", "font": {"color": "#1F4E79"},
         "number_position": 14, "text_position": 28}], None, False, undo="Apply list",
        check=lambda o: "warnings" not in o and o["paragraphs"][0]["list_string"] == "■")
    step("sub-point to level 2", lists._set_list_level_impl, 8, None, 2, undo="List level",
         check=lambda o: o["paragraphs"][0].get("list_string") == "・")
    step("read the list back", lists._get_list_impl, 6,
         check=lambda o: o["levels"][0]["font"]["color"] == "#C00000")
    step("numbered list in full-width digits", lists._apply_list_impl, 6, 7, [
        {"level": 1, "number_style": "arabic_full_width", "number_format": "（%1）"}],
        None, False, undo="Apply list",
        check=lambda o: [p.get("list_string") for p in o["paragraphs"]] == ["（１）", "（２）"])
    step("remove the list", lists._remove_list_impl, 6, 8, undo="Remove list",
         check=lambda o: all(p["list"] == "none" for p in o["paragraphs"]))
    step("delete the list items", text._delete_paragraphs_impl, 6, 8,
         undo="Delete paragraphs", check=lambda o: o.get("paragraphs") == 5)
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
    # Equations, fields and tables of contents last: they add paragraphs. The
    # document is now: 1 title, 2 Heading 1, 3 body, 4 Heading 2.
    step("numbered equation from LaTeX", equations._insert_equation_impl,
         r"\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}", "latex", True, "Equation",
         None, None, undo="Insert equation",
         check=lambda o: o.get("index") == 5 and o.get("number") == "1"
         and "▒" in o.get("word_reads_it_as", ""))
    step("second numbered equation, from UnicodeMath", equations._insert_equation_impl,
         "a^2+b^2=c^2", "unicodemath", True, "Equation", None, None, undo="Insert equation",
         check=lambda o: o.get("index") == 6 and o.get("number") == "2")
    step("a numbered equation above them renumbers", equations._insert_equation_impl,
         r"E=mc^2", "latex", True, "Equation", None, 5, undo="Insert equation",
         check=lambda o: o.get("index") == 5 and o.get("number") == "1")
    step("list the equations", equations._list_equations_impl,
         check=lambda o: [e.get("number") for e in o.get("equations", [])]
         == ["1", "2", "3"])
    step("TOC 1 by its English name", formatting._get_style_impl, "TOC 1",
         check=lambda o: o.get("success") is True)
    step("zero characters of indent on a style also zeroes its points",
         formatting._set_style_impl, "Heading 2", None, None, None, None,
         {"first_line_indent_chars": 0}, None, None, None, None, undo="Define style",
         check=lambda o: o["style"]["paragraph"].get("first_line_indent") == 0
         and "warnings" not in o)
    step("read an equation back", equations._get_equation_impl, 6,
         check=lambda o: "√" in o.get("unicodemath", ""))
    step("page number at the end", text._insert_paragraphs_impl,
         [{"text": "Page ", "style": None}], None, None, undo="Insert paragraphs",
         check=lambda o: o["inserted"][0]["index"] == 8)
    step("PAGE field", fields._insert_field_impl, 8, "PAGE", "end", None,
         undo="Insert field", check=lambda o: o.get("result") == "1")
    step("table of contents for part of the document", fields._insert_toc_impl, None,
         None, 1, 3, True, False, 2, 4, undo="Insert table of contents",
         check=lambda o: "bookmark" in o and o.get("entries") == 2)
    step("table of contents at the start", fields._insert_toc_impl, 0, None, 1, 3, True,
         True, None, None, undo="Insert table of contents",
         check=lambda o: o.get("index") == 1 and o.get("entries", 0) >= 2)
    step("list the fields", fields._list_fields_impl, None, None,
         check=lambda o: {"TOC", "SEQ", "PAGE"} <= {f["kind"] for f in o["fields"]})
    step("update the equation numbers", fields._update_fields_impl, "numbers", None, None,
         undo="Update fields", check=lambda o: o.get("updated") == 3)
    step("update the tables of contents' page numbers", fields._update_fields_impl,
         "page_numbers", None, None, undo="Update fields",
         check=lambda o: o.get("updated") == 2)
    step("update every field", fields._update_fields_impl, "all", None, None,
         undo="Update fields", check=lambda o: o.get("updated", 0) >= 6)
    step("text with a caret after the equations", text._insert_paragraphs_impl,
         [{"text": "Then x^2 grows.", "style": None}], None, None,
         undo="Insert paragraphs", check=lambda o: "error" not in o)
    step("replace text that follows equations", text._find_replace_impl, "x^2",
         "x squared", True, False, False, undo="Replace",
         check=lambda o: o.get("replaced") == 1 and "warnings" not in o)
    step("the replacement landed where it was found", text._get_outline_impl,
         check=lambda o: all("squared" not in h["text"] for h in o["headings"]))
    step("picture of page 1", screen._get_page_image_impl, 1, None,
         check=lambda o: o.get("width") == 1200 and os.path.getsize(o["path"]) > 10000)
    step("picture of the page a paragraph is on", screen._get_page_image_impl, None, 5,
         check=lambda o: o.get("page", 0) >= 1 and os.path.exists(o["path"]))
    step("save", documents._save_document_impl)
    pdf = os.path.splitext(path)[0] + ".pdf"
    step("export PDF", documents._save_document_as_impl, pdf, True,
         check=lambda o: o.get("bytes", 0) > 0)


if __name__ == "__main__":
    raise SystemExit(main())
