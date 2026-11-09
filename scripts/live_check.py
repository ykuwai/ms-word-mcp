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
import re
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.backend import call_document  # noqa: E402
from ms_word_mcp.tools import (  # noqa: E402
    _common, documents, edit, equations, fields, formatting, links, lists, notes, page,
    references, screen, tables, text)

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
    # A Windows console or a redirected file may be cp932 or the like, which
    # cannot write the math characters the equation steps answer.
    for stream in (sys.stdout, sys.stderr):
        stream.reconfigure(errors="backslashreplace")
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
    step("styles in use", formatting._list_styles_impl, True, "text",
         check=lambda o: len(o.get("styles", [])) >= 3)
    step("every table style", formatting._list_styles_impl, False, "table",
         check=lambda o: len(o.get("styles", [])) >= 20
         and all(s["type"] == "table" for s in o["styles"]))
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
    full = json.loads(_common.run(text._get_outline_impl, True))
    step("the outline from Word's heading list is the one from every paragraph",
         text._get_outline_impl, False,
         check=lambda o: o["headings"] == full["headings"] and len(o["headings"]) >= 2)
    run_table_steps()
    run_picture_and_page_steps(Path(path).parent)
    run_review_steps()
    run_heading_number_steps()
    run_border_steps()
    run_note_steps()
    run_link_steps()
    step("picture of page 1", screen._get_page_image_impl, 1, None,
         check=lambda o: o.get("width") == 1200 and os.path.getsize(o["path"]) > 10000)
    step("picture of the page a paragraph is on", screen._get_page_image_impl, None, 5,
         check=lambda o: o.get("page", 0) >= 1 and os.path.exists(o["path"]))
    step("save", documents._save_document_impl)
    pdf = os.path.splitext(path)[0] + ".pdf"
    step("export PDF", documents._save_document_as_impl, pdf, True,
         check=lambda o: o.get("bytes", 0) > 0)


def run_review_steps() -> None:
    """What a code review found: spaces, many matches, ruby, look-alike headings.

    In a document of its own, so the paragraph numbers are short and nothing
    above moves.
    """
    other = step("a document for the review's cases", documents._create_document_impl, None, None)
    if "error" in other:
        return
    token = call_document.set(other["full_name"])
    try:
        made = step("headings, and a body paragraph with a heading's words",
                    text._insert_paragraphs_impl,
                    [{"text": "Intro", "style": "Heading 1"},
                     {"text": "Next", "style": "Heading 1"},
                     {"text": "Next", "style": None},
                     {"text": "\u3000indented, with spaces after  ", "style": None},
                     {"text": "aa aa x", "style": None},
                     {"text": "zeta " * 520, "style": None},
                     {"text": "swap me", "style": None},
                     {"text": "漢字と文章", "style": None}], None, None)
        if "error" in made:
            return
        step("an equation that adds marks to the text", equations._insert_equation_impl,
             "E=mc^2", "latex", True, "Equation", 1, None)
        step("the body paragraph gets outline level 1", formatting._set_paragraph_format_impl,
             4, 4, {"outline_level": "1"}, undo="Paragraph format")
        step("the outline holds the heading, not the body paragraph", text._get_outline_impl, False,
             check=lambda o: [h["index"] for h in o["headings"]] == [1, 3])
        step("leading and trailing spaces are written", text._get_paragraphs_impl, 5, 5, False,
             check=lambda o: o["paragraphs"][0]["text"] == "\u3000indented, with spaces after  ")
        step("a search ending in a space", text._find_replace_impl, "aa ", "b ", False, False,
             False, undo="Replace", check=lambda o: o.get("replaced") == 2)
        step("only the spaces searched for went", text._get_paragraphs_impl, 6, 6, False,
             check=lambda o: o["paragraphs"][0]["text"] == "b b x")
        # One call may run out of time (Windows: 433 of these in a call) and says
        # so; calling again does the rest.
        replaced = 0
        for call_number in range(1, 6):
            out = step(f"more than 500 matches, call {call_number}", text._find_replace_impl,
                       "zeta", "eta", True, False, False, undo="Replace",
                       check=lambda o: o.get("replaced", 0) > 0
                       and (o.get("complete", True) or "Call again" in o["warnings"][-1]))
            replaced += out.get("replaced", 0)
            if out.get("complete", True):
                break
        step("all 520 were replaced", text._find_replace_impl, "zeta", None, True, False, False,
             check=lambda o: o.get("matches") == 0 and replaced == 520)
        step("a replacement beyond the basic plane reads back", text._find_replace_impl, "swap me",
             "\U0001F600 done", True, False, False, undo="Replace",
             check=lambda o: o.get("replaced") == 1 and "warnings" not in o)
        step("ruby to keep", page._add_ruby_impl, 9, [{"text": "漢字", "ruby": "かんじ"}], "center",
             None, undo="Add ruby", check=lambda o: len(o.get("ruby", [])) == 1)
        step("a changed reading with a span that is not there is refused", page._add_ruby_impl, 9,
             [{"text": "漢字", "ruby": "カンジ"}, {"text": "ない", "ruby": "x"}], "center", None,
             undo="Add ruby", check=lambda o: "Nothing was changed" in o.get("error", ""))
        step("the old reading is still there", text._get_paragraphs_impl, 9, 9, False,
             check=lambda o: [r["ruby"] for r in o["paragraphs"][0].get("ruby", [])] == ["かんじ"])
        step("close the document of the review's cases", documents._close_document_impl, False)
    finally:
        call_document.reset(token)


def run_heading_number_steps() -> None:
    """Numbered headings through linked list levels, and a read that runs out of time."""
    other = step("a document for numbered headings", documents._create_document_impl, None, None)
    if "error" in other:
        return
    token = call_document.set(other["full_name"])
    try:
        heads = [("Intro", "Heading 1"), ("Body", None), ("Scope", "Heading 2"),
                 ("Aim", "Heading 2"), ("Method", "Heading 1"), ("Steps", "Heading 2")]
        made = step("headings and body text", text._insert_paragraphs_impl,
                    [{"text": t, "style": s} for t, s in heads], None, None)
        if "error" in made:
            return
        step("number the first heading, with levels linked to the heading styles",
             lists._apply_list_impl, 1, 1,
             [{"level": 1, "number_style": "arabic", "number_format": "%1",
               "linked_style": "Heading 1"},
              {"level": 2, "number_style": "arabic", "number_format": "%1.%2",
               "linked_style": "Heading 2"}], None, False,
             undo="Apply list", check=lambda o: o["paragraphs"][0]["list_string"] == "1")
        wanted = {1: "1", 3: "1.1", 4: "1.2", 5: "2", 6: "2.1"}
        for index, number in wanted.items():
            step(f"heading {index} is numbered {number}", lists._get_list_impl, index,
                 check=lambda o, n=number: o["paragraph"].get("list_string") == n)
        step("body text is not numbered", lists._get_list_impl, 2,
             check=lambda o: o["paragraph"]["list"] == "none")
        step("a heading added later takes the next number", text._insert_paragraphs_impl,
             [{"text": "Later", "style": "Heading 2"}], None, None, undo="Insert paragraphs")
        step("and reads 2.2", lists._get_list_impl, 7,
             check=lambda o: o["paragraph"].get("list_string") == "2.2")
        step("a numbered list after the headings", text._insert_paragraphs_impl,
             [{"text": t, "style": "Normal"} for t in ("a", "b", "c", "d")], None, None,
             undo="Insert paragraphs")
        step("numbered 1 to 4", lists._apply_list_impl, 8, 11,
             [{"level": 1, "number_style": "arabic", "number_format": "%1."}], None, False,
             undo="Apply list", check=lambda o: o["paragraphs"][3]["list_string"] == "4.")
        if backend.IS_MACOS:
            step("starting over is refused on the Mac", lists._restart_numbering_impl, 10,
                 check=lambda o: "Nothing was changed" in o.get("error", ""))
        else:
            step("start over at the third", lists._restart_numbering_impl, 10,
                 undo="Restart numbering",
                 check=lambda o: "warnings" not in o and o["paragraph"]["list_string"] == "1."
                 and o["next"]["list_string"] == "2.")
        step("the ones before keep their numbers", lists._get_list_impl, 9,
             check=lambda o: o["paragraph"]["list_string"] == "2.")
        step("the headings keep theirs", lists._get_list_impl, 7,
             check=lambda o: o["paragraph"]["list_string"] == "2.2")
        limit = text.READ_SECONDS
        text.READ_SECONDS = -1.0
        try:
            step("a read out of time answers what it has", text._get_paragraphs_impl, 1, 3, False,
                 check=lambda o: len(o["paragraphs"]) == 1 and o["complete"] is False
                 and "start=2" in o["more"])
        finally:
            text.READ_SECONDS = limit
        step("close the document of numbered headings", documents._close_document_impl, False)
    finally:
        call_document.reset(token)


def run_border_steps() -> None:
    """Paragraph borders and background colors, directly and through a style."""
    other = step("a document for borders", documents._create_document_impl, None, None)
    if "error" in other:
        return
    token = call_document.set(other["full_name"])
    try:
        made = step("three paragraphs", text._insert_paragraphs_impl,
                    [{"text": "ruled", "style": None}, {"text": "plain", "style": None},
                     {"text": "styled", "style": "Heading 1"}], None, None)
        if "error" in made:
            return
        rule = {"style": "single", "width": 1.5, "color": "#1F4E79"}
        step("a rule under a paragraph, and a background", formatting._set_paragraph_format_impl,
             1, 1, {"borders": {"bottom": rule}, "shading": "#DDEEFF"}, undo="Paragraph format",
             check=lambda o: "warnings" not in o and o["paragraphs"][0]["borders"]["bottom"] == rule
             and o["paragraphs"][0]["shading"] == "#DDEEFF")
        step("the next paragraph is untouched", formatting._set_paragraph_format_impl, 2, 2,
             {"borders": {"bottom": {"style": "none"}}, "shading": "none"}, undo="Paragraph format",
             check=lambda o: "warnings" not in o and o["paragraphs"][0]["borders"] == {}
             and o["paragraphs"][0]["shading"] is None)
        step("a border and a background come off", formatting._set_paragraph_format_impl, 1, 1,
             {"borders": {"bottom": {"style": "none"}}, "shading": "none"}, undo="Paragraph format",
             check=lambda o: "warnings" not in o and o["paragraphs"][0]["borders"] == {}
             and o["paragraphs"][0]["shading"] is None)
        double = {"style": "double", "width": 0.75, "color": "#C00000"}
        step("a style holds a border and a background", formatting._set_style_impl,
             "Heading 1", False, None, None, None,
             {"borders": {"top": double}, "shading": "#FFEEDD"}, None, None, None, None,
             undo="Set style",
             check=lambda o: "warnings" not in o
             and o["style"]["paragraph"]["borders"]["top"] == double
             and o["style"]["paragraph"]["shading"] == "#FFEEDD")
        step("close the document of borders", documents._close_document_impl, False)
    finally:
        call_document.reset(token)


def run_note_steps() -> None:
    """Footnotes and endnotes: where the mark lands, the numbers, and the text."""
    other = step("a document for notes", documents._create_document_impl, None, None)
    if "error" in other:
        return
    token = call_document.set(other["full_name"])
    try:
        made = step("two paragraphs for notes", text._insert_paragraphs_impl,
                    [{"text": "Alpha beta gamma.", "style": "Normal"},
                     {"text": "日本語の本文です。", "style": "Normal"}], None, None)
        if "error" in made:
            return
        step("a footnote at the end of a paragraph", notes._insert_note_impl, 1, "脚注 𠮷 first",
             "footnote", None, undo="Insert footnote",
             check=lambda o: (o["number"], o["index"], o["text"]) == (1, 1, "脚注 𠮷 first"))
        step("a footnote before it takes number 1", notes._insert_note_impl, 1, "earlier",
             "footnote", "Alpha", undo="Insert footnote",
             check=lambda o: (o["number"], o["index"], o["text"]) == (1, 1, "earlier"))
        step("the marks are in the text where they were put", text._get_paragraphs_impl, 1, 1,
             False, check=lambda o: o["paragraphs"][0]["text"] == "Alpha\x02 beta gamma.\x02")
        step("an endnote after some Japanese text", notes._insert_note_impl, 2, "文末脚注",
             "endnote", "本文", undo="Insert endnote",
             check=lambda o: (o["number"], o["index"], o["text"]) == (1, 2, "文末脚注"))
        step("the Japanese paragraph holds its mark", text._get_paragraphs_impl, 2, 2, False,
             check=lambda o: o["paragraphs"][0]["text"] == "日本語の本文\x02です。")
        step("the footnotes in order, and both counts", notes._list_notes_impl, "footnote", 1, 100,
             check=lambda o: [(n["number"], n["text"]) for n in o["notes"]]
             == [(1, "earlier"), (2, "脚注 𠮷 first")]
             and (o["footnotes"], o["endnotes"]) == (2, 1))
        step("one note at a time says what follows", notes._list_notes_impl, "footnote", 1, 1,
             check=lambda o: len(o["notes"]) == 1 and "start=2" in o["more"])
        step("a note's text changes", notes._set_note_impl, 2, "changed", "footnote",
             undo="Change footnote",
             check=lambda o: (o["number"], o["text"]) == (2, "changed"))
        step("text that is not there is refused", notes._insert_note_impl, 1, "no", "footnote",
             "nowhere", check=lambda o: "is not in paragraph 1" in o["error"])
        step("a note that is not there is refused", notes._delete_note_impl, 9, "footnote",
             check=lambda o: "There is no footnote 9" in o["error"])
        step("deleting a note renumbers the next", notes._delete_note_impl, 1, "footnote",
             undo="Delete footnote", check=lambda o: (o["deleted"], o["left"]) == (1, 1))
        step("the one left is number 1, its mark in place", notes._list_notes_impl, "footnote",
             1, 100, check=lambda o: [(n["number"], n["text"]) for n in o["notes"]]
             == [(1, "changed")])
        step("and the deleted mark is gone", text._get_paragraphs_impl, 1, 1, False,
             check=lambda o: o["paragraphs"][0]["text"] == "Alpha beta gamma.\x02")
        step("close the document of notes", documents._close_document_impl, False)
    finally:
        call_document.reset(token)


def run_link_steps() -> None:
    """Hyperlinks to addresses and bookmarks, and the user's own bookmarks."""
    other = step("a document for links", documents._create_document_impl, None, None)
    if "error" in other:
        return
    token = call_document.set(other["full_name"])
    url = "https://example.com/a?b=1&c=2"
    try:
        made = step("two paragraphs for links", text._insert_paragraphs_impl,
                    [{"text": "Alpha beta gamma.", "style": "Normal"},
                     {"text": "日本語の本文です。", "style": "Normal"}], None, None)
        if "error" in made:
            return
        step("a bookmark over Japanese text", links._add_bookmark_impl, "本文_1", 2, "本文",
             undo="Add bookmark",
             check=lambda o: (o["name"], o["index"], o["text"], o["empty"])
             == ("本文_1", 2, "本文", False) and "note" not in o)
        step("a bookmark over a whole paragraph", links._add_bookmark_impl, "Whole", 1, None,
             undo="Add bookmark", check=lambda o: o["text"] == "Alpha beta gamma.")
        step("the same name moves it, and says so", links._add_bookmark_impl, "Whole", 2, None,
             undo="Add bookmark",
             check=lambda o: o["index"] == 2 and "moved" in o["note"])
        step("bookmarks in the order of the text", links._list_bookmarks_impl,
             check=lambda o: sorted(b["name"] for b in o["bookmarks"]) == ["Whole", "本文_1"])
        step("a link over some text", links._insert_hyperlink_impl, 1, url, None, "gamma", None,
             None, undo="Insert hyperlink",
             check=lambda o: (o["number"], o["index"], o["text"], o["address"])
             == (1, 1, "gamma", url))
        step("a link before it takes number 1, past a field", links._insert_hyperlink_impl, 1,
             "https://example.com/b", None, "beta", None, "A tip", undo="Insert hyperlink",
             check=lambda o: (o["number"], o["text"], o["screen_tip"]) == (1, "beta", "A tip"))
        step("a new link at the end, to a bookmark", links._insert_hyperlink_impl, 1, None,
             "本文_1", None, " 続き", None, undo="Insert hyperlink",
             check=lambda o: (o["number"], o["text"], o["address"], o["bookmark"])
             == (3, " 続き", None, "本文_1"))
        step("the paragraph reads as text", text._get_paragraphs_impl, 1, 1, False,
             check=lambda o: o["paragraphs"][0]["text"] == "Alpha beta gamma. 続き")
        step("a bookmark that is not there is refused", links._insert_hyperlink_impl, 2, None,
             "Nope", None, None, None, check=lambda o: "no bookmark named 'Nope'" in o["error"])
        step("the links in order", links._list_hyperlinks_impl, 1, 100,
             check=lambda o: [h["text"] for h in o["hyperlinks"]] == ["beta", "gamma", " 続き"])
        step("a link's address, text and tip change", links._set_hyperlink_impl, 2,
             "https://example.com/new", None, "GAMMA", "New tip", undo="Change hyperlink",
             check=lambda o: (o["address"], o["text"], o["screen_tip"])
             == ("https://example.com/new", "GAMMA", "New tip"))
        step("a tip comes off", links._set_hyperlink_impl, 2, None, None, None, "",
             undo="Change hyperlink", check=lambda o: o["screen_tip"] is None)
        step("a bookmark link goes to an address", links._set_hyperlink_impl, 3,
             "https://example.com/z", None, None, None, undo="Change hyperlink",
             check=lambda o: (o["address"], o["bookmark"]) == ("https://example.com/z", None))
        step("an address link becomes a bookmark link, its tip kept", links._set_hyperlink_impl,
             1, None, "Whole", None, None, undo="Change hyperlink",
             check=lambda o: (o["number"], o["text"], o["address"], o["bookmark"], o["screen_tip"])
             == (1, "beta", None, "Whole", "A tip"))
        step("the paragraph has not moved", text._get_paragraphs_impl, 1, 1, False,
             check=lambda o: o["paragraphs"][0]["text"] == "Alpha beta GAMMA. 続き")
        step("a link comes off, its text stays", links._remove_hyperlink_impl, 2,
             undo="Remove hyperlink",
             check=lambda o: (o["removed"], o["text_kept"], o["left"]) == (2, "GAMMA", 2))
        step("and the paragraph still reads the same", text._get_paragraphs_impl, 1, 1, False,
             check=lambda o: o["paragraphs"][0]["text"] == "Alpha beta GAMMA. 続き")
        step("a link that is not there is refused", links._remove_hyperlink_impl, 9,
             check=lambda o: "There is no hyperlink 9" in o["error"])
        step("a bookmark is deleted", links._delete_bookmark_impl, "Whole",
             undo="Delete bookmark", check=lambda o: (o["deleted"], o["left"]) == ("Whole", 1))
        step("a bookmark that is not there is refused", links._delete_bookmark_impl, "Whole",
             check=lambda o: "no bookmark named 'Whole'" in o["error"])
        step("close the document of links", documents._close_document_impl, False)
    finally:
        call_document.reset(token)


def run_picture_and_page_steps(folder: Path) -> None:
    picture = folder / "live-picture.png"
    picture.write_bytes(screen.png(60, 30, [b"\x40\x80\xc0" * 60] * 30))
    pic = step("a picture, 150 pt wide, height in proportion", references._insert_picture_impl,
               str(picture), None, None, 150.0, None, "center", "A blue box",
               undo="Insert picture",
               check=lambda o: o["width"] == 150 and abs(o["height"] - 75) < 1)
    picture.unlink()
    if "error" in pic:
        return
    step("picture read back", references._list_pictures_impl,
         check=lambda o: o["pictures"][-1]["alt_text"] == "A blue box")
    step("a picture wider than the text warns", references._set_picture_impl, pic["picture"],
         900.0, None, None, None, undo="Change picture",
         check=lambda o: "right margin" in " ".join(o.get("warnings", [])))
    step("and back", references._set_picture_impl, pic["picture"], 150.0, None, None, None,
         undo="Change picture", check=lambda o: "warnings" not in o)
    cap = step("a caption under the picture", references._insert_caption_impl, "A blue box",
               "Figure", ": ", pic["picture"], None, None, None, undo="Insert caption",
               check=lambda o: o["number"] == "1" and o["text"] == "Figure 1: A blue box")
    tab = step("a table caption in another label, in Japanese", references._insert_caption_impl,
               "売上", "表", " ", None, 2, None, None, undo="Insert caption",
               check=lambda o: o["number"] == "1" and o["text"] == "表 1 売上")
    step("captions are listed", references._list_captions_impl,
         check=lambda o: {"Figure", "表", "Equation"} <= {c["label"] for c in o["captions"]})
    para = step("a paragraph to refer from", text._insert_paragraphs_impl,
                [{"text": "As PLACEHOLDER shows, and SECOND, and on page .", "style": "Normal"}], None, None,
                undo="Insert paragraphs", check=lambda o: "error" not in o)
    if "error" in cap or "error" in para:
        return
    at = para["inserted"][0]["index"]
    # The table's caption went in above the figure's, so its number is asked again.
    figure = next(c["index"] for c in
                  json.loads(_common.run(references._list_captions_impl))["captions"]
                  if c["label"] == "Figure")
    step("a reference to the figure", references._insert_cross_reference_impl, at, figure,
         "label_and_number", "end", "PLACEHOLDER", undo="Insert cross-reference",
         check=lambda o: o["result"] == "Figure 1")
    table_caption = next(c["index"] for c in
                         json.loads(_common.run(references._list_captions_impl))["captions"]
                         if c["label"] == "表")
    step("a second reference in the paragraph", references._insert_cross_reference_impl, at,
         table_caption, "label_and_number", "end", "SECOND", undo="Insert cross-reference",
         check=lambda o: o.get("result") == "表 1")
    step("both references are in one piece", fields._list_fields_impl, at, at,
         check=lambda o: [f["kind"] for f in o["fields"]] == ["REF", "REF"]
         and all(f["code"].count("REF") == 1 for f in o["fields"]))
    step("a reference to its page", references._insert_cross_reference_impl, at, figure,
         "page", "end", None, undo="Insert cross-reference",
         check=lambda o: o["result"].isdigit())
    step("a heading has no number to refer to", references._insert_cross_reference_impl, at, 4,
         "number", "end", None, undo="Insert cross-reference",
         check=lambda o: "no number" in o.get("error", ""))
    equation = next(c for c in json.loads(_common.run(references._list_captions_impl))["captions"]
                    if c["label"] == "Equation")
    step("a reference to an equation's number", references._insert_cross_reference_impl, at,
         equation["index"], "number", "end", None, undo="Insert cross-reference",
         check=lambda o: o.get("result") == equation["number"])
    earlier = step("a figure caption above renumbers the others",
                   references._insert_caption_impl, "Earlier", "Figure", ": ", None, None, 2, None,
                   undo="Insert caption", check=lambda o: o["number"] == "1")
    if "error" not in earlier:
        step("a caption before a table's first cell stays out of the table",
             tables._list_tables_impl,
             check=lambda o: all(not t["first_row"][0].startswith(("Figure", "表"))
                                 for t in o["tables"]))
        step("the reference followed", text._get_paragraphs_impl, at + 1, at + 1, False,
             undo=None, check=lambda o: "Figure 2" in o["paragraphs"][0]["text"])
    step("a header", page._set_header_footer_impl, "header", "Live check", "primary", 1, "right",
         undo="Set header or footer", check=lambda o: o["reads"] == "Live check")
    step("a footer with the page number and count", page._set_header_footer_impl, "footer",
         "Page {page} of {pages}", "primary", 1, "center", undo="Set header or footer",
         check=lambda o: re.fullmatch(r"Page \d+ of \d+", o["reads"]))
    step("a first page of its own", page._set_header_footer_impl, "header", "", "first", 1, None,
         undo="Set header or footer", check=lambda o: o["reads"] == "")
    step("roman page numbers from 5", page._set_page_numbers_impl, 1, "lowercase_roman", 5, None,
         undo="Set page numbering",
         check=lambda o: o["style"] == "lowercase_roman" and o["start_at"] == 5)
    step("back to arabic from 1", page._set_page_numbers_impl, 1, "arabic", 1, None,
         undo="Set page numbering", check=lambda o: o["style"] == "arabic")
    step("headers and footers read back", page._get_headers_footers_impl,
         check=lambda o: o["sections"][0]["header"]["primary"]["text"] == "Live check"
         and o["sections"][0]["header"]["first"]["text"] == "")
    step("not a first page of its own", page._set_page_setup_impl, None,
         {"different_first_page": False}, undo="Set page setup",
         check=lambda o: o["sections"][0]["different_first_page"] is False)
    step("A4, margins of 72", page._set_page_setup_impl, None,
         {"paper": "a4", "top_margin": 72.0, "bottom_margin": 72.0}, undo="Set page setup",
         check=lambda o: o["sections"][0]["paper"] == "a4"
         and o["sections"][0]["margins"]["top"] == 72 and "warnings" not in o)
    step("vertical text keeps the page upright", page._set_page_setup_impl, None,
         {"text_direction": "vertical"}, undo="Set page setup",
         check=lambda o: o["sections"][0]["text_direction"] == "vertical"
         and o["sections"][0]["orientation"] == "portrait" and "warnings" not in o)
    step("back to horizontal", page._set_page_setup_impl, None,
         {"text_direction": "horizontal"}, undo="Set page setup",
         check=lambda o: o["sections"][0]["text_direction"] == "horizontal"
         and o["sections"][0]["orientation"] == "portrait")
    step("no table floats after the change of direction", _tables_float,
         check=lambda o: o["float"] == [False] * o["tables"] and o["tables"] >= 2)
    step("two columns, and back to one", page._set_page_setup_impl, None, {"columns": 2},
         undo="Set page setup", check=lambda o: o["sections"][0]["columns"] == 2)
    step("one column", page._set_page_setup_impl, None, {"columns": 1}, undo="Set page setup",
         check=lambda o: o["sections"][0]["columns"] == 1)
    ruby = step("a paragraph for ruby", text._insert_paragraphs_impl,
                [{"text": "漢字と日本語の文章です。", "style": "Normal"}], None, None,
                undo="Insert paragraphs", check=lambda o: "error" not in o)
    if "error" not in ruby:
        at = ruby["inserted"][0]["index"]
        step("ruby on two words", page._add_ruby_impl, at,
             [{"text": "漢字", "ruby": "かんじ"}, {"text": "文章", "ruby": "ぶんしょう"}], "center", None,
             undo="Add ruby",
             check=lambda o: o["text"] == "漢字と日本語の文章です。"
             and [r["ruby"] for r in o["ruby"]] == ["かんじ", "ぶんしょう"])
        step("ruby next to existing ruby", page._add_ruby_impl, at,
             [{"text": "日本語", "ruby": "にほんご"}], "center", 5, undo="Add ruby",
             check=lambda o: len(o["ruby"]) == 3 and o["text"] == "漢字と日本語の文章です。")
        step("ruby can be changed", page._add_ruby_impl, at,
             [{"text": "漢字", "ruby": "カンジ"}], "center", None, undo="Add ruby",
             check=lambda o: o["text"] == "漢字と日本語の文章です。" and len(o["ruby"]) == 3
             and o["ruby"][0]["ruby"] == "カンジ")
        step("a paragraph with ruby can be rewritten", text._set_paragraph_text_impl, at,
             "Rewritten after ruby.", undo="Rewrite",
             check=lambda o: o["paragraph"]["text"] == "Rewritten after ruby.")
        step("text that is not there changes nothing", page._add_ruby_impl, at,
             [{"text": "ない", "ruby": "x"}], "center", None, undo="Add ruby",
             check=lambda o: "Nothing was changed" in o.get("error", ""))


def _tables_float():
    doc = backend.word.document()
    if backend.IS_MACOS:
        from appscript import k

        n = int(doc.count(each=k.table))
        return {"tables": n,
                "float": [bool(doc.tables[i].row_options.wrap_around_text.get())
                          for i in range(1, n + 1)]}
    n = doc.Tables.Count
    return {"tables": n, "float": [bool(doc.Tables.Item(i).Rows.WrapAroundText)
                                   for i in range(1, n + 1)]}


_common.MAC_IMPLS[_tables_float.__name__] = _tables_float


def _active_document():
    return {"name": backend.word.get_app().active_document.name.get()}


_common.MAC_IMPLS[_active_document.__name__] = _active_document  # a Mac-only check


def run_table_steps() -> None:
    t = step("table at the end, bold heading row, numbers on the right",
             tables._insert_table_impl,
             [["Item", "Qty", "Price"], ["Apple", "3", "1.20"], ["Pear", "10", "0.80"]],
             None, None, 1, None, None, None, None, None, [None, "right", "right"],
             undo="Insert table",
             check=lambda o: (o["rows"], o["columns"], o["header_rows"]) == (3, 3, 1)
             and o["cells"][1] == ["Apple", "3", "1.20"])
    if "error" in t:
        return
    n = t["table"]
    step("paragraphs in a table carry its number", text._get_paragraphs_impl,
         t["paragraphs"]["first"] + 1, t["paragraphs"]["first"] + 1, False,
         check=lambda o: o["paragraphs"][0].get("table") == n)
    second = step("a table right after it stays a table of its own",
                  tables._insert_table_impl, [["A", "B"], ["1", "2"]],
                  t["paragraphs"]["last"], None, 0, None, None, None, None, None, None,
                  undo="Insert table", check=lambda o: o["table"] == n + 1)
    before = json.loads(_common.run(text._get_paragraphs_impl, 3, 3, False))
    mid = step("a styled table before paragraph 3 leaves no empty paragraph",
               tables._insert_table_impl, [["Key", "Value"], ["x", "1"]], None, 3, 1,
               "Light Grid", None, [120.0, 80.0], None, "center", None,
               undo="Insert table", check=lambda o: o["style"] and o["table"] == 1)
    if "error" not in mid:
        n += 1
        after = json.loads(_common.run(text._get_paragraphs_impl,
                                       mid["paragraphs"]["last"] + 1,
                                       mid["paragraphs"]["last"] + 1, False))
        results.append(("the paragraph after the table is the one it was put before",
                         after["paragraphs"][0]["text"] == before["paragraphs"][0]["text"]))
        print(f"{'PASS' if results[-1][1] else 'FAIL'}  {results[-1][0]}")
        step("a table inside a cell is refused", tables._insert_table_impl, [["no"]], None,
             mid["paragraphs"]["first"] + 1, 0, None, None, None, None, None, None,
             undo="Insert table", check=lambda o: "inside a table" in o.get("error", ""))
    step("write a block of cells", tables._set_cells_impl, n, 2, 2, [["4", None], ["12"]],
         undo="Write table cells",
         check=lambda o: o["written"] == 2 and o["cells"][0][1] == "4")
    step("shade the heading row", tables._format_cells_impl, n, 1, 1, 1, None,
         "#D9E2F3", None, None, None, None, "center", "center", undo="Format table cells",
         check=lambda o: o["cells"] == 3 and o["first_cell_reads"]["shading"] == "#D9E2F3")
    step("rows at the bottom are data, not headings", tables._insert_rows_impl, n, None,
         None, [["Plum", "7", "0.50"], ["Fig", "2", "2.10"]], undo="Insert table rows",
         check=lambda o: o["rows"] == 5 and o["cells"][1][0] == "Fig")
    step("still one heading row", tables._get_table_impl, n, 1,
         check=lambda o: o["header_rows"] == 1 and o["cells"][4][0] == "Fig")
    step("delete a row", tables._delete_rows_impl, n, 4, None, undo="Delete table rows",
         check=lambda o: o["rows"] == 4)
    if backend.IS_MACOS:
        step("adding a column is refused on the Mac", tables._insert_columns_impl, n, 1,
             None, [["Unit", "kg", "kg", "kg"]], undo="Insert table columns",
             check=lambda o: "cannot add a column" in o.get("error", ""))
    else:
        step("add a column", tables._insert_columns_impl, n, 1, None,
             [["Unit", "kg", "kg", "kg"]], undo="Insert table columns",
             check=lambda o: o["columns"] == 4 and o["cells"][0][1] == "Unit")
        step("delete the column", tables._delete_columns_impl, n, 2, None,
             undo="Delete table columns", check=lambda o: o["columns"] == 3)
    step("widths, centred, Medium Shading 1 Accent 1", tables._set_table_impl, n,
         "Medium Shading 1 Accent 1", None, None, [120.0, 60.0, 60.0], None, "center",
         undo="Change table", check=lambda o: o["style"] and o["header_rows"] == 1)
    step("columns wider than the text warn", tables._set_table_impl, n, None, None, None,
         [300.0, 120.0, 120.0], None, None, undo="Change table",
         check=lambda o: "right margin" in " ".join(o.get("warnings", [])))
    step("and back", tables._set_table_impl, n, None, None, None, [120.0, 60.0, 60.0],
         None, None, undo="Change table", check=lambda o: "warnings" not in o)
    step("merge the last row's first two cells", tables._merge_cells_impl, n, 4, 1, 4, 2,
         "Total", undo="Merge table cells",
         check=lambda o: o["merged_cell"] == "Total" and len(o["cells"][3]) == 2)
    step("unshade, and bold, italic, color and size", tables._format_cells_impl, n, 2, 2, 1,
         1, "none", True, True, "#C00000", 12.0, None, None, undo="Format table cells",
         check=lambda o: o["cells"] == 1 and o["first_cell_reads"]["shading"] is None
         and o["first_cell_reads"]["bold"] is True)
    if "error" not in mid:
        step("merge down a column", tables._merge_cells_impl, 1, 1, 1, 2, 1, None,
             undo="Merge table cells", check=lambda o: len(o["cells"][1]) == 1)
        step("read a table merged down a column", tables._get_table_impl, 1, 1,
             check=lambda o: o["rows"] == 2 and o["cells"][1] == ["1"])
        step("restyle it", tables._set_table_impl, 1, "Light Grid", None, None, None, None,
             None, undo="Change table", check=lambda o: o["style"])
        step("rows of it are added or refused with a reason", tables._insert_rows_impl, 1,
             None, 1, None, undo="Insert table rows",
             check=lambda o: "error" not in o or "merged down a column" in o["error"])
    other = step("another document becomes the active one",
                 documents._create_document_impl, None, None,
                 check=lambda o: "error" not in o)
    if "error" not in other:
        step("add a row in a document that is not the active one",
             tables._insert_rows_impl, n, 1, None, [["Kiwi", "1", "0.30"]],
             undo="Insert table rows", check=lambda o: o["cells"][0][0] == "Kiwi")
        step("merge in it", tables._merge_cells_impl, n, 2, 2, 2, 3, None,
             undo="Merge table cells", check=lambda o: "error" not in o)
        if backend.IS_MACOS:
            step("the other document is still the active one", _active_document,
                 check=lambda o: o["name"] == other["name"])
        token = call_document.set(other["full_name"])
        try:
            step("close the other document", documents._close_document_impl, False)
        finally:
            call_document.reset(token)
    step("list the tables", tables._list_tables_impl,
         check=lambda o: len(o["tables"]) == 3 and o["tables"][1]["first_row"][0] == "Item")
    if "error" not in second:
        step("delete a table", tables._delete_table_impl, n + 1, undo="Delete table",
             check=lambda o: o["tables"] == 2)


if __name__ == "__main__":
    raise SystemExit(main())
