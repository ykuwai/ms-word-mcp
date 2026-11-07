"""Probe Word for Mac's behaviour for the core tools, in a scratch document.

    uv run python scripts/probes/mac_core.py

Creates a new document, tries each operation, prints what Word answered, and
closes the document without saving. Touches no other document.
"""

import appscript
from appscript import k

w = appscript.app(id="com.microsoft.Word", terms="sdef")


def attempt(label, f):
    try:
        print(f"{label} -> {f()!r}")
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {type(e).__name__} {str(e).splitlines()[0][:120]} "
              f"{getattr(e, 'errornumber', '')}")


def para_count(d):
    lo, hi = 0, 1
    def exists(i):
        try:
            d.paragraphs[i].text_object.start_of_content.get()
            return True
        except Exception:
            return False
    while exists(hi):
        lo, hi = hi, hi * 2
    while hi - lo > 1:
        mid = (lo + hi) // 2
        lo, hi = (mid, hi) if exists(mid) else (lo, mid)
    return lo


d = w.create_new_document()
try:
    r = d.create_range(start=0, end_=0)
    r.content.set("Title\rHeading one\rBody 本文 cat\rLast")
    attempt("count by probing", lambda: para_count(d))
    end = d.text_object.end_of_content.get()
    attempt("end_of_content", lambda: end)
    r2 = d.create_range(start=end - 1, end_=end - 1)
    attempt("insert at end", lambda: r2.content.set("\rAppended"))
    attempt("count after append", lambda: para_count(d))
    attempt("p5 text", lambda: d.paragraphs[5].text_object.content.get())
    attempt("set style by enum", lambda: d.paragraphs[2].style.set(k.style_heading1))
    attempt("p2 style", lambda: d.paragraphs[2].style.name_local.get())
    attempt("p2 outline", lambda: d.paragraphs[2].outline_level.get())
    attempt("p3 outline", lambda: d.paragraphs[3].outline_level.get())
    attempt("set style on range", lambda: d.paragraphs[1].text_object.style.set(k.style_title))
    attempt("p1 style", lambda: d.paragraphs[1].style.name_local.get())
    attempt("style by name", lambda: d.Word_styles["見出し 2"].name_local.get())
    pf = d.paragraphs[3].paragraph_format
    attempt("set multiple 1.5", lambda: (pf.line_spacing_rule.set(k.line_space_multiple),
                                         pf.line_spacing.set(18)))
    attempt("read rule/spacing", lambda: (pf.line_spacing_rule.get(), pf.line_spacing.get()))
    attempt("space before 6", lambda: pf.space_before.set(6))
    attempt("line unit before", lambda: pf.line_unit_before.get())
    attempt("alignment", lambda: (pf.alignment.set(k.align_paragraph_justify), pf.alignment.get()))
    attempt("para alignment via paragraph", lambda: d.paragraphs[3].alignment.get())
    f = d.paragraphs[3].text_object.font_object
    attempt("font size/bold", lambda: (f.font_size.set(14), f.bold.set(True), f.font_size.get(), f.bold.get()))
    attempt("color", lambda: (f.color.set([255, 0, 0]), f.color.get()))
    attempt("selection before", lambda: (w.selection.selection_start.get(), w.selection.selection_end.get()))
    rng = d.paragraphs[3].text_object
    fo = rng.find_object
    attempt("find execute", lambda: fo.execute_find(find_text="cat", match_case=False))
    attempt("found range", lambda: (rng.start_of_content.get(), rng.end_of_content.get()))
    attempt("selection after", lambda: (w.selection.selection_start.get(), w.selection.selection_end.get()))
    attempt("undo 1", lambda: d.undo(times=1))
    attempt("font size after undo", lambda: f.font_size.get())
    attempt("window count", lambda: d.active_window.get())
    attempt("zoom", lambda: d.active_window.view.zoom.percentage.get())
    attempt("statistics pages", lambda: d.compute_statistics(statistic=k.statistic_pages))
    attempt("saved", lambda: d.saved.get())
finally:
    d.close(saving=k.no)
    print("closed")
