"""Measure tables on Word for Mac, in scratch documents closed unsaved.

    uv run python scripts/probes/mac_tables.py
"""

import appscript
from appscript import k

from ms_word_mcp.backend.mac import _decode_wide_integers

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")


def attempt(label, f):
    try:
        print(f"{label} -> {f()!r}")
    except Exception as e:  # noqa: BLE001
        lines = str(e).splitlines()
        print(f"{label} FAILED: {' '.join(x.strip() for x in lines[:3])[:200]}")


def texts(d):
    n = d.count(each=k.paragraph)
    return [d.paragraphs[i].text_object.content.get() for i in range(1, n + 1)]


d = w.create_new_document()
other = None
try:
    d.create_range(start=0, end_=0).content.set("Before\r\rAfter")
    s2 = d.paragraphs[2].text_object.start_of_content.get()
    attempt("make table at empty paragraph", lambda: d.make(
        new=k.table, at=d.create_range(start=s2, end_=s2),
        with_properties={k.number_of_rows: 2, k.number_of_columns: 3}))
    tb = d.tables[1]
    attempt("paragraphs", lambda: texts(d))
    attempt("style (no style given)", lambda: tb.style.get())
    attempt("table range", lambda: (tb.text_object.start_of_content.get(),
                                    tb.text_object.end_of_content.get()))
    for r in (1, 2):
        for c in (1, 2, 3):
            attempt(f"fill {r},{c}", lambda: tb.get_cell_from_table(row=r, column=c)
                    .text_object.content.set(f"r{r}c{c}"))
    attempt("make new row at end", lambda: d.make(new=k.row, at=tb.rows.end))
    attempt("make new row at after row 2", lambda: d.make(new=k.row, at=tb.rows[2].after))
    attempt("rows now", lambda: tb.number_of_rows.get())
    # Selection route (Alon's): put the insertion point in the last row, insert below.
    sel = w.selection
    attempt("selection before", lambda: (sel.selection_start.get(), sel.selection_end.get()))
    c = tb.get_cell_from_table(row=2, column=1)
    attempt("select last row cell", lambda: c.text_object.select())
    attempt("insert rows below", lambda: sel.insert_rows(position=k.insert_below,
                                                          number_of_rows=2))
    attempt("rows after insert rows", lambda: tb.number_of_rows.get())
    attempt("delete row 4", lambda: tb.rows[4].delete())
    attempt("rows after delete", lambda: tb.number_of_rows.get())
    attempt("insert columns", lambda: sel.insert_columns())
    attempt("columns now", lambda: tb.number_of_columns.get())
    attempt("delete column 4", lambda: tb.columns[4].delete())
    attempt("columns after delete", lambda: tb.number_of_columns.get())
    attempt("merge 1,1 with 1,2", lambda: tb.get_cell_from_table(row=1, column=1)
            .merge_cell(with_=tb.get_cell_from_table(row=1, column=2)))
    attempt("uniform after merge", lambda: tb.uniform.get())
    attempt("cell 1,2 after merge", lambda: tb.get_cell_from_table(row=1, column=2)
            .text_object.content.get())
    attempt("cells in row 1", lambda: tb.rows[1].count(each=k.cell))
    cell = tb.get_cell_from_table(row=2, column=1)
    attempt("shading set", lambda: cell.shading.background_pattern_color.set([65535, 0, 0]))
    attempt("shading get", lambda: cell.shading.background_pattern_color.get())
    attempt("vertical alignment set", lambda: cell.vertical_alignment.set(
        k.cell_align_vertical_center))
    attempt("vertical alignment get", lambda: cell.vertical_alignment.get())
    attempt("auto fit window", lambda: tb.auto_fit_behavior(behavior=k.auto_fit_window))
    attempt("preferred width", lambda: (tb.preferred_width_type.get(), tb.preferred_width.get()))
    attempt("auto fit content", lambda: tb.auto_fit_behavior(behavior=k.auto_fit_content))
    attempt("column widths", lambda: [tb.columns[i].width.get() for i in (1, 2, 3)])
    attempt("row alignment get", lambda: tb.rows[1].alignment.get())
    attempt("row options alignment set", lambda: tb.row_options.alignment.set(k.align_row_center))
    attempt("row alignment after", lambda: [tb.rows[i].alignment.get() for i in (1, 2)])
    attempt("heading rows", lambda: tb.rows[1].heading_format.set(True))
    attempt("style by name", lambda: tb.style.set("Table Grid"))
    attempt("style after", lambda: tb.style.get())
    attempt("apply style heading rows", lambda: tb.apply_style_heading_rows.get())
    attempt("tables count", lambda: d.count(each=k.table))
    attempt("paragraphs", lambda: texts(d))
    # A second table, while the document is not active.
    other = w.create_new_document()
    attempt("active is other", lambda: w.active_document.name.get())
    end = d.text_object.end_of_content.get()
    d.create_range(start=end - 1, end_=end - 1).content.set("\r")
    n = d.count(each=k.paragraph)
    s = d.paragraphs[n].text_object.start_of_content.get()
    attempt("make table in inactive doc", lambda: d.make(
        new=k.table, at=d.create_range(start=s, end_=s),
        with_properties={k.number_of_rows: 1, k.number_of_columns: 2}))
    attempt("tables count", lambda: d.count(each=k.table))
    attempt("delete table 1", lambda: d.tables[1].delete())
    attempt("paragraphs after delete", lambda: texts(d))
finally:
    if other is not None:
        other.close(saving=k.no)
    d.close(saving=k.no)
    print("closed")
