"""Table calls on Word for Mac while another document is the active one, and
a table merged across rows. Scratch documents, closed unsaved.

    uv run python scripts/probes/mac_tables_inactive.py
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
        print(f"{label} FAILED: {' '.join(x.strip() for x in str(e).splitlines()[:3])[:200]}")


d = w.create_new_document()
dn = d.name.get()
d = w.documents[dn]
other = None
try:
    d.make(new=k.table, at=d.create_range(start=0, end_=0),
           with_properties={k.number_of_rows: 3, k.number_of_columns: 3})
    tb = d.tables[1]
    for r in (1, 2, 3):
        for c in (1, 2, 3):
            tb.get_cell_from_table(row=r, column=c).text_object.content.set(f"r{r}c{c}")
    other = w.create_new_document()
    on = other.name.get()
    attempt("active is the other document", lambda: w.active_document.name.get() == on)
    attempt("set cell text", lambda: tb.get_cell_from_table(row=1, column=1)
            .text_object.content.set("R1C1"))
    attempt("make new row after row 3", lambda: d.make(new=k.row, at=tb.rows[3].after))
    attempt("make new row before row 1", lambda: d.make(new=k.row, at=tb.rows[1].before))
    attempt("rows", lambda: tb.number_of_rows.get())
    attempt("delete row 1", lambda: tb.rows[1].delete())
    attempt("heading format", lambda: tb.rows[1].heading_format.set(True))
    attempt("style by enum", lambda: tb.style.set(k.style_table_light_grid))
    attempt("style by local name", lambda: (tb.style.set(d.Word_styles["表 (格子)"]),
                                            tb.style.name_local.get()))
    attempt("enable borders", lambda: tb.border_options.enable_borders.set(True))
    attempt("auto fit window", lambda: tb.auto_fit_behavior(behavior=k.auto_fit_window))
    attempt("column width", lambda: (tb.columns[1].preferred_width_type.set(
        k.preferred_width_points), tb.columns[1].preferred_width.set(90)))
    attempt("row alignment", lambda: tb.row_options.alignment.set(k.align_row_center))
    cell = tb.get_cell_from_table(row=2, column=2)
    attempt("shading color", lambda: cell.shading.background_pattern_color.set([65535, 0, 0]))
    attempt("shading back to auto by index", lambda: (
        cell.shading.background_pattern_color_index.set(k.auto),
        cell.shading.background_pattern_color.get()))
    attempt("bold/italic/color/size", lambda: (
        cell.text_object.bold.set(True), cell.text_object.italic.set(True),
        cell.text_object.font_object.color.set([0, 0, 65535]),
        cell.text_object.font_object.font_size.set(14)))
    attempt("delete column 3", lambda: tb.columns[3].delete())
    attempt("merge 1,1 to 2,2 (across rows)", lambda: tb.get_cell_from_table(row=1, column=1)
            .merge_cell(with_=tb.get_cell_from_table(row=2, column=2)))
    attempt("active is still the other document", lambda: w.active_document.name.get() == on)
    attempt("size", lambda: (tb.number_of_rows.get(), tb.number_of_columns.get(),
                             tb.uniform.get()))
    for r in (1, 2, 3, 4):
        attempt(f"row {r} cells", lambda: tb.rows[r].count(each=k.cell))
        attempt(f"row {r} heading", lambda: tb.rows[r].heading_format.get())
    attempt("table text", lambda: tb.text_object.content.get())
    attempt("cell 2,1", lambda: tb.get_cell_from_table(row=2, column=1).text_object.content.get())
    attempt("delete row 3 after merge", lambda: tb.rows[3].delete())
    attempt("delete table", lambda: tb.delete())
finally:
    if other is not None:
        other.close(saving=k.no)
    d.close(saving=k.no)
    print("closed")


# A cell merged down a column, in a 3 x 3 table.
d = w.create_new_document()
try:
    d.make(new=k.table, at=d.create_range(start=0, end_=0),
           with_properties={k.number_of_rows: 3, k.number_of_columns: 3})
    tb = d.tables[1]
    for r in (1, 2, 3):
        for c in (1, 2, 3):
            tb.get_cell_from_table(row=r, column=c).text_object.content.set(f"r{r}c{c}")
    fresh = tb.get_cell_from_table(row=3, column=3)
    attempt("shading of a cell never shaded", lambda: (
        fresh.shading.background_pattern_color.get(),
        fresh.shading.background_pattern_color_index.get()))
    attempt("merge 1,1 with 2,1", lambda: tb.get_cell_from_table(row=1, column=1)
            .merge_cell(with_=tb.get_cell_from_table(row=2, column=1)))
    attempt("size", lambda: (tb.number_of_rows.get(), tb.number_of_columns.get(),
                             tb.uniform.get()))
    for r in (1, 2, 3):
        attempt(f"row {r} cells", lambda: tb.rows[r].count(each=k.cell))
    attempt("table text", lambda: tb.text_object.content.get())
    attempt("cell 2,1", lambda: tb.get_cell_from_table(row=2, column=1).text_object.content.get())
    attempt("heading row 1", lambda: (tb.rows[1].heading_format.set(True),
                                      tb.rows[1].heading_format.get()))
    attempt("make new row after row 3", lambda: d.make(new=k.row, at=tb.rows[3].after))
    attempt("delete row 2", lambda: tb.rows[2].delete())
    attempt("table text after", lambda: tb.text_object.content.get())
finally:
    d.close(saving=k.no)
    print("closed")
