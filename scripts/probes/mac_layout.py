"""Pictures, captions, references, headers, vertical text and ruby on Word for Mac.

    uv run python scripts/probes/mac_layout.py

A scratch document, closed unsaved. The picture is written in .word-test/.
"""

from pathlib import Path

import appscript
from appscript import k

from ms_word_mcp.backend.mac import _decode_wide_integers
from ms_word_mcp.tools.screen import png

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")


def attempt(label, f):
    try:
        print(f"{label} -> {f()!r}")
    except Exception as e:  # noqa: BLE001
        print(f"{label} FAILED: {' '.join(x.strip() for x in str(e).splitlines()[:3])[:200]}")


# Word asks the user for access to each file outside its own container, and
# the call times out waiting (measured), so the picture goes inside it.
folder = Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents"
picture = folder / "ms-word-mcp-probe-picture.png"
picture.write_bytes(png(40, 20, [b"\xc0\x40\x40" * 40] * 20))


def para(d, i):
    return d.paragraphs[i].text_object.content.get()


def bounds(d, i):
    t = d.paragraphs[i].text_object
    return t.start_of_content.get(), t.end_of_content.get()


d = w.create_new_document()
other = None
try:
    d.create_range(start=0, end_=0).content.set("Intro\r\rFigure X: A red box\rSee here.\rHeading text\rLast")
    d.paragraphs[5].style.set(k.style_heading1)
    s, _ = bounds(d, 2)
    other = w.create_new_document()
    on = other.name.get()
    attempt("picture, other document active", lambda: d.make(
        new=k.inline_picture, at=d.create_range(start=s, end_=s),
        with_properties={k.file_name: str(picture), k.link_to_file: False,
                         k.save_with_document: True}))
    attempt("inline shapes", lambda: d.count(each=k.inline_shape))
    attempt("active still other", lambda: w.active_document.name.get() == on)
    if d.count(each=k.inline_shape) == 0:
        d.activate_object()
        attempt("picture, document active", lambda: d.make(
            new=k.inline_picture, at=d.create_range(start=s, end_=s),
            with_properties={k.file_name: str(picture), k.link_to_file: False,
                             k.save_with_document: True}))
        other.activate_object()
    shape = d.inline_shapes[1]
    attempt("size", lambda: (shape.width.get(), shape.height.get(), shape.lock_aspect_ratio.get()))
    attempt("set width only", lambda: (shape.width.set(100), shape.width.get(), shape.height.get()))
    attempt("alt text", lambda: (shape.alternative_text.set("A red box"), shape.alternative_text.get()))
    attempt("type", lambda: shape.inline_shape_type.get())
    attempt("range", lambda: (shape.text_object.start_of_content.get(), shape.text_object.end_of_content.get()))
    attempt("para 2 text", lambda: para(d, 2))
    # A caption by hand: the X replaced by a SEQ field.
    s3, e3 = bounds(d, 3)
    attempt("SEQ field over X", lambda: w.create_new_field(
        text_range=d.create_range(start=s3 + 7, end_=s3 + 8), field_type=k.field_empty,
        field_text="SEQ Figure \\* ARABIC", preserve_formatting=False))
    attempt("active after field", lambda: w.active_document.name.get() == on)
    other.activate_object()
    attempt("para 3 text", lambda: para(d, 3))
    f = d.fields[1]
    attempt("field result", lambda: (f.result_range.content.get(), f.result_range.start_of_content.get(),
                                     f.result_range.end_of_content.get()))
    attempt("field code start", lambda: f.field_code.start_of_content.get())
    attempt("caption style", lambda: (d.paragraphs[3].style.set(k.style_caption), d.paragraphs[3].style.name_local.get()))
    attempt("items for 'Figure'", lambda: d.get_cross_reference_items(reference_type="Figure"))
    end = f.result_range.end_of_content.get()
    attempt("bookmark over Figure 1", lambda: d.make(new=k.bookmark, at=d, with_properties={
        k.name: "_Ref100000001", k.text_object: d.create_range(start=s3, end_=end + 1)}))
    attempt("bookmark text", lambda: d.bookmarks["_Ref100000001"].text_object.content.get())
    attempt("visible bookmarks", lambda: d.count(each=k.bookmark))
    s4, e4 = bounds(d, 4)
    attempt("REF field", lambda: w.create_new_field(
        text_range=d.create_range(start=s4 + 4, end_=s4 + 8), field_type=k.field_empty,
        field_text="REF _Ref100000001 \\h", preserve_formatting=False))
    other.activate_object()
    attempt("para 4 text", lambda: para(d, 4))
    attempt("fields", lambda: [(x.field_code.content.get(), x.result_range.content.get())
                               for x in [d.fields[1], d.fields[2]]])
    # Headers and footers
    sec = d.sections[1]
    attempt("header", lambda: sec.get_header(index=k.header_footer_primary))
    hdr = sec.get_header(index=k.header_footer_primary)
    ftr = sec.get_footer(index=k.header_footer_primary)
    attempt("header set", lambda: (hdr.text_object.content.set("Report title"), hdr.text_object.content.get()))
    attempt("header alignment", lambda: (hdr.text_object.paragraph_format.alignment.set(k.align_paragraph_right),
                                         hdr.text_object.paragraph_format.alignment.get()))
    attempt("footer placeholder", lambda: ftr.text_object.content.set("Page X of Y"))
    attempt("footer chars", lambda: ftr.text_object.count(each=k.character))
    attempt("field over char 11 (Y)", lambda: w.create_new_field(
        text_range=ftr.text_object.characters[11], field_type=k.field_empty, field_text="NUMPAGES",
        preserve_formatting=False))
    attempt("field over char 6 (X)", lambda: w.create_new_field(
        text_range=ftr.text_object.characters[6], field_type=k.field_empty, field_text="PAGE",
        preserve_formatting=False))
    other.activate_object()
    attempt("footer text", lambda: ftr.text_object.content.get())
    attempt("footer fields", lambda: ftr.text_object.count(each=k.field))
    attempt("page numbers props", lambda: (ftr.page_number_options.starting_number.get(),
                                           ftr.page_number_options.restart_numbering_at_section.get(),
                                           ftr.page_number_options.number_style.get()))
    attempt("page number style set", lambda: (ftr.page_number_options.number_style.set(k.page_number_style_lowercase_roman),
                                              ftr.page_number_options.number_style.get()))
    attempt("restart at 5", lambda: (ftr.page_number_options.restart_numbering_at_section.set(True),
                                     ftr.page_number_options.starting_number.set(5),
                                     ftr.page_number_options.starting_number.get()))
    attempt("different first page", lambda: (sec.page_setup.different_first_page_header_footer.set(True),
                                             sec.page_setup.different_first_page_header_footer.get()))
    first = sec.get_header(index=k.header_footer_first_page)
    attempt("first page header", lambda: (first.text_object.content.set("First"), first.text_object.content.get()))
    attempt("odd even", lambda: (sec.page_setup.odd_and_even_pages_header_footer.set(True),
                                 sec.page_setup.odd_and_even_pages_header_footer.get()))
    attempt("main text unchanged", lambda: para(d, 1))
    # Vertical text
    ps = sec.page_setup
    attempt("orientation now", lambda: (ps.orientation.get(), ps.page_width.get(), ps.page_height.get()))
    attempt("text orientation of the section", lambda: sec.text_object.orientation.get())
    attempt("set vertical", lambda: (sec.text_object.orientation.set(k.text_orientation_vertical),
                                              sec.text_object.orientation.get()))
    attempt("page after", lambda: (ps.orientation.get(), ps.page_width.get(), ps.page_height.get()))
    attempt("grid", lambda: (ps.layout_mode.get(), ps.chars_line.get(), ps.lines_page.get()))
    attempt("set layout grid", lambda: (ps.layout_mode.set(k.layout_mode_grid), ps.chars_line.set(40), ps.lines_page.set(30),
                                        ps.layout_mode.get(), ps.chars_line.get(), ps.lines_page.get()))
    attempt("set horizontal", lambda: (sec.text_object.orientation.set(k.text_orientation_horizontal),
                                       sec.text_object.orientation.get()))
    # Ruby
    s1, e1 = bounds(d, 1)
    d.create_range(start=e1 - 1, end_=e1 - 1).content.set(" 漢字です")
    s1, e1 = bounds(d, 1)
    at = e1 - 5
    attempt("target", lambda: d.create_range(start=at, end_=at + 2).content.get())
    attempt("phonetic guide", lambda: d.create_range(start=at, end_=at + 2).phonetic_guide(
        text="かんじ", alignment_type=k.phonetic_guide_alignment_center, raise_=0, font_size=5, font_name=""))
    attempt("para 1 after ruby", lambda: para(d, 1))
    attempt("fields after ruby", lambda: [x.field_code.content.get() for x in d.fields.get()][:5])
    attempt("active still other at the end", lambda: w.active_document.name.get() == on)
finally:
    if other is not None:
        attempt("close other", lambda: other.close(saving=k.no))
    attempt("close", lambda: d.close(saving=k.no))
    picture.unlink(missing_ok=True)
    print("closed")
