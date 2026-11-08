"""How long reading paragraphs takes on Word for Mac, one by one and in bulk.

    uv run python scripts/probes/mac_bulk_reads.py

Makes a scratch document of 1,000 paragraphs (every tenth a heading, a
table in the middle), times the reads, and closes it unsaved.
"""

import time

import appscript
from appscript import k

from ms_word_mcp.backend.mac import _decode_wide_integers

_decode_wide_integers()
w = appscript.app(id="com.microsoft.Word", terms="sdef")


def timed(label, f):
    t = time.perf_counter()
    try:
        out = f()
        print(f"{label}: {time.perf_counter() - t:.2f} s -> {repr(out)[:160]}")
        return out
    except Exception as e:  # noqa: BLE001
        print(f"{label}: FAILED after {time.perf_counter() - t:.2f} s: "
              f"{' '.join(str(e).splitlines()[:2])[:200]}")


d = w.create_new_document()
try:
    lines = [f"Heading {i}" if i % 10 == 0 else f"Body paragraph {i} with some text."
             for i in range(1, 1001)]
    d.create_range(start=0, end_=0).content.set("\r".join(lines))
    n = d.count(each=k.paragraph)
    for i in range(10, 1001, 10):
        d.paragraphs[i].style.set(k.style_heading1)
    s = d.paragraphs[500].text_object.start_of_content.get()
    d.make(new=k.table, at=d.create_range(start=s, end_=s),
           with_properties={k.number_of_rows: 2, k.number_of_columns: 2})
    n = d.count(each=k.paragraph)
    print("paragraphs", n)
    timed("one by one: outline level of 200", lambda: [
        d.paragraphs[i].outline_level.get() for i in range(1, 201)])
    timed("bulk: every outline level", lambda: len(d.paragraphs.outline_level.get()))
    timed("bulk: every paragraph's text", lambda: len(d.paragraphs.text_object.content.get()))
    timed("bulk: every style name", lambda: d.paragraphs.style.name_local.get()[:3])
    timed("bulk: every style (references)", lambda: len(d.paragraphs.style.get()))
    timed("bulk: whole text", lambda: len(d.text_object.content.get()))
    timed("bulk: alignment", lambda: d.paragraphs.paragraph_format.alignment.get()[:3])
    timed("bulk: bold", lambda: d.paragraphs.text_object.bold.get()[:3])
    timed("bulk: latin font", lambda: d.paragraphs.text_object.font_object.ascii_name.get()[:3])
    timed("range of paragraphs 300-500 texts", lambda: len(
        d.paragraphs[300:500].text_object.content.get()))
    timed("range of paragraphs 300-500 styles", lambda: d.paragraphs[300:500]
          .style.name_local.get()[:2])
    timed("starts of every paragraph", lambda: d.paragraphs.text_object.start_of_content
          .get()[:3])
finally:
    d.close(saving=k.no)
    print("closed")
