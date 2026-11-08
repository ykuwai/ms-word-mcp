"""Measure equations on Word for Mac, in a scratch document closed unsaved.

    uv run python scripts/probes/mac_equations.py
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
        print(f"{label} FAILED: {' '.join(x.strip() for x in lines[:3])[:160]}")


d = w.create_new_document()
try:
    cases = [("unicodemath fraction", "a/b"), ("power and root", "E=mc^2+√(x^2+1)"),
             ("integral", "∫_0^∞ e^(-x^2) dx=√π/2"), ("tex words", "\\alpha+\\beta=\\gamma"),
             ("latex frac", "\\frac{a}{b}"), ("numbered with #", "E=mc^2#(1)")]
    for i, (label, linear) in enumerate(cases, start=1):
        end = d.text_object.end_of_content.get()
        start = end - 1
        prefix = "" if i == 1 else "\r"
        d.create_range(start=start, end_=start).content.set(prefix + linear)
        p = d.paragraphs[i]
        s, e = p.text_object.start_of_content.get(), p.text_object.end_of_content.get()
        rng = d.create_range(start=s, end_=e - 1)
        attempt(f"[{label}] create new equation", lambda: w.create_new_equation(from_range=rng, in_document=d))
        attempt(f"[{label}] math objects in doc", lambda: d.count(each=k.math_object))
        n = d.count(each=k.math_object)
        m = d.math_objects[n]
        attempt(f"[{label}] build up", lambda: m.build_up())
        attempt(f"[{label}] text after build up", lambda: d.paragraphs[i].text_object.content.get())
        attempt(f"[{label}] display type", lambda: m.display_type.get())
        attempt(f"[{label}] linearize", lambda: m.linearize())
        attempt(f"[{label}] text after linearize", lambda: d.paragraphs[i].text_object.content.get())
        attempt(f"[{label}] build up again", lambda: m.build_up())
    # SEQ field
    end = d.text_object.end_of_content.get()
    d.create_range(start=end - 1, end_=end - 1).content.set("\r(X)")
    n = d.count(each=k.paragraph)
    s = d.paragraphs[n].text_object.start_of_content.get()
    attempt("SEQ field", lambda: w.create_new_field(text_range=d.create_range(start=s + 1, end_=s + 2),
                                                    field_type=k.field_sequence, field_text="Equation",
                                                    preserve_formatting=False))
    attempt("SEQ paragraph text", lambda: d.paragraphs[n].text_object.content.get())
finally:
    d.close(saving=k.no)
    print("closed")
