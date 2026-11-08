"""Equations: from LaTeX or UnicodeMath, built by Word, numbered by a field.

How an equation is made, all measured (docs/PROBES.md, "Equations"):

1. Its UnicodeMath (LaTeX is converted first, tools/latex.py) goes into a
   paragraph of its own, and Word makes it a math zone and builds it up.
2. A numbered equation is written "...#(1)". Word builds that into an
   equation centred at display size with its number at the right margin, as
   Word does when a user types it. The "1" is then replaced by a SEQ field,
   so the numbers count themselves: (1), (2), (3), and they renumber when an
   equation is added above.
3. What Word built is read back as UnicodeMath (Linearize, read, BuildUp).
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, field_validator, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, Input, annotations, call, follow, mac_impl
from ._word import clean_text, paragraph
from .latex import LatexError, latex_to_unicodemath

WD_FIELD_SEQUENCE = 12  # WdFieldType wdFieldSequence (reference/equations.md)
NUMBER_MARK = "#(1)"


class InsertEquationInput(Input):
    equation: str = Field(..., min_length=1, max_length=4000, description=(
        "The equation, in LaTeX (\\frac{a}{b}, \\sum_{i=1}^n, \\begin{pmatrix}...) "
        "or UnicodeMath (a/b, ∑_(i=1)^n▒, ■(a&b@c&d)), as `format` says. "
        "Without $ signs, or with them: they are taken off."))
    format: Literal["latex", "unicodemath"] = Field("latex", description=(
        "How `equation` is written. LaTeX is converted to UnicodeMath, which "
        "Word builds; the result says what it was converted to."))
    numbered: bool = Field(False, description=(
        "Number the equation (1), (2)... at the right margin, counted by a SEQ "
        "field so the numbers stay in order when equations are added."))
    sequence: str = Field("Equation", pattern=r"^[A-Za-z][A-Za-z0-9_]{0,39}$", description=(
        "The SEQ name the number counts in. Equations in different sequences "
        "are numbered separately."))
    after: int | None = Field(None, ge=0, description=(
        "Insert after this paragraph (0 = at the very start)."))
    before: int | None = Field(None, ge=1, description="Insert before this paragraph.")

    @model_validator(mode="after")
    def _one_place(self):
        if self.after is not None and self.before is not None:
            raise ValueError("Give `after` or `before`, not both. Leave both out to add "
                             "at the end.")
        return self

    @field_validator("equation")
    @classmethod
    def _no_paragraph_marks(cls, value: str) -> str:
        if "\r" in value or "\n" in value:
            raise ValueError("An equation is one line. For several aligned lines use "
                             "\\begin{aligned} ... \\\\ ... \\end{aligned} in LaTeX, or "
                             "█(a&=b@&=c) in UnicodeMath.")
        return value


def to_unicodemath(equation: str, fmt: str) -> str:
    if fmt == "unicodemath":
        return equation.strip()
    try:
        return latex_to_unicodemath(equation)
    except LatexError as exc:
        raise ValueError(f"{exc}. Nothing was inserted.") from None


def _last_one(char_at, start: int, end: int) -> int | None:
    """Word's position of the last "1" in [start, end): the number of "#(1)".

    Asked of Word one character at a time from the end, rather than worked
    out from the paragraph's text: after a root is built up, positions in
    the text and in Word no longer agree (seen: the number of an equation
    holding √π lost its parentheses when the field went one place off).
    """
    for pos in range(end - 1, max(start, end - 40) - 1, -1):
        if char_at(pos) == "1":
            return pos
    return None


def _seq_fields(doc, sequence: str):
    """The document's SEQ fields counting in `sequence`, in order."""
    out = []
    fields = doc.Fields
    for i in range(1, fields.Count + 1):
        f = fields.Item(i)
        code = (f.Code.Text or "").split()
        if len(code) >= 2 and code[0].upper() == "SEQ" and code[1] == sequence:
            out.append(f)
    return out


def _insert_equation_impl(equation, fmt, numbered, sequence, after, before):
    from . import text

    linear = to_unicodemath(equation, fmt)
    # In Normal, whatever paragraph it lands next to (a heading's style would
    # otherwise carry over).
    written = text._insert_paragraphs_impl(
        [{"text": linear + (NUMBER_MARK if numbered else ""), "style": "Normal"}], after,
        before)
    index = written["inserted"][0]["index"]
    doc = backend.word.document()
    p = paragraph(doc, index)
    rng = p.Range
    zone = doc.OMaths.Add(doc.Range(rng.Start, rng.End - 1))
    math = zone.OMaths.Item(1)
    math.BuildUp()
    number = None
    warnings = []
    if numbered:
        rng = paragraph(doc, index).Range
        start = _last_one(lambda pos: doc.Range(pos, pos + 1).Text, rng.Start, rng.End)
        if start is None:
            warnings.append("The equation number could not be found to make it a field; "
                            "it stays (1).")
        else:
            field = doc.Fields.Add(doc.Range(start, start + 1), WD_FIELD_SEQUENCE,
                                   sequence, False)
            for f in _seq_fields(doc, sequence):
                f.Update()
            number = field.Result.Text
    math = paragraph(doc, index).Range.OMaths.Item(1)
    math.Linearize()
    built = clean_text(paragraph(doc, index).Range.Text)
    math = paragraph(doc, index).Range.OMaths.Item(1)
    math.BuildUp()
    follow(doc, paragraph(doc, index).Range)
    result = {"success": True, "index": index, "unicodemath": linear,
              "word_reads_it_as": built}
    if numbered:
        result["number"] = number
    if warnings:
        result["warnings"] = warnings
    return result


def _list_equations_impl():
    doc = backend.word.document()
    maths = doc.OMaths
    out = []
    numbers = {}
    for f in doc.Fields:
        code = (f.Code.Text or "").split()
        if code and code[0].upper() == "SEQ":
            numbers.setdefault(f.Code.Start, (code[1] if len(code) > 1 else "",
                                              f.Result.Text))
    for i in range(1, maths.Count + 1):
        m = maths.Item(i)
        rng = m.Range
        entry = {"index": doc.Range(0, rng.Start + 1).Paragraphs.Count,
                 "display": m.Type == 0}
        inside = [v for k, v in numbers.items() if rng.Start <= k <= rng.End]
        if inside:
            entry["sequence"], entry["number"] = inside[0]
        out.append(entry)
    return {"success": True, "equations": out}


class GetEquationInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the equation is in.")


def _get_equation_impl(index):
    doc = backend.word.document()
    p = paragraph(doc, index)
    if p.Range.OMaths.Count == 0:
        raise ValueError(f"Paragraph {index} holds no equation. word_list_equations "
                         "lists them.")
    math = p.Range.OMaths.Item(1)
    math.Linearize()
    linear = clean_text(paragraph(doc, index).Range.Text)
    paragraph(doc, index).Range.OMaths.Item(1).BuildUp()
    return {"success": True, "index": index, "unicodemath": linear,
            "note": "Read by turning the equation into its linear form and back, which "
                    "Word records as two undo steps."}


# -- Mac -------------------------------------------------------------------------------

@mac_impl(_insert_equation_impl)
def _mac_insert_equation(equation, fmt, numbered, sequence, after, before):
    from appscript import k

    from . import mac as m
    from . import text
    from ._common import MAC_IMPLS

    linear = to_unicodemath(equation, fmt)
    insert = MAC_IMPLS[text._insert_paragraphs_impl.__name__]
    written = insert([{"text": linear + (NUMBER_MARK if numbered else ""),
                       "style": "Normal"}], after, before)
    index = written["inserted"][0]["index"]
    word = backend.word
    doc = word.document()
    app = word.get_app()
    start, end = m._bounds(doc, index)
    # Measured: both `from range` and `in document` are needed.
    app.create_new_equation(from_range=doc.create_range(start=start, end_=end - 1),
                            in_document=doc)
    count = int(doc.count(each=k.math_object))
    target = None
    for j in range(1, count + 1):
        rng = doc.math_objects[j].text_range
        if start <= rng.start_of_content.get() <= end:
            target = doc.math_objects[j]
    if target is None:
        raise RuntimeError("Word answered without error, but no equation is in paragraph "
                           f"{index}.")
    target.build_up()
    number, warnings = None, []
    if numbered:
        pstart, pend = m._bounds(doc, index)
        content = doc.paragraphs[index].text_object.content.get()
        pos = _last_one(lambda at: doc.create_range(start=at, end_=at + 1).content.get(),
                        pstart, pend)
        if pos is None:
            warnings.append("The equation number could not be found to make it a field; "
                            "it stays (1).")
        else:
            app.create_new_field(text_range=doc.create_range(start=pos, end_=pos + 1),
                                 field_type=k.field_sequence, field_text=sequence,
                                 preserve_formatting=False)
            for j in range(1, int(doc.count(each=k.field)) + 1):
                f = doc.fields[j]
                code = (f.field_code.content.get() or "").split()
                if len(code) >= 2 and code[0].upper() == "SEQ" and code[1] == sequence:
                    f.update_field()
                    s0 = f.field_code.start_of_content.get()
                    if pstart <= s0 <= pstart + len(content) + 4:
                        number = f.result_range.content.get()
    target.linearize()
    built = m._clean(doc.paragraphs[index].text_object.content.get())
    target.build_up()
    m._scroll(doc, *m._bounds(doc, index))
    result = {"success": True, "index": index, "unicodemath": linear,
              "word_reads_it_as": built,
              "note": "On macOS each step is its own undo step."}
    if numbered:
        result["number"] = number
    if warnings:
        result["warnings"] = warnings
    return result


@mac_impl(_list_equations_impl)
def _mac_list_equations():
    from appscript import k

    doc = backend.word.document()
    out = []
    for j in range(1, int(doc.count(each=k.math_object)) + 1):
        mo = doc.math_objects[j]
        start = mo.text_range.start_of_content.get()
        index = int(doc.create_range(start=0, end_=start + 1).count(each=k.paragraph))
        out.append({"index": index,
                    "display": "inline" not in str(mo.display_type.get())})
    return {"success": True, "equations": out}


@mac_impl(_get_equation_impl)
def _mac_get_equation(index):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    m._para(doc, index)
    start, end = m._bounds(doc, index)
    for j in range(1, int(doc.count(each=k.math_object)) + 1):
        mo = doc.math_objects[j]
        if start <= mo.text_range.start_of_content.get() < end:
            mo.linearize()
            linear = m._clean(doc.paragraphs[index].text_object.content.get())
            mo.build_up()
            return {"success": True, "index": index, "unicodemath": linear,
                    "note": "Read by turning the equation into its linear form and back."}
    raise ValueError(f"Paragraph {index} holds no equation. word_list_equations lists them.")


def register(mcp) -> None:
    @mcp.tool(name="word_insert_equation", annotations=annotations("Insert an equation",
                                                                   WRITES))
    async def word_insert_equation(params: InsertEquationInput) -> str:
        """Insert a display equation, built by Word's own equation editor, from
        LaTeX or UnicodeMath, optionally numbered (1), (2)... at the right margin.

        The number is a SEQ field, so equations renumber themselves. The result
        gives the UnicodeMath Word built and reads back what Word made of it;
        compare the two to check the equation.
        """
        p = params
        return await call(_insert_equation_impl, p.equation, p.format, p.numbered,
                          p.sequence, p.after, p.before, undo="Insert equation")

    @mcp.tool(name="word_list_equations", annotations=annotations("List equations",
                                                                  READ_ONLY))
    async def word_list_equations() -> str:
        """The document's equations: their paragraph, whether each is a display
        equation, and its number if it has one."""
        return await call(_list_equations_impl)

    @mcp.tool(name="word_get_equation", annotations=annotations("Read an equation", WRITES))
    async def word_get_equation(params: GetEquationInput) -> str:
        """An equation as UnicodeMath, read by turning it into its linear form
        and back (two undo steps in Word)."""
        return await call(_get_equation_impl, params.index)
