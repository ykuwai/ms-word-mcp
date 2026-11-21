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

from .fields import may_update

from typing import Literal

from pydantic import Field, field_validator, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, Input, Trimmed, annotations, call, follow, mac_impl, mac_with_active
from ._word import clean_text, paragraph
from .latex import LatexError, latex_to_unicodemath
from .references import ChapterNumber, chapter_codes, chapter_warning, joined_number

WD_FIELD_SEQUENCE = 12  # WdFieldType wdFieldSequence (reference/equations.md)
NUMBER_MARK = "#(1)"


class InsertEquationInput(Input):
    equation: Trimmed = Field(..., min_length=1, max_length=4000, description=(
        "The equation, in LaTeX (\\frac{a}{b}, \\sum_{i=1}^n, \\begin{pmatrix}...) "
        "or UnicodeMath (a/b, ∑_(i=1)^n▒, ■(a&b@c&d)), as `format` says. "
        "Without $ signs, or with them: they are taken off."))
    format: Literal["latex", "unicodemath"] = Field("latex", description=(
        "How `equation` is written. LaTeX is converted to UnicodeMath, which "
        "Word builds; the result says what it was converted to."))
    numbered: bool = Field(False, description=(
        "Number the equation (1), (2)... at the right margin, counted by a SEQ "
        "field so the numbers stay in order when equations are added."))
    sequence: Trimmed = Field("Equation", pattern=r"^[A-Za-z][A-Za-z0-9_]{0,39}$", description=(
        "The SEQ name the number counts in. Equations in different sequences "
        "are numbered separately."))
    after: int | None = Field(None, ge=0, description=(
        "Insert after this paragraph (0 = at the very start)."))
    before: int | None = Field(None, ge=1, description="Insert before this paragraph.")
    chapter: ChapterNumber | None = Field(None, description=(
        "With `numbered`: number with the chapter, (2.1), counting again in each chapter. "
        "The chapter's number is that of the Heading `level` before the equation, so the "
        "headings need numbers (word_apply_list with levels linked to the heading styles)."))

    @model_validator(mode="after")
    def _one_place(self):
        if self.chapter is not None and not self.numbered:
            raise ValueError("`chapter` numbers a numbered equation: give `numbered` too.")
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


DISPLAY_PARAGRAPH = ("Paragraph {index} is a display equation: an inline equation goes in "
                     "a paragraph of text. Put it in the paragraph before or after, or "
                     "change this one with word_set_equation.")
NO_SUCH_MATH = ("Paragraph {index} holds {count} equation(s), not {nth}. "
                "word_list_equations lists them.")


def inline_answer(index, nth, display, linear, built) -> dict:
    out = {"success": True, "index": index, "nth": nth, "inline": not display,
           "unicodemath": linear, "word_reads_it_as": built}
    if display:
        out["warnings"] = [
            "Word made it a display equation, on a line of its own: the paragraph holds "
            "nothing else. An inline equation sits among the paragraph's words."]
    return out


def nth_of(starts: list[int], at: int) -> int:
    """Which of a paragraph's equations, by their starts, begins at or after
    `at` first (1-based)."""
    order = sorted(starts)
    return next((i for i, s in enumerate(order, 1) if s >= at), len(order))


def number_mark(chapter) -> str:
    """What a numbered equation is written with: "#(1)", or "#(1.1)" with its
    chapter's, each 1 made a field afterwards. (Quoting the separator as
    ordinary text left the equation in its linear form: measured on the Mac.)"""
    return f"#(1{chapter['separator']}1)" if chapter else NUMBER_MARK


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


def _insert_equation_impl(equation, fmt, numbered, sequence, after, before, chapter=None):
    from . import text

    linear = to_unicodemath(equation, fmt)
    # In Normal, whatever paragraph it lands next to (a heading's style would
    # otherwise carry over).
    written = text._insert_paragraphs_impl(
        [{"text": linear + (number_mark(chapter) if numbered else ""), "style": "Normal"}],
        after, before)
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
            seq_text = sequence if not chapter else chapter_codes(
                sequence, chapter["level"])[1].removeprefix("SEQ ")
            field = doc.Fields.Add(doc.Range(start, start + 1), WD_FIELD_SEQUENCE,
                                   seq_text, False)
            chap = None
            if chapter:
                # The chapter's "1" is the one before the count's field.
                begin = field.Code.Start - 1
                at = _last_one(lambda pos: doc.Range(pos, pos + 1).Text, begin - 6, begin)
                if at is None:
                    warnings.append("The chapter's place in the number could not be found; "
                                    "the number counts without it.")
                else:
                    chap = doc.Fields.Add(doc.Range(at, at + 1), -1,
                                          chapter_codes(sequence, chapter["level"])[0], False)
                    chap.Update()
            for f in _seq_fields(doc, sequence):
                if may_update(f.Code.Text):
                    f.Update()
            number = field.Result.Text
            if chap is not None:
                between = doc.Range(chap.Result.End + 1, field.Code.Start - 1).Text or ""
                number = joined_number(chap.Result.Text, between, number)
                warnings += chapter_warning(chap.Result.Text, chapter["level"])
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


def _math(doc, index, nth):
    maths = paragraph(doc, index).Range.OMaths
    if nth > maths.Count:
        if not maths.Count:
            raise ValueError(f"Paragraph {index} holds no equation. word_list_equations "
                             "lists them.")
        raise ValueError(NO_SUCH_MATH.format(index=index, count=maths.Count, nth=nth))
    return maths.Item(nth)


def _insert_inline_equation_impl(index, equation, fmt, position, match):
    from ._word import locate_spans

    linear = to_unicodemath(equation, fmt)
    doc = backend.word.document()
    rng = paragraph(doc, index).Range
    maths = rng.OMaths
    if any(maths.Item(i).Type == 0 for i in range(1, maths.Count + 1)):
        raise ValueError(DISPLAY_PARAGRAPH.format(index=index))
    text = rng.Text
    if match is not None:
        if text.find(match) < 0:
            raise ValueError(f"'{match}' is not in paragraph {index}. Nothing was inserted.")
        hidden = max(0, rng.End - rng.Start - len(text.encode("utf-16-le")) // 2)
        s, e = locate_spans(text, [match], lambda a, b: doc.Range(a, b).Text, rng.Start,
                            hidden)[0]
    elif position == "start":
        s = e = rng.Start
    else:
        s = e = rng.End - 1
    target = doc.Range(s, e)
    target.Text = linear
    doc.OMaths.Add(doc.Range(s, s + len(linear.encode("utf-16-le")) // 2))
    starts = [m.Range.Start for m in (paragraph(doc, index).Range.OMaths.Item(i) for i in
              range(1, paragraph(doc, index).Range.OMaths.Count + 1))]
    nth = nth_of(starts, s)
    _math(doc, index, nth).BuildUp()
    math = _math(doc, index, nth)
    display = math.Type == 0
    math.Linearize()
    built = clean_text(_math(doc, index, nth).Range.Text)
    _math(doc, index, nth).BuildUp()
    follow(doc, _math(doc, index, nth).Range)
    return inline_answer(index, nth, display, linear, built)


def _list_equations_impl():
    doc = backend.word.document()
    maths = doc.OMaths
    out = []
    numbers, chapters = {}, []
    for f in doc.Fields:
        code = (f.Code.Text or "").split()
        if code and code[0].upper() == "STYLEREF":
            chapters.append(f)
    for f in doc.Fields:
        code = (f.Code.Text or "").split()
        if code and code[0].upper() == "SEQ":
            begin = f.Code.Start - 1
            chap = next((c for c in chapters if 0 <= begin - (c.Result.End + 1) <= 3), None)
            number = f.Result.Text if chap is None else joined_number(
                chap.Result.Text, doc.Range(chap.Result.End + 1, begin).Text or "",
                f.Result.Text)
            numbers.setdefault(f.Code.Start, (code[1] if len(code) > 1 else "", number))
    for i in range(1, maths.Count + 1):
        m = maths.Item(i)
        rng = m.Range
        index = doc.Range(0, rng.Start + 1).Paragraphs.Count
        entry = {"index": index, "nth": 1 + sum(e["index"] == index for e in out),
                 "display": m.Type == 0}
        inside = [v for k, v in numbers.items() if rng.Start <= k <= rng.End]
        if inside:
            entry["sequence"], entry["number"] = inside[0]
        out.append(entry)
    return {"success": True, "equations": out}


class InsertInlineEquationInput(Input):
    index: int = Field(..., ge=1, description="The paragraph to put the equation in.")
    equation: Trimmed = Field(..., min_length=1, max_length=4000, description=(
        "The equation, in LaTeX or UnicodeMath, as `format` says, without $ signs."))
    format: Literal["latex", "unicodemath"] = Field("latex", description=(
        "How `equation` is written."))
    position: Literal["end", "start"] = Field("end", description=(
        "At the end of the paragraph's text or at its start."))
    match: str | None = Field(None, min_length=1, max_length=255, description=(
        "Instead of `position`: replace the first occurrence of this text in the "
        "paragraph with the equation, as a placeholder written for it (\"where X is\")."))

    _one_line = field_validator("equation")(
        classmethod(lambda cls, v: InsertEquationInput._no_paragraph_marks(v)))


NTH = Field(1, ge=1, le=100, description=(
    "Which equation in the paragraph, counting from 1: an inline equation shares its "
    "paragraph with text and other equations (word_list_equations gives each its `nth`)."))


class GetEquationInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the equation is in.")
    nth: int = NTH


def _get_equation_impl(index, nth=1):
    doc = backend.word.document()
    math = _math(doc, index, nth)
    inline = math.Type != 0
    math.Linearize()
    # An inline equation is read from its own range: the paragraph holds text
    # besides it. A display one from the paragraph, with its number.
    linear = clean_text(_math(doc, index, nth).Range.Text if inline
                        else paragraph(doc, index).Range.Text)
    _math(doc, index, nth).BuildUp()
    return {"success": True, "index": index, "nth": nth, "inline": inline,
            "unicodemath": linear,
            "note": "Read by turning the equation into its linear form and back, which "
                    "Word records as two undo steps."}


class SetEquationInput(Input):
    index: int = Field(..., ge=1, description="The paragraph the equation is in.")
    nth: int = NTH
    equation: Trimmed = Field(..., min_length=1, max_length=4000, description=(
        "The new equation, as word_insert_equation takes it, without its number: a "
        "numbered equation keeps its number."))
    format: Literal["latex", "unicodemath"] = Field("latex", description=(
        "How `equation` is written."))

    @field_validator("equation")
    @classmethod
    def _one_line_no_number(cls, value: str) -> str:
        value = InsertEquationInput._no_paragraph_marks(value)
        if "#(" in value:
            raise ValueError("Give the equation without its number (#(...)): a numbered "
                             "equation keeps the number it has.")
        return value


def kept_number_at(code_start: int, read) -> int | None:
    """Where an equation's own text ends: before the "#(" of its number.

    A numbered equation, linearized, reads "...#(n)" with n a SEQ field; the
    field's begin mark is just before its code, and "#(" just before that.
    `read(a, b)` answers Word's text from a to b. None if it is not there."""
    begin = code_start - 1
    return begin - 2 if read(begin - 2, begin) == "#(" else None


#: A numbered equation, linearized, is an array holding the equation and its
#: number, "█(E=mc^2#(1) )": the new text goes inside it.
ARRAY_OPEN = "█("


def own_text_start(start: int, read) -> int:
    """Where a linearized equation's own text starts: after the array opening
    of a numbered one."""
    return start + len(ARRAY_OPEN) if read(start, start + len(ARRAY_OPEN)) == ARRAY_OPEN else start


def equation_set_answer(index, linear, built, number) -> dict:
    out = {"success": True, "index": index, "unicodemath": linear, "word_reads_it_as": built}
    if number is not None:
        out["number"] = number
    return out


def _set_equation_impl(index, equation, fmt, nth=1):
    linear = to_unicodemath(equation, fmt)
    doc = backend.word.document()
    if _math(doc, index, nth).Type != 0:
        return _set_inline_equation(doc, index, nth, linear)
    paragraph(doc, index).Range.OMaths.Item(1).Linearize()
    math = paragraph(doc, index).Range.OMaths.Item(1).Range
    seq = [f for f in (paragraph(doc, index).Range.Fields.Item(i)
                       for i in range(1, paragraph(doc, index).Range.Fields.Count + 1))
           if (f.Code.Text or "").split()[:1] == ["SEQ"]]
    end = math.End
    # Read now: building the equation up again makes its field anew, and the
    # Field read before is then "deleted" (measured on Windows).
    number = seq[0].Result.Text if seq else None
    if seq:
        end = kept_number_at(seq[0].Code.Start, lambda a, b: doc.Range(a, b).Text)
        if end is None:
            paragraph(doc, index).Range.OMaths.Item(1).BuildUp()
            raise ValueError(f"The number of the equation in paragraph {index} is not where "
                             "it was put. Nothing was changed.")
    begin = own_text_start(math.Start, lambda a, b: doc.Range(a, b).Text) if seq else math.Start
    doc.Range(begin, end).Text = linear
    paragraph(doc, index).Range.OMaths.Item(1).BuildUp()
    # Read back as word_insert_equation does: linear, read, built up again.
    paragraph(doc, index).Range.OMaths.Item(1).Linearize()
    built = clean_text(paragraph(doc, index).Range.Text)
    paragraph(doc, index).Range.OMaths.Item(1).BuildUp()
    follow(doc, paragraph(doc, index).Range)
    return equation_set_answer(index, linear, built, number)


def _set_inline_equation(doc, index, nth, linear):
    """An inline equation's own range rewritten: the paragraph's text around
    it stays."""
    _math(doc, index, nth).Linearize()
    rng = _math(doc, index, nth).Range
    start = rng.Start
    rng.Text = linear
    if paragraph(doc, index).Range.OMaths.Count < nth or \
            _math(doc, index, nth).Range.Start != start:
        # The zone went with its text: made again over the new text.
        doc.OMaths.Add(doc.Range(start, start + len(linear.encode("utf-16-le")) // 2))
    _math(doc, index, nth).BuildUp()
    _math(doc, index, nth).Linearize()
    built = clean_text(_math(doc, index, nth).Range.Text)
    _math(doc, index, nth).BuildUp()
    follow(doc, _math(doc, index, nth).Range)
    return {**equation_set_answer(index, linear, built, None), "nth": nth, "inline": True}


# -- Mac -------------------------------------------------------------------------------

@mac_impl(_insert_equation_impl)
def _mac_insert_equation(equation, fmt, numbered, sequence, after, before, chapter=None):
    from appscript import k

    from . import mac as m
    from . import text
    from ._common import MAC_IMPLS

    linear = to_unicodemath(equation, fmt)
    insert = MAC_IMPLS[text._insert_paragraphs_impl.__name__]
    written = insert([{"text": linear + (number_mark(chapter) if numbered else ""),
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
        # Inside this paragraph only: an equation starting at `end` is the next
        # paragraph's, and building that one up left this one in its linear
        # form (seen in a Sonnet run).
        rng = doc.math_objects[j].text_range
        if start <= rng.start_of_content.get() < end:
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
            seq_text = sequence if not chapter else chapter_codes(
                sequence, chapter["level"])[1].removeprefix("SEQ ")
            mac_with_active(app, doc, lambda: app.create_new_field(
                text_range=doc.create_range(start=pos, end_=pos + 1),
                field_type=k.field_sequence, field_text=seq_text, preserve_formatting=False))
            chap_at = None
            if chapter:
                # The chapter's "1" is the one before the count's field, whose
                # begin mark now stands at `pos`.
                read = lambda at: doc.create_range(start=at, end_=at + 1).content.get()  # noqa: E731
                chap_at = _last_one(read, pos - 6, pos)
                if chap_at is None:
                    warnings.append("The chapter's place in the number could not be found; "
                                    "the number counts without it.")
                else:
                    mac_with_active(app, doc, lambda: app.create_new_field(
                        text_range=doc.create_range(start=chap_at, end_=chap_at + 1),
                        field_type=k.field_empty,
                        field_text=chapter_codes(sequence, chapter["level"])[0],
                        preserve_formatting=False))
            chap = None
            # The paragraph's bounds again, now that its fields' codes are in it: the
            # number is read from the SEQ field inside them, the first. A short
            # equation's paragraph is followed closely by the next one's field, which
            # a margin past its text took for its own (an equation inserted above
            # another read "2", measured on Word for Mac).
            pstart, pend = m._bounds(doc, index)
            for j in range(1, int(doc.count(each=k.field)) + 1):
                f = doc.fields[j]
                code = (f.field_code.content.get() or "").split()
                s0 = f.field_code.start_of_content.get()
                if code and code[0].upper() == "STYLEREF" and chap_at is not None and \
                        chap is None and pstart <= s0 < pend:
                    f.update_field()
                    chap = f
                if (len(code) >= 2 and code[0].upper() == "SEQ" and code[1] == sequence
                        and may_update(f.field_code.content.get())):
                    f.update_field()
                    s0 = f.field_code.start_of_content.get()
                    if number is None and pstart <= s0 < pend:
                        number = f.result_range.content.get()
                        seq_begin = s0 - 1
            if chap is not None and number is not None:
                a = chap.result_range.end_of_content.get() + 1
                between = doc.create_range(start=a, end_=seq_begin).content.get() \
                    if seq_begin > a else ""
                number = joined_number(chap.result_range.content.get(), between or "", number)
                warnings += chapter_warning(chap.result_range.content.get(), chapter["level"])
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
    numbers, chapters, seqs = [], [], []
    for j in range(1, int(doc.count(each=k.field)) + 1):
        f = doc.fields[j]
        code = (f.field_code.content.get() or "").split()
        if code and code[0].upper() == "STYLEREF":
            chapters.append(f)
        elif code and code[0].upper() == "SEQ":
            seqs.append((f, code))
    for f, code in seqs:
        begin = f.field_code.start_of_content.get() - 1
        count = (f.result_range.content.get() or "").strip()
        chap = next((c for c in chapters
                     if 0 <= begin - (c.result_range.end_of_content.get() + 1) <= 3), None)
        if chap is not None:
            a = chap.result_range.end_of_content.get() + 1
            between = doc.create_range(start=a, end_=begin).content.get() if begin > a else ""
            count = joined_number(chap.result_range.content.get(), between or "", count)
        numbers.append((begin + 1, code[1] if len(code) > 1 else "", count))
    out = []
    for j in range(1, int(doc.count(each=k.math_object)) + 1):
        mo = doc.math_objects[j]
        start = mo.text_range.start_of_content.get()
        end = mo.text_range.end_of_content.get()
        index = int(doc.create_range(start=0, end_=start + 1).count(each=k.paragraph))
        entry = {"index": index, "nth": 1 + sum(e["index"] == index for e in out),
                 "display": "inline" not in str(mo.display_type.get())}
        inside = [(seq, number) for at, seq, number in numbers if start <= at <= end]
        if inside:
            entry["sequence"], entry["number"] = inside[0]
        out.append(entry)
    return {"success": True, "equations": out}


@mac_impl(_get_equation_impl)
def _mac_get_equation(index, nth=1):
    from . import mac as m

    doc = backend.word.document()
    mo = _mac_nth(doc, index, nth)
    inline = _mac_inline(mo)
    mo.linearize()
    linear = (_mac_range_text(doc, _mac_nth(doc, index, nth)) if inline
              else m._clean(doc.paragraphs[index].text_object.content.get()))
    _mac_nth(doc, index, nth).build_up()
    return {"success": True, "index": index, "nth": nth, "inline": inline,
            "unicodemath": linear,
            "note": "Read by turning the equation into its linear form and back."}


def _mac_maths_in(doc, start: int, end: int) -> list:
    """The math objects starting in [start, end), in order."""
    from appscript import k

    found = []
    for j in range(1, int(doc.count(each=k.math_object)) + 1):
        mo = doc.math_objects[j]
        at = mo.text_range.start_of_content.get()
        if start <= at < end:
            found.append((at, mo))
    return [mo for _, mo in sorted(found, key=lambda x: x[0])]


def _mac_nth(doc, index: int, nth: int):
    from . import mac as m

    m._para(doc, index)
    found = _mac_maths_in(doc, *m._bounds(doc, index))
    if nth > len(found):
        if not found:
            raise ValueError(f"Paragraph {index} holds no equation. word_list_equations "
                             "lists them.")
        raise ValueError(NO_SUCH_MATH.format(index=index, count=len(found), nth=nth))
    return found[nth - 1]


def _mac_inline(mo) -> bool:
    return "inline" in str(mo.display_type.get())


def _mac_range_text(doc, mo) -> str:
    from . import mac as m

    r = mo.text_range
    return m._clean(doc.create_range(start=r.start_of_content.get(),
                                     end_=r.end_of_content.get()).content.get())


@mac_impl(_insert_inline_equation_impl)
def _mac_insert_inline_equation(index, equation, fmt, position, match):
    from . import mac as m
    from ._word import locate_spans

    linear = to_unicodemath(equation, fmt)
    word = backend.word
    doc = word.document()
    m._para(doc, index)
    start, end = m._bounds(doc, index)
    if any(not _mac_inline(mo) for mo in _mac_maths_in(doc, start, end)):
        raise ValueError(DISPLAY_PARAGRAPH.format(index=index))
    if match is not None:
        content = doc.paragraphs[index].text_object.content.get()
        if content.find(match) < 0:
            raise ValueError(f"'{match}' is not in paragraph {index}. Nothing was inserted.")

        def read(a, b):
            try:
                return doc.create_range(start=a, end_=b).content.get()
            except Exception:  # noqa: BLE001 - half a character: not here
                return None

        hidden = max(0, end - start - m._u16(content, len(content)))
        s, e = locate_spans(content, [match], read, start, hidden)[0]
    elif position == "start":
        s = e = start
    else:
        s = e = end - 1
    doc.create_range(start=s, end_=e).content.set(linear)
    stop = s + m._u16(linear, len(linear))
    word.get_app().create_new_equation(from_range=doc.create_range(start=s, end_=stop),
                                       in_document=doc)
    starts = [mo.text_range.start_of_content.get()
              for mo in _mac_maths_in(doc, *m._bounds(doc, index))]
    nth = nth_of(starts, s)
    _mac_nth(doc, index, nth).build_up()
    mo = _mac_nth(doc, index, nth)
    display = not _mac_inline(mo)
    mo.linearize()
    built = _mac_range_text(doc, _mac_nth(doc, index, nth))
    _mac_nth(doc, index, nth).build_up()
    m._scroll(doc, *m._bounds(doc, index))
    return {**inline_answer(index, nth, display, linear, built),
            "note": "On macOS each step is its own undo step."}


def _mac_math_in(doc, start: int, end: int):
    from appscript import k

    for j in range(1, int(doc.count(each=k.math_object)) + 1):
        mo = doc.math_objects[j]
        if start <= mo.text_range.start_of_content.get() < end:
            return mo
    return None


@mac_impl(_set_equation_impl)
def _mac_set_equation(index, equation, fmt, nth=1):
    from appscript import k

    from . import mac as m

    linear = to_unicodemath(equation, fmt)
    doc = backend.word.document()
    mo = _mac_nth(doc, index, nth)
    if _mac_inline(mo):
        return _mac_set_inline_equation(doc, index, nth, linear)
    mo.linearize()
    start, end = m._bounds(doc, index)
    mo = _mac_math_in(doc, start, end)
    math_start, math_end = mo.text_range.start_of_content.get(), mo.text_range.end_of_content.get()
    seq = None
    for j in range(1, int(doc.count(each=k.field)) + 1):
        f = doc.fields[j]
        if (f.field_code.content.get() or "").split()[:1] == ["SEQ"] and \
                start <= f.field_code.start_of_content.get() < end:
            seq = f
            break
    stop = math_end
    number = (seq.result_range.content.get() or "").strip() if seq is not None else None
    if seq is not None:
        stop = kept_number_at(seq.field_code.start_of_content.get(),
                              lambda a, b: doc.create_range(start=a, end_=b).content.get())
        if stop is None:
            mo.build_up()
            raise ValueError(f"The number of the equation in paragraph {index} is not where "
                             "it was put. Nothing was changed.")
    read = lambda a, b: doc.create_range(start=a, end_=b).content.get()  # noqa: E731
    begin = own_text_start(math_start, read) if seq is not None else math_start
    doc.create_range(start=begin, end_=stop).content.set(linear)
    mo = _mac_math_in(doc, *m._bounds(doc, index))
    mo.build_up()
    mo = _mac_math_in(doc, *m._bounds(doc, index))
    mo.linearize()
    built = m._clean(doc.paragraphs[index].text_object.content.get())
    _mac_math_in(doc, *m._bounds(doc, index)).build_up()
    m._scroll(doc, *m._bounds(doc, index))
    return {**equation_set_answer(index, linear, built, number),
            "note": "On macOS each step is its own undo step."}


def _mac_set_inline_equation(doc, index, nth, linear):
    """_set_inline_equation, on Word for Mac."""
    from . import mac as m

    _mac_nth(doc, index, nth).linearize()
    r = _mac_nth(doc, index, nth).text_range
    start, end = r.start_of_content.get(), r.end_of_content.get()
    doc.create_range(start=start, end_=end).content.set(linear)
    found = _mac_maths_in(doc, *m._bounds(doc, index))
    if len(found) < nth or found[nth - 1].text_range.start_of_content.get() != start:
        backend.word.get_app().create_new_equation(
            from_range=doc.create_range(start=start, end_=start + m._u16(linear, len(linear))),
            in_document=doc)
    _mac_nth(doc, index, nth).build_up()
    _mac_nth(doc, index, nth).linearize()
    built = _mac_range_text(doc, _mac_nth(doc, index, nth))
    _mac_nth(doc, index, nth).build_up()
    m._scroll(doc, *m._bounds(doc, index))
    return {**equation_set_answer(index, linear, built, None), "nth": nth, "inline": True,
            "note": "On macOS each step is its own undo step."}


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
                          p.sequence, p.after, p.before,
                          p.chapter.model_dump() if p.chapter else None, undo="Insert equation")

    @mcp.tool(name="word_list_equations", annotations=annotations("List equations",
                                                                  READ_ONLY))
    async def word_list_equations() -> str:
        """The document's equations: their paragraph, which one in it each is
        (`nth`: inline equations share a paragraph), whether each is a display
        equation, and its number if it has one."""
        return await call(_list_equations_impl)

    @mcp.tool(name="word_set_equation", annotations=annotations("Change an equation", WRITES))
    async def word_set_equation(params: SetEquationInput) -> str:
        """Change an equation in place, from LaTeX or UnicodeMath: the paragraph,
        its place and a numbered equation's number stay, and so do the
        cross-references to it. Reads back what Word built."""
        p = params
        return await call(_set_equation_impl, p.index, p.equation, p.format, p.nth,
                          undo="Change equation")

    @mcp.tool(name="word_get_equation", annotations=annotations("Read an equation", WRITES))
    async def word_get_equation(params: GetEquationInput) -> str:
        """An equation as UnicodeMath, read by turning it into its linear form
        and back (two undo steps in Word)."""
        return await call(_get_equation_impl, params.index, params.nth)

    @mcp.tool(name="word_insert_inline_equation",
              annotations=annotations("Insert an inline equation", WRITES))
    async def word_insert_inline_equation(params: InsertInlineEquationInput) -> str:
        """Insert an equation among a paragraph's words ("where $E=mc^2$ is..."),
        built by Word's equation editor from LaTeX or UnicodeMath: at the start
        or end of the paragraph's text, or in place of a placeholder (`match`).
        Reads back what Word built. For an equation on a line of its own, or a
        numbered one, use word_insert_equation."""
        p = params
        return await call(_insert_inline_equation_impl, p.index, p.equation, p.format,
                          p.position, p.match, undo="Insert equation")
