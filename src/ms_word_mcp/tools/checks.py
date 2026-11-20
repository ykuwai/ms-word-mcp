"""Checks: spelling and grammar, the language text is checked in, and a health
check of the document before it is sent.

Word checks spelling in the language the text is marked with. Text marked with
a region Office has no dictionary for, such as English (Israel), becomes a
custom locale (0x0C00 on the Mac, 0x1000 on Windows) and is never checked:
Word found none of three misspellings marked en-IL on either platform, and all
three marked en-US or en-GB (measured 2026-11-08). word_check_spelling says so,
and word_set_language marks the text with a language Word checks.

Reading Word for Mac's list of spelling errors costs about half a second an
error (Word checks the document again for each), so on the Mac the count comes
from Word and the words are found by looking each distinct Latin-script word up
(about 10 ms a word), which found the same words as Word's list (measured
2026-11-06). Other scripts are left to Word's count.

The health check reads the document's text in one call, the outline from
Word's heading list, and a few counts, so it is quick on a long document.
https://learn.microsoft.com/en-us/office/vba/api/word.document.spellingerrors
https://learn.microsoft.com/en-us/office/vba/api/word.range.languageid
https://learn.microsoft.com/en-us/office/vba/api/word.wdlanguageid
"""

from __future__ import annotations

import re
from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, MAC_IMPLS, Input, annotations, call, mac_impl
from ._word import check_span, clean_text, index_of, span_range

#: Windows language IDs (WdLanguageID) by name. Word for Mac's values are
#: 0x0246 followed by the same ID.
LANGUAGES = {
    "english_us": 1033, "english_uk": 2057, "english_au": 3081, "english_ca": 4105,
    "english_nz": 5129, "english_ie": 6153, "english_za": 7177, "english_in": 16393,
    "french": 1036, "french_ca": 3084, "german": 1031, "spanish": 3082, "spanish_mexico": 2058,
    "italian": 1040, "portuguese": 2070, "portuguese_brazil": 1046, "dutch": 1043, "danish": 1030,
    "swedish": 1053, "norwegian": 1044, "finnish": 1035, "polish": 1045, "czech": 1029,
    "greek": 1032, "turkish": 1055, "russian": 1049, "ukrainian": 1058, "hebrew": 1037,
    "arabic": 1025, "japanese": 1041, "korean": 1042, "chinese_simplified": 2052,
    "chinese_traditional": 1028, "no_proofing": 1024,
}
LANGUAGE_NAMES = {v: k for k, v in LANGUAGES.items()}
Language = Literal[tuple(LANGUAGES)]  # type: ignore[valid-type]

#: A range holding more than one language answers wdUndefined.
MIXED = 9999999
#: Locale IDs Windows gives a region it has no ID for; Word has no dictionary
#: for any of them.
CUSTOM_LOCALES = (0x0400, 0x0800, 0x0C00, 0x1000, 0x1400)
#: The most distinct words the Mac looks up in one call (about 10 ms each).
MAX_WORDS_LOOKED_UP = 3000
#: The most paragraphs whose language is asked when a document holds several.
MAX_LANGUAGE_READS = 300

LATIN_WORD = re.compile(r"[A-Za-zÀ-ÖØ-öø-ɏ]+(?:['’][A-Za-zÀ-ÖØ-öø-ɏ]+)*")
SENTENCE = re.compile(r"[^.!?。！？]+[.!?。！？]*")
PLACEHOLDER = re.compile(
    r"\b(TODO|TBD|TBC|FIXME|XXX+)\b|lorem ipsum|\[(insert|add|your)[^\]]*\]|○○|〇〇|××", re.I)
VAGUE_LINKS = {"here", "click here", "this", "link", "read more", "more", "this link", "learn more",
               "ここ", "こちら", "ここをクリック", "詳細"}
LONG_SENTENCE_WORDS, LONG_PARAGRAPH_WORDS = 35, 200
#: A sentence or paragraph of Japanese or Chinese is measured in characters.
LONG_SENTENCE_CJK, LONG_PARAGRAPH_CJK = 120, 600
CJK = re.compile(r"[぀-ヿ㐀-鿿豈-﫿]")


class CheckSpellingInput(Input):
    limit: int = Field(50, ge=1, le=500, description="At most this many errors of each kind.")


class SetLanguageInput(Input):
    language: Language = Field(..., description=(
        "The language the text is written in. no_proofing turns checking off."))
    start: int | None = Field(None, ge=1, description=(
        "First paragraph. Leave out, with `end`, for the whole document."))
    end: int | None = Field(None, ge=1, description="Last paragraph; leave out for `start` alone.")

    @model_validator(mode="after")
    def _start_with_end(self):
        if self.end is not None and self.start is None:
            raise ValueError("Give `start` with `end`, or neither for the whole document.")
        return self


class CheckDocumentInput(Input):
    limit: int = Field(100, ge=1, le=1000, description="At most this many findings.")


# -- shared ----------------------------------------------------------------------------

def language_name(lcid: int | None) -> str:
    if lcid is None or lcid == MIXED:
        return "mixed"
    if lcid in LANGUAGE_NAMES:
        return LANGUAGE_NAMES[lcid]
    if lcid in CUSTOM_LOCALES:
        return f"a custom locale (0x{lcid:04x}) Word has no dictionary for"
    return f"language ID {lcid}"


def unchecked(lcid: int | None) -> bool:
    return lcid in CUSTOM_LOCALES


def language_note(lcid: int | None) -> str | None:
    if unchecked(lcid):
        return (f"The text is marked as {language_name(lcid)}, so Word checks none of its "
                "spelling. Text typed on a Mac set to a region Office has no dictionary "
                "for, such as English (Israel), is marked so. word_set_language marks it "
                "with a language Word checks.")
    if lcid == LANGUAGES["no_proofing"]:
        return ("The text is marked 'do not check spelling or grammar'. word_set_language "
                "sets the language it is written in.")
    return None


def unchecked_paragraphs(texts: list[dict], lcid, language_of) -> dict:
    """Which paragraphs Word never checks: {"paragraphs", "language", "complete"}.

    When the whole text is one language, that settles it. A document holding
    several (English and Japanese, say) answers wdUndefined as a whole, and text
    in a custom locale hides behind it, so each paragraph with text is asked,
    `language_of(index)`, up to MAX_LANGUAGE_READS of them."""
    written = [p for p in texts if _printable(p["text"]).strip()]
    if lcid is not None and lcid != MIXED:
        return {"paragraphs": [p["index"] for p in written] if unchecked(lcid) else [],
                "language": lcid, "complete": True}
    found, code = [], None
    for read, p in enumerate(written):
        if read >= MAX_LANGUAGE_READS:
            return {"paragraphs": found, "language": code, "complete": False}
        got = language_of(p["index"])
        if unchecked(got):
            found.append(p["index"])
            code = code or got
    return {"paragraphs": found, "language": code, "complete": True}


def unchecked_note(found: dict, whole: bool) -> str | None:
    """What to say about text Word never checks, if there is any."""
    if not found["paragraphs"]:
        if not found["complete"]:
            return (f"The document holds several languages; the first {MAX_LANGUAGE_READS} "
                    "paragraphs with text are all in languages Word checks.")
        return None
    numbers = found["paragraphs"]
    where = "The text" if whole else (
        f"Paragraph {numbers[0]}" if len(numbers) == 1 else
        f"Paragraphs {', '.join(map(str, numbers[:10]))}{' and more' if len(numbers) > 10 else ''}")
    note = (f"{where} {'is' if whole or len(numbers) == 1 else 'are'} marked as "
            f"{language_name(found['language'])}, so Word checks none of {'its' if whole or len(numbers) == 1 else 'their'} "
            "spelling. Text typed on a Mac set to a region Office has no dictionary for, such as "
            "English (Israel), is marked so. word_set_language marks it with a language Word checks.")
    if not found["complete"]:
        note += f" Only the first {MAX_LANGUAGE_READS} paragraphs with text were asked."
    return note


def paragraph_texts(content: str) -> list[dict]:
    """The document's text in one string, as paragraphs: their text and whether
    each is a table cell (Word ends a cell and a row with \\r\\x07)."""
    raw = (content or "").split("\r")
    if raw and raw[-1] == "":
        raw.pop()
    out = []
    for i, text in enumerate(raw):
        # A cell's end is \r\x07, so the \x07 opens the next piece of text.
        cell = i + 1 < len(raw) and raw[i + 1].startswith("\x07")
        out.append({"index": i + 1, "text": clean_text(text.lstrip("\x07")), "cell": cell})
    return out


def _printable(text: str) -> str:
    """Text without Word's marks for fields, notes, ruby and pictures."""
    return re.sub(r"[\x00-\x08\x0b\x0e-\x1f]", "", text)


def spelling_answer(lcid, count: int, errors: list[dict], grammar_count: int,
                    grammar: list[dict], limit: int, texts: list[dict], found: dict) -> dict:
    """What word_check_spelling answers, from either platform's reads.

    `errors` are {"word", "index" (or None), "suggestions"}; a word with no
    place is found in `texts`."""
    words: dict[str, dict] = {}
    for e in errors:
        item = words.setdefault(e["word"], {"word": e["word"], "count": 0, "paragraphs": [],
                                            "suggestions": e["suggestions"]})
        item["count"] += 1
        if e.get("index") is not None and e["index"] not in item["paragraphs"] \
                and len(item["paragraphs"]) < 10:
            item["paragraphs"].append(e["index"])
    for item in words.values():
        if not item["paragraphs"]:
            pattern = re.compile(r"(?<!\w)" + re.escape(item["word"]) + r"(?!\w)")
            hits = [p["index"] for p in texts if pattern.search(p["text"])]
            item["paragraphs"] = hits[:10]
            item["count"] = max(item["count"], sum(len(pattern.findall(p["text"])) for p in texts))
    result = {"success": True, "language": language_name(lcid), "spelling_errors": count,
              "misspelled": sorted(words.values(), key=lambda w: (w["paragraphs"][:1] or [0])[0]),
              "grammar_errors": grammar_count, "grammar": grammar}
    if found["paragraphs"]:
        result["unchecked_paragraphs"] = found["paragraphs"][:50]
    whole = lcid is not None and lcid != MIXED
    notes = [n for n in (unchecked_note(found, whole) if not lcid == LANGUAGES["no_proofing"]
                         else language_note(lcid),) if n]
    if len(errors) >= limit or len(grammar) >= limit:
        notes.append("There are more than are listed; raise `limit` to see them.")
    if notes:
        result["note"] = " ".join(notes)
    return result


def _words_of(text: str) -> int:
    """Words in Latin text, and a CJK character as one."""
    return len(LATIN_WORD.findall(text)) + len(CJK.findall(text))


def _syllables(word: str) -> int:
    word = word.lower()
    groups = re.findall(r"[aeiouy]+", word)
    n = len(groups) - (1 if word.endswith("e") and len(groups) > 1 and not word.endswith("le") else 0)
    return max(n, 1)


def readability(texts: list[str]) -> dict | None:
    """Flesch reading ease and Flesch-Kincaid grade, for English prose only."""
    sentences = [s for t in texts for s in SENTENCE.findall(t) if re.search(r"[A-Za-z]", s)]
    words = [w for t in texts for w in re.findall(r"[A-Za-z]+(?:'[a-z]+)?", t)]
    if len(words) < 30 or not sentences:
        return None
    syllables = sum(_syllables(w) for w in words)
    wps, spw = len(words) / len(sentences), syllables / len(words)
    ease = 206.835 - 1.015 * wps - 84.6 * spw
    return {"reading_ease": round(ease, 1), "grade_level": round(0.39 * wps + 11.8 * spw - 15.59, 1),
            "words_per_sentence": round(wps, 1),
            "reading": ("very easy" if ease >= 80 else "easy" if ease >= 70 else "plain" if ease >= 60
                        else "fairly hard" if ease >= 50 else "hard" if ease >= 30 else "very hard")}


def health(parts: dict) -> list[dict]:
    """The findings of the health check, from what each platform read:
    `paragraphs` (paragraph_texts), `headings` (index, level, text), `pictures`
    (index, alt_text), `links` (index, text), `table_headers` (one bool a table),
    `comments`, `revisions`, `tracking` and `language`."""
    found = []

    def add(check, severity, message, index=None, text=None):
        item = {"check": check, "severity": severity, "message": message}
        if index:
            item["index"] = index
        if text:
            item["text"] = text[:120]
        found.append(item)

    paragraphs, headings = parts["paragraphs"], parts["headings"]
    previous = 0
    for h in headings:
        if previous and h["level"] > previous + 1:
            add("heading_skip", "warning", f"Heading level {h['level']} follows level {previous}; "
                "a level is skipped, which breaks the outline and screen-reader navigation.",
                h["index"], h["text"])
        previous = h["level"]
    heading_at = {h["index"] for h in headings}
    total = sum(_words_of(p["text"]) for p in paragraphs)
    if not headings and total > 500:
        add("no_headings", "info", f"About {total} words and no headings: headings give the "
            "document an outline, a table of contents and navigation.")
    run: list[int] = []
    for p in paragraphs + [None]:
        if p is not None and not p["cell"] and not _printable(p["text"]).strip() \
                and "\x0c" not in p["text"] and p["index"] < len(paragraphs):
            run.append(p["index"])
            continue
        if len(run) >= 2:
            add("blank_paragraphs", "info", f"{len(run)} empty paragraphs in a row, used as "
                "spacing; space before or after, or a page break, holds up better.", run[0])
        run = []
    for p in paragraphs:
        t = _printable(p["text"])
        if not t.strip():
            continue
        if "  " in t.strip():
            add("double_space", "info", "Two spaces in a row.", p["index"], t)
        if re.search(r"[A-Za-z0-9] ([,.;:!?])(\s|$)", t):
            add("space_before_punctuation", "info", "A space before punctuation.", p["index"], t)
        m = re.search(r"\b([A-Za-z]+) \1\b", t, re.I)
        if m and m.group(1).lower() not in ("had", "that", "bye", "no", "so"):
            add("repeated_word", "warning", f"'{m.group(0)}': a word twice in a row.", p["index"], t)
        m = PLACEHOLDER.search(t)
        if m:
            add("placeholder", "warning", f"'{m.group(0)}' looks like a placeholder left in.",
                p["index"], t)
        if p["index"] in heading_at or p["cell"]:
            continue
        cjk = len(CJK.findall(t)) > len(t) / 3
        long_paragraph = len(t) > LONG_PARAGRAPH_CJK if cjk else _words_of(t) > LONG_PARAGRAPH_WORDS
        if long_paragraph:
            add("long_paragraph", "info", f"A paragraph of {len(t) if cjk else _words_of(t)} "
                f"{'characters' if cjk else 'words'}.", p["index"], t)
        for s in SENTENCE.findall(t):
            size = len(s.strip()) if cjk else len(s.split())
            if size > (LONG_SENTENCE_CJK if cjk else LONG_SENTENCE_WORDS):
                add("long_sentence", "info", f"A sentence of {size} "
                    f"{'characters' if cjk else 'words'}.", p["index"], s.strip())
                break
    for picture in parts["pictures"]:
        if not (picture.get("alt_text") or "").strip():
            add("picture_without_alt_text", "warning", "A picture without alternative text: "
                "screen readers cannot describe it. word_set_picture sets it.", picture["index"])
    for number, has_header in enumerate(parts["table_headers"], start=1):
        if not has_header:
            add("table_without_header", "info", f"Table {number} has no header row, repeated on "
                "each page and read out by screen readers.")
    for link in parts["links"]:
        words = re.sub(r"[\s.,:;!?…»›>。、]+", " ", link.get("text") or "").strip().lower()
        if words in VAGUE_LINKS:
            add("vague_link_text", "info", f"A link reads '{link['text']}', which says nothing "
                "out of context: a screen reader can list a page's links by their words alone.",
                link["index"])
    found_unchecked = parts["unchecked"]
    if found_unchecked["paragraphs"]:
        whole = parts["language"] is not None and parts["language"] != MIXED
        add("unchecked_language", "warning", unchecked_note(found_unchecked, whole),
            found_unchecked["paragraphs"][0])
    if parts["comments"]:
        n = parts["comments"]
        add("comments_left", "warning", f"{n} comment{'s are' if n > 1 else ' is'} still in the "
            "document (word_list_comments).")
    if parts["revisions"]:
        n = parts["revisions"]
        add("tracked_changes_left", "warning", f"{n} tracked change{'s are' if n > 1 else ' is'} "
            "not accepted or rejected yet (word_list_revisions).")
    elif parts["tracking"]:
        add("track_changes_on", "info", "Track changes is on.")
    return found


def check_answer(parts: dict, limit: int) -> dict:
    found = health(parts)
    counts: dict[str, int] = {}
    for f in found:
        counts[f["check"]] = counts.get(f["check"], 0) + 1
    heading_at = {h["index"] for h in parts["headings"]}
    prose = [_printable(p["text"]) for p in parts["paragraphs"]
             if not p["cell"] and p["index"] not in heading_at and len(p["text"].split()) >= 4]
    result = {"success": True, "findings": found[:limit], "counts": counts,
              "paragraphs": len(parts["paragraphs"]), "headings": len(parts["headings"]),
              "language": language_name(parts["language"])}
    score = readability(prose)
    if score:
        result["readability"] = score
    notes = ["Spelling and grammar are word_check_spelling's."]
    if len(found) > limit:
        notes.append(f"{len(found) - limit} more findings; raise `limit` to see them.")
    if score:
        notes.append("Readability uses the English formulas (Flesch): 60 to 70 reads plainly; "
                     "a higher grade is harder.")
    result["note"] = " ".join(notes)
    return result


def _same(impl):
    """The implementation of `impl` for this platform, to read with it here."""
    return MAC_IMPLS[impl.__name__] if backend.IS_MACOS else impl


def _shared_parts() -> dict:
    """What the health check reads through the other tools' own reads."""
    from . import links, references
    from . import text as text_tools

    return {"headings": _same(text_tools._get_outline_impl)(False)["headings"],
            "pictures": _same(references._list_pictures_impl)()["pictures"],
            "links": _same(links._list_hyperlinks_impl)(1, 200)["hyperlinks"]}


# -- Windows ---------------------------------------------------------------------------

def _language(rng) -> int:
    return int(rng.LanguageID)


def _check_spelling_impl(limit):
    doc = backend.word.document()
    texts = paragraph_texts(doc.Content.Text)
    spelling = doc.SpellingErrors
    errors = []
    for i in range(1, min(spelling.Count, limit) + 1):
        rng = spelling.Item(i)
        try:
            suggestions = [s.Name for s in rng.GetSpellingSuggestions()][:3]
        except Exception:  # noqa: BLE001 - the suggestions are a nicety
            suggestions = []
        errors.append({"word": rng.Text, "index": index_of(doc, rng), "suggestions": suggestions})
    grammar_errors = doc.GrammaticalErrors
    grammar = [{"index": index_of(doc, g), "text": clean_text(g.Text)[:200]}
               for g in (grammar_errors.Item(i)
                         for i in range(1, min(grammar_errors.Count, limit) + 1))]
    lcid = _language(doc.Content)
    found = unchecked_paragraphs(texts, lcid, lambda i: _language(doc.Paragraphs.Item(i).Range))
    return spelling_answer(lcid, spelling.Count, errors, grammar_errors.Count, grammar, limit,
                           texts, found)


def language_set(language, start, end, got) -> dict:
    lcid = LANGUAGES[language]
    if got != lcid:
        raise RuntimeError(f"Paragraphs {start} to {end} read back as {language_name(got)}, not "
                           f"{language}.")
    return {"success": True, "language": language, "start": start, "end": end,
            "note": "Word checks spelling and grammar in this language now; "
                    "word_check_spelling lists what it finds."}


def _set_language_impl(language, start, end):
    doc = backend.word.document()
    if start is None:
        start, end = 1, doc.Paragraphs.Count
    start, end = check_span(doc, start, end)
    rng = span_range(doc, start, end)
    rng.LanguageID = LANGUAGES[language]
    return language_set(language, start, end, _language(span_range(doc, start, end)))


def _check_document_impl(limit):
    doc = backend.word.document()
    tables = doc.Tables
    texts = paragraph_texts(doc.Content.Text)
    lcid = _language(doc.Content)
    parts = {"paragraphs": texts, **_shared_parts(),
             "table_headers": [bool(tables.Item(i).Rows.Item(1).HeadingFormat)
                               for i in range(1, tables.Count + 1)],
             "comments": doc.Comments.Count, "revisions": doc.Revisions.Count,
             "tracking": bool(doc.TrackRevisions), "language": lcid,
             "unchecked": unchecked_paragraphs(
                 texts, lcid, lambda i: _language(doc.Paragraphs.Item(i).Range))}
    return check_answer(parts, limit)


# -- Mac -------------------------------------------------------------------------------

def _mac_code(value) -> int | None:
    """The Windows language ID in a Word for Mac language value: an enumerator
    it names (english_us), or a code it does not (a custom locale)."""
    from . import mac as m

    if value is None or m._name(value) == "missing_value":
        return None
    raw = getattr(value, "code", None)
    if raw is None:
        enum = backend.word.get_app().AS_appdata.typebyname().get(getattr(value, "AS_name", ""))
        if enum is None:
            return None
        raw = enum.code
    return int.from_bytes(raw, "big") & 0xFFFF


def _mac_language_value(lcid: int):
    import aem

    return aem.AEEnum((0x02460000 | lcid).to_bytes(4, "big"))


def _mac_paragraph_language(doc, index: int) -> int | None:
    return _mac_code(doc.paragraphs[index].text_object.language_ID.get())


def _mac_texts(doc) -> list[dict]:
    return paragraph_texts(doc.text_object.content.get())


def _mac_misspelled(texts: list[dict], limit: int) -> list[dict]:
    """Each distinct Latin-script word Word does not know, looked up one by one.
    A word in capitals is left alone, as Word leaves it by default."""
    from . import mac as m

    k = m._k()
    app = backend.word.get_app()
    seen, out = set(), []
    for word in LATIN_WORD.findall("\n".join(_printable(p["text"]) for p in texts)):
        if word in seen or (word.isupper() and len(word) > 1):
            continue
        seen.add(word)
        if len(seen) > MAX_WORDS_LOOKED_UP or len(out) >= limit:
            break
        answer = app.get_spelling_suggestions(item_to_check=word)
        if answer.get(k.type_class) != k.spelling_correct:
            out.append({"word": word, "index": None, "suggestions": list(answer.get(k.list) or [])[:3]})
    return out


@mac_impl(_check_spelling_impl)
def _mac_check_spelling(limit):
    from . import mac as m

    k = m._k()
    doc = backend.word.document()
    texts = _mac_texts(doc)
    lcid = _mac_code(doc.text_object.language_ID.get())
    count = int(doc.count(each=k.spelling_error))
    found = unchecked_paragraphs(texts, lcid, lambda i: _mac_paragraph_language(doc, i))
    skip = set(found["paragraphs"])
    checked = [p for p in texts if p["index"] not in skip]
    errors = _mac_misspelled(checked, limit) if count and checked else []
    grammar_count = int(doc.count(each=k.grammatical_error))
    grammar = []
    for i in range(1, min(grammar_count, limit) + 1):
        g = doc.grammatical_errors[i]
        at = g.start_of_content.get()
        grammar.append({"index": int(doc.create_range(start=0, end_=at + 1).count(each=k.paragraph)),
                        "text": m._clean(g.content.get())[:200]})
    return spelling_answer(lcid, count, errors, grammar_count, grammar, limit, texts, found)


@mac_impl(_set_language_impl)
def _mac_set_language(language, start, end):
    from . import mac as m

    doc = backend.word.document()
    if start is None:
        start, end = 1, m._count(doc)
    start, end = m._span(doc, start, end)
    first, _ = m._bounds(doc, start)
    _, last = m._bounds(doc, end)
    doc.create_range(start=first, end_=last).language_ID.set(_mac_language_value(LANGUAGES[language]))
    got = _mac_code(doc.create_range(start=first, end_=last).language_ID.get())
    return language_set(language, start, end, MIXED if got is None else got)


@mac_impl(_check_document_impl)
def _mac_check_document(limit):
    from . import mac as m

    k = m._k()
    doc = backend.word.document()
    tables = int(doc.count(each=k.table))
    texts = _mac_texts(doc)
    lcid = _mac_code(doc.text_object.language_ID.get())
    parts = {"paragraphs": texts, **_shared_parts(),
             "table_headers": [bool(doc.tables[i].rows[1].heading_format.get())
                               for i in range(1, tables + 1)],
             "comments": int(doc.count(each=k.Word_comment)),
             "revisions": int(doc.count(each=k.revision)),
             "tracking": bool(doc.track_revisions.get()), "language": lcid,
             "unchecked": unchecked_paragraphs(texts, lcid, lambda i: _mac_paragraph_language(doc, i))}
    return check_answer(parts, limit)


def register(mcp) -> None:
    @mcp.tool(name="word_check_spelling", annotations=annotations("Check spelling", READ_ONLY))
    async def word_check_spelling(params: CheckSpellingInput) -> str:
        """Word's spelling and grammar findings: each misspelled word with the
        paragraphs it is in and Word's suggestions, each grammar issue with its
        paragraph, and the language the text is checked in. Text marked with a
        language Word has no dictionary for is never checked; the answer says so."""
        return await call(_check_spelling_impl, params.limit)

    @mcp.tool(name="word_set_language", annotations=annotations("Set the language", WRITES))
    async def word_set_language(params: SetLanguageInput) -> str:
        """Mark text with the language it is written in, which Word checks its
        spelling and grammar in: some paragraphs, or the whole document."""
        p = params
        return await call(_set_language_impl, p.language, p.start, p.end, undo="Set language")

    @mcp.tool(name="word_check_document", annotations=annotations("Check the document", READ_ONLY))
    async def word_check_document(params: CheckDocumentInput) -> str:
        """A check before the document is sent: skipped heading levels, empty
        paragraphs used as spacing, double spaces, a word typed twice,
        placeholders left in (TODO, TBD, lorem ipsum, ○○), very long sentences
        and paragraphs, pictures without alt text, tables without a header row,
        links whose words say nothing ("click here"), text Word never
        spell-checks, comments and tracked changes left; and, for English,
        readability."""
        return await call(_check_document_impl, params.limit)
