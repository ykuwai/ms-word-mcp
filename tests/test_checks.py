import json
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _common, checks, mac


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def parts(*texts, **extra):
    """What the health check is given: paragraphs and the rest empty."""
    base = {"paragraphs": checks.paragraph_texts("".join(t + "\r" for t in texts)),
            "headings": [], "pictures": [], "links": [], "table_headers": [],
            "comments": 0, "revisions": 0, "tracking": False, "language": 1033,
            "unchecked": {"paragraphs": [], "language": None, "complete": True}}
    merged = {**base, **extra}
    if "language" in extra and "unchecked" not in extra:
        merged["unchecked"] = checks.unchecked_paragraphs(merged["paragraphs"], extra["language"], None)
    return merged


NONE_UNCHECKED = {"paragraphs": [], "language": None, "complete": True}


def found(*texts, **extra):
    return [f["check"] for f in checks.health(parts(*texts, **extra))]


# -- shared ----------------------------------------------------------------------------

def test_the_text_is_split_into_paragraphs_as_word_numbers_them():
    got = checks.paragraph_texts("Title\rcell one\r\x07cell two\r\x07\r\x07After\r")
    assert [(p["index"], p["text"], p["cell"]) for p in got] == [
        (1, "Title", False), (2, "cell one", True), (3, "cell two", True), (4, "", True),
        (5, "After", False)]


@pytest.mark.parametrize("lcid, name", [
    (1033, "english_us"), (1041, "japanese"), (9999999, "mixed"), (None, "mixed"),
    (0x0C00, "a custom locale (0x0c00) Word has no dictionary for"),
    (0x1000, "a custom locale (0x1000) Word has no dictionary for"), (1234, "language ID 1234"),
])
def test_languages_are_named(lcid, name):
    assert checks.language_name(lcid) == name


def test_text_in_a_custom_locale_is_said_to_be_unchecked():
    texts = checks.paragraph_texts("One.\rTwo.\r")
    found = checks.unchecked_paragraphs(texts, 0x1000, None)
    out = checks.spelling_answer(0x1000, 0, [], 0, [], 50, texts, found)
    assert out["unchecked_paragraphs"] == [1, 2]
    assert out["note"].startswith("The text is marked as a custom locale (0x1000)")
    assert "checks none of its spelling" in out["note"] and "word_set_language" in out["note"]


def test_in_a_document_of_several_languages_each_paragraph_is_asked():
    """English in a custom locale beside Japanese: the whole reads as mixed."""
    texts = checks.paragraph_texts("Thiss is it.\r\r日本語です。\rMore text.\r")
    langs = {1: 0x0C00, 3: 1041, 4: 0x0C00}
    asked = []
    found = checks.unchecked_paragraphs(texts, checks.MIXED, lambda i: asked.append(i) or langs[i])
    assert asked == [1, 3, 4]  # the empty paragraph is not asked
    assert found == {"paragraphs": [1, 4], "language": 0x0C00, "complete": True}
    note = checks.unchecked_note(found, whole=False)
    assert note.startswith("Paragraphs 1, 4 are marked as a custom locale (0x0c00)")


def test_the_paragraphs_asked_stop_at_a_limit(monkeypatch):
    monkeypatch.setattr(checks, "MAX_LANGUAGE_READS", 2)
    texts = checks.paragraph_texts("a\rb\rc\r")
    found = checks.unchecked_paragraphs(texts, checks.MIXED, lambda i: 1033)
    assert found == {"paragraphs": [], "language": None, "complete": False}
    assert "first 2 paragraphs" in checks.unchecked_note(found, whole=False)


def test_a_misspelled_word_is_counted_and_placed():
    texts = checks.paragraph_texts("Thiss is fine.\rAnd thiss too, thiss.\r")
    out = checks.spelling_answer(1033, 3, [{"word": "thiss", "index": None, "suggestions": ["this"]}],
                                 0, [], 50, texts, NONE_UNCHECKED)
    [item] = out["misspelled"]
    assert (item["word"], item["paragraphs"], item["count"]) == ("thiss", [2], 2)
    assert "note" not in out


def test_more_errors_than_listed_says_so():
    errors = [{"word": f"w{i}", "index": 1, "suggestions": []} for i in range(5)]
    assert "raise `limit`" in checks.spelling_answer(1033, 9, errors, 0, [], 5, [], NONE_UNCHECKED)["note"]


@pytest.mark.parametrize("texts, check", [
    (("Two  spaces here.",), "double_space"),
    (("A space before , a comma.",), "space_before_punctuation"),
    (("It was the the end.",), "repeated_word"),
    (("TODO: write this.",), "placeholder"),
    (("会社名は○○です。",), "placeholder"),
    (("One.", "", "", "Two."), "blank_paragraphs"),
    ((" ".join(["word"] * 210) + ".",), "long_paragraph"),
    (("あ" * 130 + "。",), "long_sentence"),
])
def test_slips_in_the_text_are_found(texts, check):
    assert check in found(*texts)


@pytest.mark.parametrize("texts", [
    ("He had had enough.",), ("全角　スペースは字下げ。",), ("One.", "", "Two."),
    ("cell\r\x07", ), ("日本語の文です。" * 5,),
])
def test_text_that_is_meant_is_not(texts):
    assert found(*texts) == []


def test_a_skipped_heading_level_is_found():
    headings = [{"index": 1, "level": 1, "text": "A"}, {"index": 2, "level": 3, "text": "B"}]
    [f] = checks.health(parts("A", "B", headings=headings))
    assert (f["check"], f["index"]) == ("heading_skip", 2)


def test_what_is_left_for_review_is_found():
    got = found("One.", pictures=[{"index": 1, "alt_text": ""}, {"index": 2, "alt_text": "A chart"}],
                links=[{"index": 1, "text": "click here"}, {"index": 1, "text": "the 2026 report"}],
                table_headers=[True, False], comments=2, revisions=1, language=0x0C00)
    assert got == ["picture_without_alt_text", "table_without_header", "vague_link_text",
                   "unchecked_language", "comments_left", "tracked_changes_left"]


def test_tracking_on_with_nothing_left_is_said():
    assert found("One.", tracking=True) == ["track_changes_on"]


def test_readability_is_given_for_english_prose_only():
    prose = "The quick brown fox jumps over the lazy dog near the river bank today. " * 4
    assert checks.check_answer(parts(prose), 100)["readability"]["reading"]
    assert "readability" not in checks.check_answer(parts("日本語の本文です。" * 20), 100)


def test_the_findings_are_cut_to_the_limit_and_counted():
    out = checks.check_answer(parts(*["TODO."] * 5), 2)
    assert len(out["findings"]) == 2 and out["counts"] == {"placeholder": 5}
    assert "3 more findings" in out["note"]


def test_end_without_start_is_refused():
    with pytest.raises(ValidationError):
        checks.SetLanguageInput(language="english_us", end=3)


# -- Windows, against the fake Word ----------------------------------------------------

def test_words_word_finds_are_listed_with_their_paragraphs(make_doc):
    doc = make_doc("Thiss is it.", "No errors here.", "Thiss again.")
    doc.spelling_spans = [(0, 5), (29, 34)]
    doc.grammar_spans = [(13, 28)]
    doc.suggestions = {"Thiss": ["This", "Thus"]}
    out = answer(checks._check_spelling_impl, 50)
    assert out["language"] == "english_us" and out["spelling_errors"] == 2
    [item] = out["misspelled"]
    assert (item["word"], item["paragraphs"], item["count"], item["suggestions"]) == (
        "Thiss", [1, 3], 2, ["This", "Thus"])
    assert out["grammar"] == [{"index": 2, "text": "No errors here."}]


def test_the_language_is_set_and_read_back(make_doc):
    doc = make_doc("One.", "Two.", "Three.")
    out = answer(checks._set_language_impl, "japanese", 2, 3)
    assert (out["language"], out["start"], out["end"]) == ("japanese", 2, 3)
    assert doc.Range(0, 5).LanguageID == 1033 and doc.Range(5, 16).LanguageID == 1041


def test_the_whole_document_by_default(make_doc):
    doc = make_doc("One.", "Two.")
    out = answer(checks._set_language_impl, "english_uk", None, None)
    assert (out["start"], out["end"]) == (1, 2) and doc.Content.LanguageID == 2057


def test_a_span_that_is_not_there_is_refused(make_doc):
    make_doc("One.")
    assert "are not all in the document" in answer(checks._set_language_impl, "japanese", 1, 4)["error"]


def test_the_health_check_reads_the_document(make_doc):
    doc = make_doc("Title", "TODO: fill in.", "", "", "Done.")
    doc.Range(0, 6).LanguageID = 0x1000
    out = answer(checks._check_document_impl, 100)
    assert out["paragraphs"] == 5 and out["language"] == "mixed"
    assert out["counts"] == {"placeholder": 1, "blank_paragraphs": 1, "unchecked_language": 1}
    [f] = [f for f in out["findings"] if f["check"] == "unchecked_language"]
    assert f["index"] == 1 and f["message"].startswith("Paragraph 1 is marked")


# -- Mac, against stand-ins for Word for Mac's scripting ------------------------------

def _keys():
    return SimpleNamespace(paragraph="paragraph", spelling_error="spelling error",
                           grammatical_error="grammatical error", table="table",
                           Word_comment="Word comment", revision="revision",
                           type_class="class", spelling_correct="correct", list="list")


class _Value:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


def _mac_word(monkeypatch, text, known=(), counts=None, language=b"\x02F\x04\x09",
              paragraph_codes=None):
    """A document whose text is `text`, and an app that knows `known` words."""
    counts = {"spelling error": 0, "grammatical error": 0, "table": 0, "Word comment": 0,
              "revision": 0, **(counts or {})}
    asked = []
    value = _Value(SimpleNamespace(code=language) if language else None)
    paragraph_codes = paragraph_codes or {}

    class Paragraphs:
        def __getitem__(self, i):
            code = paragraph_codes.get(i, language)
            return SimpleNamespace(text_object=SimpleNamespace(
                language_ID=_Value(SimpleNamespace(code=code))))

    doc = SimpleNamespace(
        text_object=SimpleNamespace(content=_Value(text), language_ID=value),
        count=lambda each: text.count("\r") if each == "paragraph" else counts[each],
        track_revisions=_Value(False), paragraphs=Paragraphs())

    def lookup(item_to_check):
        asked.append(item_to_check)
        if item_to_check in known:
            return {"class": "correct"}
        return {"class": "wrong", "list": [item_to_check.lower() + "?"]}

    app = SimpleNamespace(get_spelling_suggestions=lookup)
    monkeypatch.setattr(mac, "_k", _keys)
    monkeypatch.setattr(checks.backend, "word", SimpleNamespace(document=lambda: doc, get_app=lambda: app))
    monkeypatch.setattr(checks, "_shared_parts", lambda: {"headings": [], "pictures": [], "links": []})
    return doc, asked


def test_on_the_mac_each_distinct_latin_word_is_looked_up(monkeypatch):
    _, asked = _mac_word(monkeypatch, "Thiss is it.\rThiss NASA 日本語です。\r",
                         known={"is", "it"}, counts={"spelling error": 2})
    out = checks._mac_check_spelling(50)
    assert asked == ["Thiss", "is", "it"]  # once each; capitals and Japanese are left to Word
    [item] = out["misspelled"]
    assert (item["word"], item["paragraphs"], item["count"]) == ("Thiss", [1, 2], 2)
    assert out["language"] == "english_us"


def test_on_the_mac_nothing_is_looked_up_when_word_counts_no_errors(monkeypatch):
    _, asked = _mac_word(monkeypatch, "Thiss is it.\r")
    assert checks._mac_check_spelling(50)["misspelled"] == [] and asked == []


def test_on_the_mac_a_mixed_document_skips_the_paragraphs_word_never_checks(monkeypatch):
    _, asked = _mac_word(monkeypatch, "Thiss is it.\r日本語です。\rFine words.\r",
                         known={"Fine", "words"}, counts={"spelling error": 1},
                         language=None, paragraph_codes={1: b"\x02F\x0c\x00", 2: b"\x02F\x04\x11",
                                                         3: b"\x02F\x04\x09"})
    out = checks._mac_check_spelling(50)
    assert asked == ["Fine", "words"] and out["unchecked_paragraphs"] == [1]
    assert out["language"] == "mixed" and out["note"].startswith("Paragraph 1 is marked")


def test_on_the_mac_text_in_a_custom_locale_is_not_looked_up(monkeypatch):
    _, asked = _mac_word(monkeypatch, "Thiss is it.\r", counts={"spelling error": 1},
                         language=b"\x02F\x0c\x00")
    out = checks._mac_check_spelling(50)
    assert asked == [] and "checks none of its spelling" in out["note"]
    assert out["language"] == "a custom locale (0x0c00) Word has no dictionary for"


def test_on_the_mac_the_language_value_carries_the_windows_id():
    pytest.importorskip("aem")
    assert checks._mac_language_value(1041).code == b"\x02F\x04\x11"


def test_on_the_mac_the_health_check_reads_the_document(monkeypatch):
    _mac_word(monkeypatch, "Title\rTODO: fill in.\r", counts={"Word comment": 1},
              language=b"\x02F\x0c\x00")
    out = checks._mac_check_document(100)
    assert out["counts"] == {"placeholder": 1, "unchecked_language": 1, "comments_left": 1}
