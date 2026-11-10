"""Tracked changes: the tools for them, and the other tools while tracking is on."""

import json
import time
from types import SimpleNamespace

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _common, _word, mac, revisions, tables, text


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def tracking(word, make_doc, *paragraphs):
    make_doc(*paragraphs)
    word.doc.TrackRevisions = True
    return word.doc


def read(start=1, end=None):
    return answer(text._get_paragraphs_impl, start, end, False)["paragraphs"]


def test_tracking_turns_on_and_off(word, make_doc):
    make_doc("one")
    assert answer(revisions._set_track_changes_impl, True)["tracking"] is True
    assert word.doc.TrackRevisions is True
    assert answer(revisions._set_track_changes_impl, False)["tracking"] is False


def test_a_paragraph_reads_without_its_tracked_deletion(word, make_doc):
    doc = tracking(word, make_doc, "Alpha beta gamma")
    doc.Range(6, 11).Delete()
    info = read()[0]
    assert info["text"] == "Alpha gamma"
    assert info["deleted_text"] == "beta "
    assert "deleted" not in info


def test_a_paragraph_without_changes_says_nothing_of_them(word, make_doc):
    make_doc("Alpha")
    assert set(read()[0]) == {"index", "text", "style"}


def test_rewriting_while_tracking_reads_back_clean(word, make_doc):
    tracking(word, make_doc, "Alpha beta")
    out = answer(text._set_paragraph_text_impl, 1, "Rewritten")
    assert out["paragraph"]["text"] == "Rewritten"
    assert out["paragraph"]["deleted_text"] == "Alpha beta"
    assert out["was"] == "Alpha beta"
    assert "warnings" not in out


def test_deleting_while_tracking_is_no_warning_and_keeps_the_numbers(word, make_doc):
    tracking(word, make_doc, "one", "two", "three")
    out = answer(text._delete_paragraphs_impl, 2, None)
    assert "warnings" not in out
    assert "marked as deleted" in out["note"]
    assert out["paragraphs"] == 3
    infos = read()
    assert [i["index"] for i in infos] == [1, 2, 3]
    assert infos[1]["text"] == "" and infos[1]["deleted"] is True
    assert infos[1]["deleted_text"] == "two"


def test_find_leaves_out_text_marked_deleted(word, make_doc):
    doc = tracking(word, make_doc, "cat and cat")
    doc.Range(0, 3).Delete()
    assert answer(text._find_replace_impl, "cat", None, False, False, False)["matches"] == 1


def test_replacing_twice_while_tracking_does_not_replace_the_deleted_text(word, make_doc):
    tracking(word, make_doc, "cat")
    assert answer(text._find_replace_impl, "cat", "dog", False, False, False)["replaced"] == 1
    assert answer(text._find_replace_impl, "cat", "dog", False, False, False)["matches"] == 0
    assert read()[0]["text"] == "dog"


def listed(start=1, limit=100):
    return answer(revisions._list_revisions_impl, start, limit)


def test_changes_are_listed_in_order(word, make_doc):
    doc = tracking(word, make_doc, "Alpha beta", "Second")
    doc.Range(11, 17).Delete()
    doc.Range(0, 0).Text = "New "
    out = listed()
    assert (out["tracking"], out["count"]) == (True, 2)
    assert [(r["number"], r["kind"], r["index"], r["text"]) for r in out["revisions"]] == [
        (1, "insert", 1, "New "), (2, "delete", 2, "Second")]
    assert out["revisions"][0]["date"] == "2026-11-09T10:30:00"
    assert out["revisions"][0]["author"] == "Reviewer"


def test_a_long_list_says_where_to_go_on(word, make_doc):
    doc = tracking(word, make_doc, "abcdefgh")
    doc.Range(0, 1).Delete()
    doc.Range(2, 3).Delete()
    assert listed(1, 1)["more"] == "Changes 2 to 2 follow; pass start=2."


def test_long_text_is_cut_and_says_so(word, make_doc):
    doc = tracking(word, make_doc, "x" * 900)
    doc.Range(0, 900).Delete()
    change = listed()["revisions"][0]
    assert len(change["text"]) == 500 and change["text_cut"] is True


def test_accepting_a_deletion_removes_the_text(word, make_doc):
    doc = tracking(word, make_doc, "Alpha beta gamma")
    doc.Range(6, 11).Delete()
    out = answer(revisions._accept_revisions_impl, [1], False)
    assert (out["accepted"], out["left"]) == (1, 0)
    assert doc.texts() == ["Alpha gamma"]


def test_rejecting_a_deletion_brings_it_back(word, make_doc):
    doc = tracking(word, make_doc, "Alpha beta gamma")
    doc.Range(6, 11).Delete()
    assert answer(revisions._reject_revisions_impl, [1], False)["rejected"] == 1
    assert doc.texts() == ["Alpha beta gamma"]
    assert doc.Revisions.Count == 0


def test_rejecting_an_insertion_takes_it_out(word, make_doc):
    doc = tracking(word, make_doc, "Alpha")
    doc.Range(0, 0).Text = "New "
    answer(revisions._reject_revisions_impl, None, True)
    assert doc.texts() == ["Alpha"]


def test_several_numbers_are_the_ones_listed_before_the_call(word, make_doc):
    doc = tracking(word, make_doc, "abcdefgh")
    doc.Range(0, 1).Delete()  # a
    doc.Range(2, 3).Delete()  # c
    doc.Range(4, 5).Delete()  # e
    out = answer(revisions._accept_revisions_impl, [1, 3], False)
    assert (out["accepted"], out["left"]) == (2, 1)
    assert doc.texts() == ["bcdfgh"]
    assert [r["text"] for r in listed()["revisions"]] == ["c"]


def test_every_change_at_once(word, make_doc):
    doc = tracking(word, make_doc, "Alpha beta", "two")
    doc.Range(0, 6).Delete()
    doc.Range(11, 14).Delete()
    out = answer(revisions._accept_revisions_impl, None, True)
    assert (out["accepted"], out["left"]) == (2, 0)
    assert doc.texts() == ["beta", ""]


def test_a_number_that_is_not_there_changes_nothing(word, make_doc):
    doc = tracking(word, make_doc, "Alpha beta")
    doc.Range(0, 6).Delete()
    out = answer(revisions._accept_revisions_impl, [1, 9], False)
    assert "There is no change 9" in out["error"]
    assert doc.Revisions.Count == 1 and doc.texts() == ["Alpha beta"]


@pytest.mark.parametrize("arguments", [{}, {"numbers": [1], "every": True}, {"numbers": [0]}])
def test_numbers_or_every_but_one_of_them(arguments):
    with pytest.raises(ValidationError):
        revisions.ResolveRevisionsInput(**arguments)


# -- ruby marked deleted ------------------------------------------------------------

def _ruby_code(word):
    return f'EQ \\* jc0 \\* hps10 \\o(\\s\\up 9(r),{word})'


FIELDS = [(_ruby_code("OLD"), 1, 52), (_ruby_code("NEW"), 59, 110)]


def test_a_ruby_word_marked_deleted_gives_way_to_the_next():
    """Measured: the deleted field stays in Word's list, its code inside the deletion."""
    def piece(a, b):
        return SimpleNamespace(Text={(0, 53): "\x15", (53, 112): " and \x15\r"}[(a, b)])

    code = [SimpleNamespace(Text=t, Start=a, End=b) for t, a, b in FIELDS]
    rng = SimpleNamespace(
        Text="\x15 and \x15\r", Start=0, End=112,
        Document=SimpleNamespace(Range=piece),
        Revisions=SimpleNamespace(Count=1, Item=lambda i: SimpleNamespace(
            Type=2, Range=SimpleNamespace(Start=0, End=53))),
        Fields=SimpleNamespace(Count=2, Item=lambda i: SimpleNamespace(Code=code[i - 1])))
    p = SimpleNamespace(Range=rng, Style=SimpleNamespace(NameLocal="Normal"), OutlineLevel=10)
    info = _word.paragraph_info(1, p)
    assert info["text"] == " and NEW"
    assert info["deleted_text"] == "OLD"
    assert [r["text"] for r in info["ruby"]] == ["NEW"]


def test_on_the_mac_a_ruby_word_marked_deleted_gives_way_to_the_next(monkeypatch):
    value = lambda v: SimpleNamespace(get=lambda: v)  # noqa: E731
    fields = {i: SimpleNamespace(field_code=SimpleNamespace(
        content=value(t), start_of_content=value(a), end_of_content=value(b)))
        for i, (t, a, b) in enumerate(FIELDS, 1)}
    rng = SimpleNamespace(content=value("\x15 and \x15\r"), fields=fields,
                          count=lambda each: 2)
    doc = SimpleNamespace(
        paragraphs={1: SimpleNamespace(
            text_object=rng, style=SimpleNamespace(name_local=value("Normal")),
            outline_level=value(SimpleNamespace(name="outline_level_body_text")))},
        create_range=lambda start, end_: SimpleNamespace(
            content=value({(0, 53): "\x15", (53, 112): " and \x15\r"}[(start, end_)])))
    monkeypatch.setattr(mac, "_k", lambda: SimpleNamespace(field="field"))
    monkeypatch.setattr(mac, "_bounds", lambda d, i: (0, 112))
    info = mac._info(doc, 1, spans=[(0, 53)])
    assert info["text"] == " and NEW"
    assert info["deleted_text"] == "OLD"
    assert [r["text"] for r in info["ruby"]] == ["NEW"]


def test_ruby_with_nothing_deleted_reads_as_before():
    info = {"text": "\x15と\x15"}
    _word.expand_ruby(info, lambda: [(_ruby_code("漢字"), False), (_ruby_code("文章"), False)])
    assert info["text"] == "漢字と文章"
    assert [r["text"] for r in info["ruby"]] == ["漢字", "文章"]


def test_fields_are_not_read_without_a_ruby_mark():
    def boom():
        raise AssertionError("no field should be read")

    info = {"text": "plain", "deleted_text": "gone"}
    _word.expand_ruby(info, boom)
    assert info == {"text": "plain", "deleted_text": "gone"}


# -- on the Mac, only the changes near what is read ---------------------------------

def _mac_kinds():
    return SimpleNamespace(revision="revision", revision_delete="delete",
                           revision_move_from="move", revision_insert="insert", field="field")


class _NoSuch(Exception):
    errornumber = -1728


def _mac_range(changes):
    """A text range whose changes are (kind, start, end), asked for by number."""
    value = lambda v: SimpleNamespace(get=lambda: v)  # noqa: E731

    def missing():
        raise _NoSuch("The object you are trying to access does not exist")

    class Changes:
        def __getitem__(self, i):
            if i > len(changes):
                return SimpleNamespace(revision_type=SimpleNamespace(get=missing))
            kind, a, b = changes[i - 1]
            return SimpleNamespace(revision_type=value(kind), text_object=SimpleNamespace(
                start_of_content=value(a), end_of_content=value(b)))

        def get(self):
            raise AssertionError("the whole list at once left Word busy (measured)")

    return SimpleNamespace(revisions=Changes())


def test_a_range_with_no_changes_answers_none(monkeypatch):
    monkeypatch.setattr(mac, "_k", _mac_kinds)
    assert mac._deletions(_mac_range([])) == []


def test_only_deletions_of_the_range_are_read(monkeypatch):
    monkeypatch.setattr(mac, "_k", _mac_kinds)
    rng = _mac_range([("insert", 1, 3), ("delete", 5, 9), ("move", 12, 15)])
    assert mac._deletions(rng) == [(5, 9), (12, 15)]


def test_a_mac_read_asks_about_changes_only_where_it_reads(monkeypatch):
    """The document's own list of changes is never walked."""
    asked = []
    monkeypatch.setattr(mac.backend, "word", SimpleNamespace(document=lambda: "doc"))
    monkeypatch.setattr(mac, "_count", lambda doc: 500)
    monkeypatch.setattr(mac, "_has_revisions", lambda doc: True)
    monkeypatch.setattr(mac, "_info", lambda doc, i, detail=False, spans=None: (
        asked.append((i, spans)) or {"index": i}))
    monkeypatch.setattr(tables, "mac_mark_tables", lambda doc, paragraphs: None)
    mac.get_paragraphs(3, 4, False)
    assert asked == [(3, None), (4, None)]  # None: each reads its own paragraph's


def test_a_mac_read_with_no_changes_asks_once(monkeypatch):
    asked = []
    monkeypatch.setattr(mac.backend, "word", SimpleNamespace(document=lambda: "doc"))
    monkeypatch.setattr(mac, "_count", lambda doc: 500)
    monkeypatch.setattr(mac, "_has_revisions", lambda doc: False)
    monkeypatch.setattr(mac, "_info", lambda doc, i, detail=False, spans=None: (
        asked.append(spans) or {"index": i}))
    monkeypatch.setattr(tables, "mac_mark_tables", lambda doc, paragraphs: None)
    mac.get_paragraphs(3, 4, False)
    assert asked == [[], []]


def _mac_doc(deleted):
    """create_range answers a range holding the deletion, if the match lies in one."""
    def create_range(start, end_):
        return _mac_range([("delete", a, b) for a, b in deleted if a <= start and end_ <= b])
    return SimpleNamespace(create_range=create_range)


def test_a_match_marked_deleted_is_left_out_at_its_own_place(monkeypatch):
    monkeypatch.setattr(mac, "_k", _mac_kinds)
    monkeypatch.setattr(mac, "_has_revisions", lambda doc: True)
    doc = _mac_doc([(10, 20)])
    assert mac._live(doc, [(1, 3), (12, 15), (30, 33)], time.monotonic() + 60) == (
        [(1, 3), (30, 33)], True)


def test_a_match_not_reached_in_time_is_not_taken_as_live(monkeypatch):
    monkeypatch.setattr(mac, "_k", _mac_kinds)
    monkeypatch.setattr(mac, "_has_revisions", lambda doc: True)
    assert mac._live(_mac_doc([]), [(1, 3), (5, 7)], time.monotonic() - 1) == ([], False)


def test_a_document_with_no_changes_is_not_asked_about_each_match(monkeypatch):
    monkeypatch.setattr(mac, "_has_revisions", lambda doc: False)
    assert mac._live(None, [(1, 3)], time.monotonic() - 1) == ([(1, 3)], True)


# -- a listing that runs out of time ------------------------------------------------

def test_a_listing_answers_what_it_has_when_time_runs_out(monkeypatch):
    monkeypatch.setattr(revisions, "LIST_SECONDS", -1.0)
    out = revisions.listing(lambda i: {"number": i}, 300, True, 181, 100)
    assert [r["number"] for r in out["revisions"]] == [181]
    assert out["complete"] is False
    assert "start=182" in out["warnings"][0] and "start=182" in out["more"]


def test_a_listing_in_time_says_only_what_follows():
    out = revisions.listing(lambda i: {"number": i}, 5, False, 2, 2)
    assert [r["number"] for r in out["revisions"]] == [2, 3]
    assert "complete" not in out and "warnings" not in out
    assert "start=4" in out["more"]


def test_the_last_page_says_nothing_follows():
    out = revisions.listing(lambda i: {"number": i}, 3, False, 2, 100)
    assert [r["number"] for r in out["revisions"]] == [2, 3]
    assert "more" not in out and "complete" not in out


def test_a_failure_other_than_the_end_is_not_taken_as_the_end(monkeypatch):
    monkeypatch.setattr(mac, "_k", _mac_kinds)

    class TimedOut(Exception):
        errornumber = -1712

    def late():
        raise TimedOut("Apple event timed out.")

    rng = SimpleNamespace(revisions={1: SimpleNamespace(revision_type=SimpleNamespace(get=late))})
    with pytest.raises(TimedOut):
        mac._deletions(rng)
