import json

import pytest

from ms_word_mcp.tools import _common, lists


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def apply(start, end, levels, level=None, continue_previous=False):
    return answer(lists._apply_list_impl, start, end, levels, level, continue_previous)


def level(**kw):
    return lists.LevelSpec(**kw).model_dump()


def test_a_bulleted_list_with_its_own_bullet_color_and_size(make_doc):
    make_doc("one", "two", "three")
    out = apply(1, 3, [level(bullet="●", font={"color": "#C00000", "size": 8},
                             number_position=0, text_position=12)])
    assert "warnings" not in out, out.get("warnings")
    assert [p["list_string"] for p in out["paragraphs"]] == ["●", "●", "●"]
    got = out["levels"][0]
    assert got["bullet"] == "●" and got["font"]["color"] == "#C00000"
    assert got["font"]["size"] == 8 and got["text_position"] == 12


def test_a_numbered_list_reads_back_its_numbers(make_doc):
    make_doc("a", "b")
    out = apply(1, 2, [level(number_style="arabic", number_format="(%1)")])
    assert [p["list_string"] for p in out["paragraphs"]] == ["(1)", "(2)"]
    assert out["paragraphs"][0]["list"] == "simple_numbering"


def test_each_list_gets_its_own_template(word, make_doc):
    doc = make_doc("a", "b", "c")
    apply(1, 1, [level(bullet="●")])
    apply(3, 3, [level(bullet="■")])
    first = doc.paragraph_objects()[0].Range.ListFormat.ListString
    assert first == "●", "the second list must not change the first"
    assert len(doc.ListTemplates.items) == 2


def test_set_level_and_remove(make_doc):
    make_doc("a", "b")
    apply(1, 2, [level(bullet="●"), level(level=2, bullet="○")])
    out = answer(lists._set_list_level_impl, 2, None, 2)
    assert out["paragraphs"][0]["level"] == 2 and out["paragraphs"][0]["list_string"] == "○"
    out = answer(lists._remove_list_impl, 1, 2)
    assert [p["list"] for p in out["paragraphs"]] == ["none", "none"]


def test_set_level_refuses_a_paragraph_outside_a_list(make_doc):
    make_doc("a")
    assert "not in a list" in answer(lists._set_list_level_impl, 1, None, 2)["error"]


def test_get_list_gives_every_level(make_doc):
    make_doc("a")
    apply(1, 1, [level(bullet="✓")])
    out = answer(lists._get_list_impl, 1)
    assert out["paragraph"]["list"] == "bullet" and len(out["levels"]) == 9
    assert out["levels"][0]["bullet"] == "✓"


def test_level_spec_checks():
    with pytest.raises(ValueError):
        lists.LevelSpec(bullet="●", number_style="arabic")
    with pytest.raises(ValueError):
        lists.LevelSpec(bullet="●", number_format="%1.")
    with pytest.raises(ValueError):
        lists.LevelSpec(bullet="●●")
    with pytest.raises(ValueError):
        lists.ApplyListInput(start=1, levels=[{"level": 1}, {"level": 1}])


def test_a_bulleted_level_of_a_multilevel_list_is_reported_as_a_bullet():
    assert lists.list_kind("outline_numbering", True) == "bullet"
    assert lists.list_kind("outline_numbering", False) == "outline_numbering"
    assert lists.list_kind("simple_numbering", False) == "simple_numbering"
