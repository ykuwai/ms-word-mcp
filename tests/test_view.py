import json

import pytest

from ms_word_mcp.tools import _common, view


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def test_zoom_to_120_and_fit_page_width(make_doc):
    doc = make_doc("a")
    out = answer(view._set_view_impl, 120, None, None)
    assert out["zoom"] == 120 and out["fit"] == "none" and "warnings" not in out
    out = answer(view._set_view_impl, None, "page_width", "print")
    assert out["fit"] == "page_width" and out["view"] == "print"
    assert doc.window.View.Zoom.PageFit == 2


def test_view_model_checks():
    with pytest.raises(ValueError):
        view.SetViewInput()
    with pytest.raises(ValueError):
        view.SetViewInput(zoom=120, fit="page_width")
    with pytest.raises(ValueError):
        view.SetViewInput(zoom=5)


def test_word_for_macs_other_names_for_a_view_read_as_that_view():
    assert view.MAC_SAME_VIEW == {"online_view": "web", "normal_view": "draft", "page_view": "print"}


def test_a_fit_that_needs_print_view_says_so():
    assert "need print view" in view.fit_warning("whole_page", "page_width")
    assert "need print view" in view.fit_warning("text_width", "page_width")
    assert view.fit_warning("page_width", "none") == "The fit reads back as 'none'."


def test_reading_view_is_refused_on_the_mac_before_anything_changes(monkeypatch):
    from ms_word_mcp.tools._common import MAC_IMPLS

    monkeypatch.setattr(view.backend, "word", None)  # touched only after the check
    with pytest.raises(ValueError, match="no reading view"):
        MAC_IMPLS[view._set_view_impl.__name__](120, None, "reading")
