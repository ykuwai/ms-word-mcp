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
