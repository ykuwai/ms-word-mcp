"""How the document is shown: zoom and view type.

Constants are from Microsoft's pages for WdViewType and WdPageFit
(https://learn.microsoft.com/en-us/office/vba/api/word.wdviewtype,
https://learn.microsoft.com/en-us/office/vba/api/word.wdpagefit); the Mac
names are from Word.sdef.
"""

from __future__ import annotations

from typing import Literal

from pydantic import Field, model_validator

from .. import backend
from ._common import READ_ONLY, WRITES, Input, annotations, call, mac_impl

VIEW_TYPES = {"print": 3, "web": 6, "draft": 1, "outline": 2, "reading": 7}
VIEW_TYPES_MAC = {"print": "print_view", "web": "web_view", "draft": "draft_view",
                  "outline": "outline_view", "reading": "reading_view"}
PAGE_FITS = {"none": 0, "whole_page": 1, "page_width": 2, "text_width": 3}
PAGE_FITS_MAC = {"none": "page_fit_none", "whole_page": "page_fit_full_page",
                 "page_width": "page_fit_best_fit", "text_width": "page_fit_text_fit"}


class SetViewInput(Input):
    zoom: int | None = Field(None, ge=10, le=500, description=(
        "Zoom percent, such as 120."))
    fit: Literal[tuple(PAGE_FITS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "Fit the page instead of a fixed zoom: page_width, whole_page (print "
        "view only), text_width, or none to keep a fixed zoom."))
    view: Literal[tuple(VIEW_TYPES)] | None = Field(None, description=(  # type: ignore[valid-type]
        "print (page layout), web, draft, outline or reading."))

    @model_validator(mode="after")
    def _something(self):
        if self.zoom is None and self.fit is None and self.view is None:
            raise ValueError("Give zoom, fit or view.")
        if self.zoom is not None and self.fit not in (None, "none"):
            raise ValueError("Give zoom or fit, not both: a fit sets the zoom itself.")
        return self


def _window(doc):
    windows = doc.Windows
    if not windows.Count:
        raise ValueError("This document has no window (it was opened without one).")
    return windows.Item(1)


def _read(view) -> dict:
    zoom = view.Zoom
    return {"view": next((k for k, v in VIEW_TYPES.items() if v == view.Type), view.Type),
            "zoom": zoom.Percentage,
            "fit": next((k for k, v in PAGE_FITS.items() if v == zoom.PageFit), zoom.PageFit)}


def _get_view_impl():
    doc = backend.word.document()
    return {"success": True, **_read(_window(doc).View)}


def _set_view_impl(zoom, fit, view):
    doc = backend.word.document()
    target = _window(doc).View
    if view is not None:
        target.Type = VIEW_TYPES[view]
    if fit is not None:
        target.Zoom.PageFit = PAGE_FITS[fit]
    if zoom is not None:
        target.Zoom.PageFit = PAGE_FITS["none"]
        target.Zoom.Percentage = zoom
    got = _read(target)
    result = {"success": True, **got}
    warnings = []
    if view is not None and got["view"] != view:
        warnings.append(f"The view reads back as {got['view']!r}.")
    if zoom is not None and got["zoom"] != zoom:
        warnings.append(f"The zoom reads back as {got['zoom']}%.")
    if fit is not None and got["fit"] != fit:
        warnings.append(f"The fit reads back as {got['fit']!r}"
                        + (" (whole_page needs print view)." if fit == "whole_page" else "."))
    if warnings:
        result["warnings"] = warnings
    return result


# -- Mac -----------------------------------------------------------------------------

def _mac_read(view) -> dict:
    from .mac import _name

    zoom = view.zoom
    vt = _name(view.view_type.get())
    pf = _name(zoom.page_fit.get())
    return {"view": next((k for k, v in VIEW_TYPES_MAC.items() if v == vt),
                         "print" if vt == "page_view" else vt),
            "zoom": zoom.percentage.get(),
            "fit": next((k for k, v in PAGE_FITS_MAC.items() if v == pf), pf)}


@mac_impl(_get_view_impl)
def _mac_get_view():
    doc = backend.word.document()
    return {"success": True, **_mac_read(doc.active_window.view)}


@mac_impl(_set_view_impl)
def _mac_set_view(zoom, fit, view):
    from appscript import k

    doc = backend.word.document()
    target = doc.active_window.view
    if view is not None:
        target.view_type.set(getattr(k, VIEW_TYPES_MAC[view]))
    if fit is not None:
        target.zoom.page_fit.set(getattr(k, PAGE_FITS_MAC[fit]))
    if zoom is not None:
        target.zoom.page_fit.set(k.page_fit_none)
        target.zoom.percentage.set(zoom)
    got = _mac_read(target)
    result = {"success": True, **got}
    warnings = []
    if view is not None and got["view"] != view:
        warnings.append(f"The view reads back as {got['view']!r}.")
    if zoom is not None and got["zoom"] != zoom:
        warnings.append(f"The zoom reads back as {got['zoom']}%.")
    if fit is not None and got["fit"] != fit:
        warnings.append(f"The fit reads back as {got['fit']!r}.")
    if warnings:
        result["warnings"] = warnings
    return result


def register(mcp) -> None:
    @mcp.tool(name="word_get_view", annotations=annotations("Read the view", READ_ONLY))
    async def word_get_view() -> str:
        """The document window's view type, zoom percent and page fit."""
        return await call(_get_view_impl)

    @mcp.tool(name="word_set_view", annotations=annotations("Set the view", WRITES))
    async def word_set_view(params: SetViewInput) -> str:
        """Set the document window's zoom (such as 120%), fit it to the page
        width or the whole page, or switch to print, web, draft, outline or
        reading view. Changes only how the document is shown."""
        p = params
        return await call(_set_view_impl, p.zoom, p.fit, p.view)
