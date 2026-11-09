"""A picture of a page, as Word lays it out.

The model does not see Word. A picture of a page shows what reading the text
back cannot: a layout that broke, text that overflows, an equation that came
out wrong. It is made inside Word's own object model and the operating
system's own drawing, with no screen capture and so no screen-recording
permission:

- Windows: Page.EnhMetaFileBits gives the page as a metafile ("a picture
  representation of how a page of text appears",
  https://learn.microsoft.com/en-us/office/vba/api/word.page.enhmetafilebits),
  drawn into a PNG with GDI+.
- macOS: Word saves the document as a PDF (it keeps its own file), and
  Quartz draws the page into a PNG. Word for Mac's PDF shows comments and
  tracked changes in a margin (measured).
"""

from __future__ import annotations

import ctypes
import json
import os
import struct
import tempfile
import zlib

from pydantic import Field, model_validator

from .. import backend
from ._common import READ_ONLY, Input, annotations, call, mac_impl

WIDTH = 1200  # pixels across a page: small text stays legible
WD_ACTIVE_END_PAGE_NUMBER = 3  # WdInformation
WD_NUMBER_OF_PAGES = 4
WD_PRINT_VIEW = 3


def png(width: int, height: int, rows: list[bytes]) -> bytes:
    """A PNG of RGB rows (each 3 * width bytes), with no library."""
    def chunk(kind: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))

    raw = b"".join(b"\x00" + row for row in rows)
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


class PageImageInput(Input):
    page: int | None = Field(None, ge=1, description="The page, numbered from 1.")
    paragraph: int | None = Field(None, ge=1, description=(
        "Instead of `page`: the page this paragraph ends on."))

    @model_validator(mode="after")
    def _one(self):
        if self.page is not None and self.paragraph is not None:
            raise ValueError("Give page or paragraph, not both. Leave both out for page 1.")
        return self


def _which_page(doc, page, paragraph_index) -> tuple[int, int]:
    from ._word import paragraph

    total = doc.Content.Information(WD_NUMBER_OF_PAGES)
    if paragraph_index is not None:
        page = paragraph(doc, paragraph_index).Range.Information(WD_ACTIVE_END_PAGE_NUMBER)
    page = page or 1
    if page > total:
        raise ValueError(f"There is no page {page}: the document has {total} pages.")
    return page, total


class _GdiplusStartupInput(ctypes.Structure):
    _fields_ = [("GdiplusVersion", ctypes.c_uint32), ("DebugEventCallback", ctypes.c_void_p),
                ("SuppressBackgroundThread", ctypes.c_int),
                ("SuppressExternalCodecs", ctypes.c_int)]


class _Guid(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_uint32), ("Data2", ctypes.c_uint16),
                ("Data3", ctypes.c_uint16), ("Data4", ctypes.c_ubyte * 8)]


# The PNG encoder's CLSID, {557CF406-1A04-11D3-9A73-0000F81EF32E}.
_PNG_ENCODER = (0x557CF406, 0x1A04, 0x11D3, (0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E))
_PIXEL_FORMAT_24BPP_RGB = 0x00021808


def emf_to_png(emf: bytes, width: int, height: int, path: str) -> None:
    """Draw an enhanced metafile on white, width x height pixels, into a PNG."""
    gdi32, gdiplus = ctypes.windll.gdi32, ctypes.windll.gdiplus
    c_void_p = ctypes.c_void_p
    gdi32.SetEnhMetaFileBits.restype = c_void_p
    gdi32.SetEnhMetaFileBits.argtypes = [ctypes.c_uint, ctypes.c_char_p]
    token = ctypes.c_size_t()
    startup = _GdiplusStartupInput(1, None, 0, 0)
    if gdiplus.GdiplusStartup(ctypes.byref(token), ctypes.byref(startup), None):
        raise RuntimeError("GDI+ did not start.")
    metafile, bitmap, graphics = c_void_p(), c_void_p(), c_void_p()
    try:
        hemf = gdi32.SetEnhMetaFileBits(len(emf), emf)
        if not hemf:
            raise RuntimeError("Windows could not read the page's metafile.")
        if gdiplus.GdipCreateMetafileFromEmf(c_void_p(hemf), 1, ctypes.byref(metafile)):
            raise RuntimeError("GDI+ could not read the page's metafile.")
        if gdiplus.GdipCreateBitmapFromScan0(width, height, 0, _PIXEL_FORMAT_24BPP_RGB, None,
                                             ctypes.byref(bitmap)):
            raise RuntimeError("GDI+ could not make the picture.")
        gdiplus.GdipGetImageGraphicsContext(bitmap, ctypes.byref(graphics))
        gdiplus.GdipGraphicsClear(graphics, ctypes.c_uint32(0xFFFFFFFF))
        gdiplus.GdipSetSmoothingMode(graphics, 4)  # anti-alias
        gdiplus.GdipSetTextRenderingHint(graphics, 4)  # anti-alias, grid-fitted
        gdiplus.GdipSetInterpolationMode(graphics, 7)  # high-quality bicubic
        if gdiplus.GdipDrawImageRectI(graphics, metafile, 0, 0, width, height):
            raise RuntimeError("GDI+ could not draw the page.")
        d4 = (ctypes.c_ubyte * 8)(*_PNG_ENCODER[3])
        clsid = _Guid(_PNG_ENCODER[0], _PNG_ENCODER[1], _PNG_ENCODER[2], d4)
        if gdiplus.GdipSaveImageToFile(bitmap, ctypes.c_wchar_p(path), ctypes.byref(clsid),
                                       None):
            raise RuntimeError("GDI+ could not write the picture.")
    finally:
        if graphics:
            gdiplus.GdipDeleteGraphics(graphics)
        if bitmap:
            gdiplus.GdipDisposeImage(bitmap)
        if metafile:
            gdiplus.GdipDisposeImage(metafile)
        gdiplus.GdiplusShutdown(token)


def _new_png() -> str:
    """A file of its own, so that two calls at once cannot answer each other's page."""
    handle, path = tempfile.mkstemp(prefix="ms-word-mcp-page-", suffix=".png")
    os.close(handle)
    return path


def _get_page_image_impl(page, paragraph_index):
    doc = backend.word.document()
    if not doc.Windows.Count:
        raise ValueError("This document has no window, and a page is drawn from its "
                         "window. Open it with a window.")
    page, total = _which_page(doc, page, paragraph_index)
    window = doc.Windows.Item(1)
    view = window.View
    was = view.Type
    if was != WD_PRINT_VIEW:
        view.Type = WD_PRINT_VIEW  # pages exist in print layout only; put back below
    try:
        emf = bytes(window.Panes.Item(1).Pages.Item(page).EnhMetaFileBits)
    finally:
        if was != WD_PRINT_VIEW:
            view.Type = was
    setup = doc.PageSetup
    height = round(WIDTH * setup.PageHeight / setup.PageWidth)
    path = _new_png()
    emf_to_png(emf, WIDTH, height, path)
    return {"success": True, "page": page, "pages": total, "path": path,
            "width": WIDTH, "height": height}


CONTAINER = os.path.expanduser("~/Library/Containers/com.microsoft.Word/Data/Documents")


def _mac_pdf_folder(doc) -> str:
    """Where Word writes the PDF a page is drawn from.

    WORD_MAC_FILES_DIR if set; else the document's own folder, which Word has
    already been let into; else Word's container, which Word writes without
    asking but which macOS may ask about when the server reads it.
    """
    chosen = os.environ.get("WORD_MAC_FILES_DIR")
    if chosen:
        return chosen
    try:
        full = doc.posix_full_name.get()
    except Exception:  # noqa: BLE001
        full = None
    if isinstance(full, str) and os.path.isabs(full) and os.path.isdir(os.path.dirname(full)):
        return os.path.dirname(full)
    return CONTAINER


@mac_impl(_get_page_image_impl)
def _mac_get_page_image(page, paragraph_index):
    import Quartz
    from appscript import k
    from Foundation import NSURL

    from .mac import _para

    word = backend.word
    doc = word.document()
    total = doc.text_object.get_range_information(
        information_type=k.number_of_pages_in_document)
    if paragraph_index is not None:
        _para(doc, paragraph_index)
        page = doc.paragraphs[paragraph_index].text_object.get_range_information(
            information_type=k.active_end_page_number)
    page = int(page or 1)
    if page > int(total):
        raise ValueError(f"There is no page {page}: the document has {total} pages.")
    folder = _mac_pdf_folder(doc)
    pdf = os.path.join(folder, f"ms-word-mcp-page-{os.getpid()}.pdf")
    # Word for Mac is sandboxed and writes in its own folder without asking;
    # WORD_MAC_FILES_DIR names another folder Word has been allowed into, so
    # the server does not read Word's container. Saving as PDF does not change
    # the document's own file (measured).
    doc.save_as(file_name=pdf, file_format=k.format_PDF)
    try:
        url = NSURL.fileURLWithPath_(pdf)
        document = Quartz.CGPDFDocumentCreateWithURL(url)
        if document is None:
            raise RuntimeError("Word answered without error, but its PDF cannot be read.")
        pdf_page = Quartz.CGPDFDocumentGetPage(document, page)
        box = Quartz.CGPDFPageGetBoxRect(pdf_page, Quartz.kCGPDFMediaBox)
        scale = WIDTH / box.size.width
        height = round(box.size.height * scale)
        space = Quartz.CGColorSpaceCreateDeviceRGB()
        context = Quartz.CGBitmapContextCreate(None, WIDTH, height, 8, 0, space,
                                               Quartz.kCGImageAlphaNoneSkipLast)
        Quartz.CGContextSetRGBFillColor(context, 1, 1, 1, 1)
        Quartz.CGContextFillRect(context, Quartz.CGRectMake(0, 0, WIDTH, height))
        Quartz.CGContextScaleCTM(context, scale, scale)
        Quartz.CGContextDrawPDFPage(context, pdf_page)
        image = Quartz.CGBitmapContextCreateImage(context)
        path = _new_png()
        destination = Quartz.CGImageDestinationCreateWithURL(
            NSURL.fileURLWithPath_(path), "public.png", 1, None)
        Quartz.CGImageDestinationAddImage(destination, image, None)
        if not Quartz.CGImageDestinationFinalize(destination):
            raise RuntimeError("The page could not be written as a PNG.")
    finally:
        try:
            os.remove(pdf)
        except OSError:
            pass
    return {"success": True, "page": page, "pages": int(total), "path": path,
            "width": WIDTH, "height": height}


def register(mcp) -> None:
    try:
        from mcp.server.mcpserver import Image  # mcp >= 2.0
    except ModuleNotFoundError:
        from mcp.server.fastmcp import Image  # mcp 1.x

    @mcp.tool(name="word_get_page_image",
              annotations=annotations("Picture of a page", READ_ONLY))
    async def word_get_page_image(params: PageImageInput):
        """A picture of one page as Word lays it out, to look at after a change:
        it shows what reading the text back cannot (broken layout, overflowing
        text, an equation that came out wrong). Give `page`, or `paragraph`
        for the page that paragraph is on; page 1 if neither."""
        text = await call(_get_page_image_impl, params.page, params.paragraph)
        result = json.loads(text)
        if "error" in result:
            return text
        path = result.pop("path")
        try:
            with open(path, "rb") as f:
                data = f.read()
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
        return [json.dumps(result, ensure_ascii=False), Image(data=data, format="png")]
