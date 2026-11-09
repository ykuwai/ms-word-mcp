"""What a crop of 5 points does to a picture that has been scaled, in line and floating.

    uv run --no-sync python scripts/probes/picture_crop.py

A scratch document, closed unsaved, and one small picture, removed at the end.
On the Mac the picture goes in Word's container, as the tools put it there.
"""

import os
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))

from ms_word_mcp import backend  # noqa: E402
from ms_word_mcp.tools import _common, screen  # noqa: E402

PNG = screen.png(60, 30, [b"\x40\x80\xc0" * 60] * 30)


def windows(path):
    doc = backend.word.get_app().Documents.Add()
    try:
        doc.Content.Text = "One paragraph.\rTwo paragraphs.\r"
        inline = doc.InlineShapes.AddPicture(path, False, True, doc.Range(0, 0))
        print("in line, natural", inline.Width, inline.Height)
        inline.LockAspectRatio = 0
        inline.Width, inline.Height = inline.Width * 2, inline.Height * 2
        print("in line, doubled", inline.Width, inline.Height)
        inline.PictureFormat.CropLeft = 5
        print("in line, crop 5", inline.PictureFormat.CropLeft, inline.Width, inline.Height)
        inline.PictureFormat.CropLeft = 0
        print("in line, crop 0", inline.PictureFormat.CropLeft, inline.Width, inline.Height)
        s = inline.ConvertToShape()
        print("floating, as it came", s.Width, s.Height)
        s.PictureFormat.CropLeft = 5
        print("floating, crop 5", s.PictureFormat.CropLeft, s.Width, s.Height)
        s.PictureFormat.CropLeft = 0
        print("floating, crop 0", s.PictureFormat.CropLeft, s.Width, s.Height)
        s.PictureFormat.CropTop = 4
        print("floating, crop top 4", s.PictureFormat.CropTop, s.Width, s.Height)
        return {"done": True}
    finally:
        doc.Close(0)


def mac(path):
    import appscript
    from appscript import k

    from ms_word_mcp.backend.mac import _decode_wide_integers

    _decode_wide_integers()
    w = appscript.app(id="com.microsoft.Word", terms="sdef")
    d = w.create_new_document()
    try:
        d.text_object.content.set("One paragraph.\rTwo paragraphs.\r")
        d.make(new=k.inline_picture, at=d.create_range(start=0, end_=0), with_properties={
            k.file_name: path, k.link_to_file: False, k.save_with_document: True})
        i = d.inline_shapes[1]
        print("in line, natural", i.width.get(), i.height.get())
        i.lock_aspect_ratio.set(False)
        i.width.set(i.width.get() * 2)
        i.height.set(i.height.get() * 2)
        print("in line, doubled", i.width.get(), i.height.get())
        i.picture_format.crop_left.set(5)
        print("in line, crop 5", i.picture_format.crop_left.get(), i.width.get(), i.height.get())
        i.picture_format.crop_left.set(0)
        print("in line, crop 0", i.picture_format.crop_left.get(), i.width.get(), i.height.get())
        i.convert_to_shape()
        name = d.shapes[1].name.get()
        s = d.shapes[name]
        print("floating, as it came", name, s.width.get(), s.height.get())
        s.picture_format.crop_left.set(5)
        print("floating, crop 5", s.picture_format.crop_left.get(), s.width.get(), s.height.get())
        s.picture_format.crop_left.set(0)
        print("floating, crop 0", s.picture_format.crop_left.get(), s.width.get(), s.height.get())
        s.picture_format.crop_top.set(4)
        print("floating, crop top 4", s.picture_format.crop_top.get(), s.width.get(), s.height.get())
        return {"done": True}
    finally:
        d.close(saving=k.no)


if backend.IS_MACOS:
    folder = os.path.expanduser("~/Library/Containers/com.microsoft.Word/Data/Documents")
else:
    folder = tempfile.gettempdir()
picture = os.path.join(folder, "picture-crop-probe.png")
with open(picture, "wb") as f:
    f.write(PNG)
try:
    if backend.IS_MACOS:
        mac(picture)
    else:
        backend.word.start()
        try:
            print(_common.run(windows, picture))
        finally:
            backend.word.stop()
finally:
    os.remove(picture)
