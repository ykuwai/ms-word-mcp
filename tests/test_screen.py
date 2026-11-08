import struct
import zlib

import pytest

from ms_word_mcp.tools import screen


def test_png_is_well_formed():
    data = screen.png(2, 1, [b"\xff\x00\x00\x00\xff\x00"])
    assert data.startswith(b"\x89PNG\r\n\x1a\n")
    assert struct.unpack(">II", data[16:24]) == (2, 1)
    idat = data[data.index(b"IDAT") + 4:data.index(b"IEND") - 8]
    assert zlib.decompress(idat) == b"\x00\xff\x00\x00\x00\xff\x00"


def test_page_or_paragraph_not_both():
    with pytest.raises(ValueError):
        screen.PageImageInput(page=1, paragraph=2)
    screen.PageImageInput()
