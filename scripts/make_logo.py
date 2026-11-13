"""Write the Word MCP logos as plain vector SVG (no filters), text outlined from Ninter Bold.

usage: uv run --no-project --with fonttools --with uharfbuzz python scripts/make_logo.py assets

Ninter is an Inter-based typeface installed in ~/Library/Fonts; the SVGs hold the
outlines, so nothing else needs the font.
"""
import sys
from pathlib import Path

import uharfbuzz as hb
from fontTools.pens.svgPathPen import SVGPathPen
from fontTools.pens.transformPen import TransformPen
from fontTools.pens.boundsPen import BoundsPen
from fontTools.ttLib import TTFont

FONT = Path.home() / "Library/Fonts/Ninter-Bold.ttf"
OUT = Path(sys.argv[1])

# The MCP mark, from the Model Context Protocol logo.
MARK = [
    "M25 97.8528L92.8823 29.9706C102.255 20.598 117.451 20.598 126.823 29.9706V29.9706C136.196 39.3431 136.196 54.5391 126.823 63.9117L75.5581 115.177",
    "M76.2653 114.47L126.823 63.9117C136.196 54.5391 151.392 54.5391 160.765 63.9117L161.118 64.2652C170.491 73.6378 170.491 88.8338 161.118 98.2063L99.7248 159.6C96.6006 162.724 96.6006 167.789 99.7248 170.913L112.331 183.52",
    "M109.853 46.9411L59.6482 97.1457C50.2757 106.518 50.2757 121.714 59.6482 131.087V131.087C69.0208 140.459 84.2168 140.459 93.5894 131.087L143.794 80.8822",
]
CX, CY, R = 97.7, 102.5, 108.6


def defs(p):
    """Gradients for the disc; ids prefixed so two logos can share a page."""
    return f"""<defs>
    <radialGradient id="{p}base" cx="{CX + R * 0.6:.1f}" cy="{CY - R * 0.64:.1f}" r="{R * 2.1:.1f}" gradientUnits="userSpaceOnUse">
      <stop offset="0" stop-color="#BDE6FF"/>
      <stop offset="0.18" stop-color="#5DB6F5"/>
      <stop offset="0.45" stop-color="#2B7CD3"/>
      <stop offset="0.75" stop-color="#185ABD"/>
      <stop offset="1" stop-color="#0E3A86"/>
    </radialGradient>
    <radialGradient id="{p}rim" cx="{CX}" cy="{CY}" r="{R}" gradientUnits="userSpaceOnUse">
      <stop offset="0.82" stop-color="#3A8FE6" stop-opacity="0"/>
      <stop offset="1" stop-color="#3A8FE6" stop-opacity="0.5"/>
    </radialGradient>
    <linearGradient id="{p}rimFade" x1="{CX + R:.1f}" y1="{CY - R:.1f}" x2="{CX - R:.1f}" y2="{CY + R:.1f}" gradientUnits="userSpaceOnUse">
      <stop offset="0.45" stop-color="#fff" stop-opacity="0"/>
      <stop offset="1" stop-color="#fff"/>
    </linearGradient>
    <mask id="{p}rimMask"><circle cx="{CX}" cy="{CY}" r="{R}" fill="url(#{p}rimFade)"/></mask>
    <radialGradient id="{p}shadow" cx="{CX - 3}" cy="{CY + 6}" r="{R + 9}" gradientUnits="userSpaceOnUse">
      <stop offset="0.86" stop-color="#0A1E3C" stop-opacity="0.30"/>
      <stop offset="0.93" stop-color="#0A1E3C" stop-opacity="0.12"/>
      <stop offset="1" stop-color="#0A1E3C" stop-opacity="0"/>
    </radialGradient>
    <clipPath id="{p}disc"><circle cx="{CX}" cy="{CY}" r="{R}"/></clipPath>
  </defs>"""


def icon(p):
    paths = "".join(f'<path d="{d}"/>' for d in MARK)
    # A soft shadow under the line without a filter: wider, faint copies, offset toward the bottom left.
    shadow = "".join(
        f'<g opacity="{o}" stroke-width="{w}" transform="translate(-1.6 2.4)">{paths}</g>'
        for w, o in ((17.5, 0.10), (15.5, 0.12), (13.8, 0.16)))
    return f"""<circle cx="{CX - 3}" cy="{CY + 6}" r="{R + 9}" fill="url(#{p}shadow)"/>
  <circle cx="{CX}" cy="{CY}" r="{R}" fill="url(#{p}base)"/>
  <g clip-path="url(#{p}disc)">
    <circle cx="{CX}" cy="{CY}" r="{R}" fill="url(#{p}rim)" mask="url(#{p}rimMask)"/>
    <circle cx="170" cy="40" r="90" fill="#fff" fill-opacity="0.10"/>
    <g fill="none" stroke="#06214F" stroke-linecap="round">{shadow}</g>
    <g fill="none" stroke="#fff" stroke-width="12" stroke-linecap="round">{paths}</g>
  </g>"""


def text_path(text, size, x, y, tracking):
    """Outline `text` set in Ninter Bold at `size`, baseline at (x, y)."""
    blob = hb.Blob.from_file_path(str(FONT))
    face = hb.Face(blob)
    font = hb.Font(face)
    buf = hb.Buffer()
    buf.add_str(text)
    buf.guess_segment_properties()
    hb.shape(font, buf, {"kern": True, "liga": True})
    tt = TTFont(FONT)
    gs = tt.getGlyphSet()
    order = tt.getGlyphOrder()
    scale = size / tt["head"].unitsPerEm
    pen = SVGPathPen(gs)
    bounds = BoundsPen(gs)
    pen_x = 0
    for info, pos in zip(buf.glyph_infos, buf.glyph_positions):
        name = order[info.codepoint]
        t = (scale, 0, 0, -scale, x + (pen_x + pos.x_offset) * scale, y - pos.y_offset * scale)
        gs[name].draw(TransformPen(pen, t))
        gs[name].draw(TransformPen(bounds, t))
        pen_x += pos.x_advance + tracking / scale
    return pen.getCommands(), bounds.bounds


def write(name, body, viewbox, style=""):
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{viewbox}" role="img" '
           f'aria-label="Word MCP">{style}\n  {body}\n</svg>\n')
    (OUT / name).write_text(svg)


# The icon alone: the disc, its shadow and a little room.
write("word-mcp-logo.svg", defs("i") + "\n  " + icon("i"), "-24 -12 240 240")

# The icon with "Word / MCP" beside it. Units: the icon's, scaled to a disc 860 across.
k = 860 / (2 * R)
size, tracking = 275, -5.5
cap = 0.727 * size
gap = 0.11 * 860
x0 = 500 + 430 + gap
# The block runs from the cap top of "Word" to the baseline of "MCP": centre it on the disc.
top_to_base2 = cap + size
base1 = 500 - top_to_base2 / 2 + cap
d1, b1 = text_path("Word", size, x0, base1, tracking)
d2, b2 = text_path("MCP", size, x0, base1 + size, tracking)
right = max(b1[2], b2[2]) + 70
style = ('<style>.t{fill:#323232}@media (prefers-color-scheme:dark){.t{fill:#F2F2F2}}</style>')
body = (defs("l") + f'\n  <g transform="translate({500 - CX * k:.3f} {500 - CY * k:.3f}) scale({k:.5f})">'
        + icon("l") + "</g>\n"
        + f'  <path class="t" d="{d1}"/>\n  <path class="t" d="{d2}"/>')
write("word-mcp-logo-letter.svg", body, f"0 0 {right:.0f} 1000", style)
print("letter viewBox width", round(right), "text bounds", b1, b2)
