"""Shapes, text boxes and lines: objects that float on the page.

Each one is anchored to a paragraph and stays on that paragraph's page. What
was measured on both platforms (docs/PROBES.md) decides how these are written:

- The Shapes collection runs in z-order, so a shape's place in it changes when
  it moves forward or back. A shape is found by its name, or by its number in
  reading order, and every answer gives the name.
- Word takes two shapes with one name and then finds the first, so a name
  that is already used is refused.
- The numbers given when a shape is made mean different things on each
  platform, so a shape is made first and placed after: what the position is
  measured from, and then the position.
- An anchor takes a position in the text that the text does not show.
- Word for Mac answers `make` and `convert` with a reference that can point at
  another shape, so a new shape is found by the name that was not there.
"""

from __future__ import annotations

import difflib
import os
from typing import Literal

from pydantic import Field, field_validator, model_validator

from .. import backend
from . import _props as P
from ._common import DESTRUCTIVE, READ_ONLY, WRITES, Input, Trimmed, annotations, call, follow, mac_impl
from ._word import ALIGNMENTS, clean_text, paragraph

# -- names ---------------------------------------------------------------------------

# MsoAutoShapeType in order: the value is the place in this list, from 1. These
# are also Word for Mac's names, after "autoshape ", so the two platforms share
# one table (https://learn.microsoft.com/en-us/office/vba/api/office.msoautoshapetype).
_AUTOSHAPES = (
    "rectangle|parallelogram|trapezoid|diamond|rounded rectangle|octagon|isosceles triangle|"
    "right triangle|oval|hexagon|cross|regular pentagon|can|cube|bevel|folded corner|"
    "smiley face|donut|no symbol|block arc|heart|lightning bolt|sun|moon|arc|double bracket|"
    "double brace|plaque|left bracket|right bracket|left brace|right brace|right arrow|"
    "left arrow|up arrow|down arrow|left right arrow|up down arrow|quad arrow|"
    "left right up arrow|bent arrow|U turn arrow|left up arrow|bent up arrow|"
    "curved right arrow|curved left arrow|curved up arrow|curved down arrow|"
    "striped right arrow|notched right arrow|pentagon|chevron|right arrow callout|"
    "left arrow callout|up arrow callout|down arrow callout|left right arrow callout|"
    "up down arrow callout|quad arrow callout|circular arrow|flowchart process|"
    "flowchart alternate process|flowchart decision|flowchart data|"
    "flowchart predefined process|flowchart internal storage|flowchart document|"
    "flowchart multi document|flowchart terminator|flowchart preparation|"
    "flowchart manual input|flowchart manual operation|flowchart connector|"
    "flowchart offpage connector|flowchart card|flowchart punched tape|"
    "flowchart summing junction|flowchart or|flowchart collate|flowchart sort|"
    "flowchart extract|flowchart merge|flowchart stored data|flowchart delay|"
    "flowchart sequential access storage|flowchart magnetic disk|"
    "flowchart direct access storage|flowchart display|explosion one|explosion two|"
    "four point star|five point star|eight point star|sixteen point star|"
    "twenty four point star|thirty two point star|up ribbon|down ribbon|curved up ribbon|"
    "curved down ribbon|vertical scroll|horizontal scroll|wave|double wave|"
    "rectangular callout|rounded rectangular callout|oval callout|cloud callout|"
    "line callout one|line callout two|line callout three|line callout four|"
    "line callout one accent bar|line callout two accent bar|line callout three accent bar|"
    "line callout four accent bar|line callout one no border|line callout two no border|"
    "line callout three no border|line callout four no border|"
    "callout one border and accent bar|callout two border and accent bar|"
    "callout three border and accent bar|callout four border and accent bar|"
    "action button custom|action button home|action button help|action button information|"
    "action button back or previous|action button forward or next|action button beginning|"
    "action button end|action button return|action button document|action button sound|"
    "action button movie|balloon|not primitive|flowchart offline storage|left right ribbon|"
    "diagonal stripe|pie|non isosceles trapezoid|Decagon|Heptagon|Dodecagon|six points star|"
    "seven points star|ten points star|twelve points star|round one rectangle|"
    "round two same rectangle|round two diagonal rectangle|snip round rectangle|"
    "snip one rectangle|snip two same rectangle|snip two diagonal rectangle|frame|half frame|"
    "tear|chord|corner|math plus|math minus|math multiply|math divide|math equal|"
    "math not equal|corner tabs|square tabs|plaque tabs|gear six|gear nine|funnel|pie wedge|"
    "left circular arrow|left right circular arrow|swoosh arrow|cloud|chart x|chart star|"
    "chart plus|line inverse").split("|")


def _friendly(word: str) -> str:
    name = word.lower().replace(" ", "_")
    return "home_plate" if name == "pentagon" else name  # the arrow-shaped one


#: A kind as the tools take it -> its MsoAutoShapeType.
KINDS = {_friendly(word): code for code, word in enumerate(_AUTOSHAPES, 1)
         if word not in ("not primitive", "balloon")}
_KIND_NAMES = {code: name for name, code in KINDS.items()}
ALIASES = {
    "triangle": "isosceles_triangle", "pentagon": "regular_pentagon", "circle": "oval",
    "ellipse": "oval", "square": "rectangle", "rect": "rectangle", "star": "five_point_star",
    "balloon": "rounded_rectangular_callout",  # Word answers it as this one
    "explosion": "explosion_one", "l_shape": "corner", "star_4point": "four_point_star",
    "star_5point": "five_point_star", "star_6point": "six_points_star",
    "star_7point": "seven_points_star", "star_8point": "eight_point_star",
    "star_10point": "ten_points_star", "star_12point": "twelve_points_star",
    "star_16point": "sixteen_point_star", "star_24point": "twenty_four_point_star",
    "star_32point": "thirty_two_point_star"}
COMMON_KINDS = ("rectangle, rounded_rectangle, oval, triangle, diamond, right_arrow, left_arrow, "
                "up_arrow, down_arrow, left_right_arrow, chevron, star, heart, cloud, "
                "rectangular_callout, rounded_rectangular_callout, flowchart_process, "
                "flowchart_decision, flowchart_terminator, frame")


def kind_name(value: str) -> str:
    """The kind a caller means, or a refusal that names the nearest ones."""
    key = value.strip().lower().replace(" ", "_").replace("-", "_")
    key = ALIASES.get(key, key)
    if key in KINDS:
        return key
    near = difflib.get_close_matches(key, [*KINDS, *ALIASES], n=4, cutoff=0.5)
    raise ValueError(f"There is no shape kind '{value}'. "
                     + (f"Did you mean {', '.join(near)}? " if near else "")
                     + f"Common kinds: {COMMON_KINDS}.")


def _mac_kind(kind: str):
    from appscript import k

    return getattr(k, "autoshape_" + _AUTOSHAPES[KINDS[kind] - 1].replace(" ", "_"))


def _mac_kind_name(value) -> str | None:
    from . import mac as m

    word = m._name(value).removeprefix("autoshape_").replace("_", " ")
    name = _friendly(word) if word in _AUTOSHAPES else None
    return name if name in KINDS else None


#: MsoShapeType -> (what the tools call it, Word for Mac's name).
TYPES = {1: ("autoshape", "shape_type_auto"), 2: ("callout", "shape_type_callout"),
         3: ("chart", "shape_type_chart"), 5: ("freeform", "shape_type_free_form"),
         6: ("group", "shape_type_group"), 7: ("object", "shape_type_embedded_OLE_control"),
         9: ("line", "shape_type_line"), 10: ("object", "shape_type_linked_OLE_object"),
         11: ("linked_picture", "shape_type_linked_picture"),
         13: ("picture", "shape_type_picture"), 15: ("word_art", "shape_type_word_art"),
         17: ("text_box", "shape_type_text_box"), 20: ("canvas", "shape_type_canvas"),
         24: ("smart_art", "shape_type_smartart_graphic"), 28: ("graphic", "shape_type_graphic")}
_MAC_TYPES = {mac: name for name, mac in TYPES.values()}
HOLDS_TEXT = ("text_box", "callout")
NO_TEXT = ("line", "picture", "linked_picture", "freeform", "group", "chart", "object",
           "word_art", "canvas", "smart_art", "graphic")

WRAPS = {"inline": 7, "square": 0, "tight": 1, "through": 2, "top_and_bottom": 4,
         "behind": 5, "front": 3}
WRAPS_MAC = {"inline": "wrap_inline", "square": "wrap_square", "tight": "wrap_tight",
             "through": "wrap_through", "top_and_bottom": "wrap_top_bottom",
             "behind": "wrap_behind", "front": "wrap_front"}
WRAP_SIDES = {"both": 0, "left": 1, "right": 2, "largest": 3}
HORIZONTAL = {"margin": 0, "page": 1, "column": 2, "character": 3, "left_margin": 4,
              "right_margin": 5, "inside_margin": 6, "outside_margin": 7}
VERTICAL = {"margin": 0, "page": 1, "paragraph": 2, "line": 3, "top_margin": 4,
            "bottom_margin": 5, "inside_margin": 6, "outside_margin": 7}
#: WdShapePosition (https://learn.microsoft.com/en-us/office/vba/api/word.wdshapeposition).
NAMED_LEFT = {"left": -999998, "center": -999995, "right": -999996, "inside": -999994,
              "outside": -999993}
NAMED_TOP = {"top": -999999, "center": -999995, "bottom": -999997, "inside": -999994,
             "outside": -999993}
Z_ORDERS = {"front": 0, "back": 1, "forward": 2, "backward": 3}
Z_ORDERS_MAC = {"front": "bring_shape_to_front", "back": "send_shape_to_back",
                "forward": "bring_shape_forward", "backward": "send_shape_backward"}
DASHES = {"solid": 1, "square_dot": 2, "round_dot": 3, "dash": 4, "dash_dot": 5,
          "dash_dot_dot": 6, "long_dash": 7, "long_dash_dot": 8, "long_dash_dot_dot": 9}
ARROWS = {"none": 1, "triangle": 2, "open": 3, "stealth": 4, "diamond": 5, "oval": 6}
ARROWS_MAC = {"none": "no_arrowhead", "triangle": "triangle_arrowhead",
              "open": "open_arrowhead", "stealth": "stealth_arrowhead",
              "diamond": "diamond_arrowhead", "oval": "oval_arrowhead"}
ORIENTATIONS = {"horizontal": 1, "vertical": 4, "rotate_right": 3, "rotate_left": 2}
ORIENTATIONS_MAC = {"horizontal": "horizontal", "vertical": "vertical_east_asian",
                    "rotate_right": "downward", "rotate_left": "upward"}
VERTICAL_ANCHORS = {"top": 1, "middle": 3, "bottom": 4}

ROUNDED = KINDS["rounded_rectangle"]
COLOR_OR_NONE = r"^(none|#?[0-9A-Fa-f]{6})$"
COLOR = r"^#?[0-9A-Fa-f]{6}$"
NO_ALT_ON_MAC = ("Word for Mac has no alternative text on a floating shape, so alt_text was "
                 "not set. It can be set on Windows.")
NO_MIDDLE_ON_MAC = ("Word for Mac has no vertical_anchor, so it was not set: the text stays "
                    "where it was. It can be set on Windows.")


# -- arguments ---------------------------------------------------------------------

class _Position(Input):
    left: float | Literal[tuple(NAMED_LEFT)] | None = Field(None, description=(  # type: ignore[valid-type]
        "Points from where horizontal_from says, or left, center, right, inside, outside."))
    top: float | Literal[tuple(NAMED_TOP)] | None = Field(None, description=(  # type: ignore[valid-type]
        "Points from where vertical_from says, or top, center, bottom, inside, outside."))
    horizontal_from: Literal[tuple(HORIZONTAL)] | None = Field(None, description=(  # type: ignore[valid-type]
        "What left is measured from. page and margin fix it on the page; column and "
        "character move with the text."))
    vertical_from: Literal[tuple(VERTICAL)] | None = Field(None, description=(  # type: ignore[valid-type]
        "What top is measured from. paragraph and line move with the text; page and "
        "margin fix it on the page."))


class _Wrap(Input):
    wrap: Literal[tuple(WRAPS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "How text goes round it. square, tight, through: text flows beside it. "
        "top_and_bottom: text above and below only. behind: behind the text (背面). "
        "front: over the text (前面), which hides what it covers. inline: in the line "
        "of text, like a character."))
    wrap_side: Literal[tuple(WRAP_SIDES)] | None = Field(None, description=(  # type: ignore[valid-type]
        "With square, tight or through, the sides text may use."))
    wrap_distance: float | None = Field(None, ge=0, le=1584, description=(
        "Points kept clear between the shape and the text, on all four sides."))


class _Look(Input):
    fill: Trimmed | None = Field(None, pattern=COLOR_OR_NONE, description="#RRGGBB, or none.")
    fill_transparency: float | None = Field(None, ge=0, le=1, description=(
        "0 is solid, 1 is clear."))
    line_color: Trimmed | None = Field(None, pattern=COLOR_OR_NONE, description=(
        "The outline: #RRGGBB, or none."))
    line_weight: float | None = Field(None, gt=0, le=1584, description="Points.")
    line_dash: Literal[tuple(DASHES)] | None = None  # type: ignore[valid-type]


class _Text(Input):
    text: str | None = Field(None, description=(
        "The text inside. \\n starts a new paragraph."))
    font: P.FontSpec | None = Field(None, description=(
        "The text's font, as in word_format_text: latin_font, east_asian_font, size, "
        "bold, color..."))
    alignment: Literal[tuple(ALIGNMENTS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "The alignment of the text's paragraphs."))
    vertical_anchor: Literal[tuple(VERTICAL_ANCHORS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "Where the text sits from top to bottom (Windows only)."))
    orientation: Literal[tuple(ORIENTATIONS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "vertical is East Asian vertical text (縦書き). rotate_right and rotate_left "
        "turn horizontal text a quarter turn."))


class _Crop(Input):
    crop_left: float | None = Field(None, ge=0, le=1584, description=(
        "Points to cut off the left of a picture, as it is shown now. 0 takes the cut away."))
    crop_top: float | None = Field(None, ge=0, le=1584, description="Likewise the top.")
    crop_right: float | None = Field(None, ge=0, le=1584, description="Likewise the right.")
    crop_bottom: float | None = Field(None, ge=0, le=1584, description="Likewise the bottom.")


class _Common(Input):
    name: Trimmed | None = Field(None, min_length=1, max_length=255, description=(
        "A name of its own in this document, to find it by later. Word gives one if "
        "this is left out."))
    alt_text: str | None = Field(None, description=(
        "What it shows, for readers who cannot see it (Windows only)."))


class _Made(_Position, _Wrap, _Common):
    index: int = Field(..., ge=1, description=(
        "The paragraph it is anchored to. It stays on that paragraph's page and is "
        "deleted with it."))

    @model_validator(mode="after")
    def _a_place_or_a_line(self):
        if self.wrap == "inline" and any(
                v is not None for v in (self.left, self.top, self.horizontal_from,
                                        self.vertical_from)):
            raise ValueError("A shape in line with the text has no position of its own: "
                             "leave out left, top, horizontal_from and vertical_from.")
        return self


class _Sized(_Made, _Look, _Text):
    width: float = Field(..., gt=0, le=1584, description="Points (72 = 1 inch).")
    height: float = Field(..., gt=0, le=1584, description="Points.")
    rotation: float | None = Field(None, ge=-360, le=360, description="Degrees, clockwise.")


class InsertShapeInput(_Sized):
    kind: str = Field(..., description=(
        f"Which shape. Common kinds: {COMMON_KINDS}. Any kind in Word's shape gallery "
        "works; a name close to one is answered with the nearest."))
    corner_radius: float | None = Field(None, ge=0, le=1, description=(
        "For a rounded_rectangle: 0 is square corners, 1 is as round as it goes."))

    @field_validator("kind")
    @classmethod
    def _known(cls, value: str) -> str:
        return kind_name(value)

    @model_validator(mode="after")
    def _rounded(self):
        if self.corner_radius is not None and self.kind != "rounded_rectangle":
            raise ValueError("corner_radius is for a rounded_rectangle.")
        return self


class InsertTextBoxInput(_Sized):
    pass


class InsertLineInput(_Made):
    begin_x: float = Field(..., description="Where it starts, in points from horizontal_from.")
    begin_y: float = Field(..., description="Where it starts, in points from vertical_from.")
    end_x: float = Field(..., description="Where it ends.")
    end_y: float = Field(..., description="Where it ends.")
    line_color: Trimmed | None = Field(None, pattern=COLOR, description="#RRGGBB.")
    line_weight: float | None = Field(None, gt=0, le=1584, description="Points.")
    line_dash: Literal[tuple(DASHES)] | None = None  # type: ignore[valid-type]
    begin_arrow: Literal[tuple(ARROWS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "The arrowhead at the start."))
    end_arrow: Literal[tuple(ARROWS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "The arrowhead at the end."))

    @model_validator(mode="after")
    def _a_line_or_a_point(self):
        if self.left is not None or self.top is not None:
            raise ValueError("A line is placed by begin_x, begin_y, end_x and end_y; leave "
                             "out left and top.")
        if self.begin_x == self.end_x and self.begin_y == self.end_y:
            raise ValueError("A line's two ends must be in different places.")
        return self


class InsertFloatingPictureInput(_Made, _Crop):
    path: Trimmed = Field(..., min_length=1, description="Full path of the picture file.")
    width: float | None = Field(None, gt=0, le=1584, description=(
        "Width in points. Given alone, the height follows in proportion."))
    height: float | None = Field(None, gt=0, le=1584, description=(
        "Height in points. Given alone, the width follows in proportion."))

    @field_validator("path")
    @classmethod
    def _absolute(cls, value: str) -> str:
        if not os.path.isabs(value):
            raise ValueError("path must be a full path.")
        return value

    @model_validator(mode="after")
    def _floats(self):
        if self.wrap == "inline":
            raise ValueError("A picture in line with the text is word_insert_picture's.")
        return self


class ListShapesInput(Input):
    start: int = Field(1, ge=1, description="The first shape to list, numbered from 1.")
    limit: int = Field(50, ge=1, le=200)


class ShapeRef(Input):
    shape: int | Trimmed = Field(..., description=(
        "Its name, or its number in reading order (word_list_shapes). The name stays "
        "the same while the numbers move."))

    @field_validator("shape")
    @classmethod
    def _real(cls, value):
        if isinstance(value, int) and value < 1:
            raise ValueError("Shapes are numbered from 1.")
        return value


class SetShapeInput(ShapeRef, _Position, _Wrap, _Common, _Look, _Text, _Crop):
    dleft: float | None = Field(None, description=(
        "Move it this many points right (negative: left) from where it is."))
    dtop: float | None = Field(None, description=(
        "Move it this many points down (negative: up) from where it is."))
    width: float | None = Field(None, gt=0, le=1584)
    height: float | None = Field(None, gt=0, le=1584)
    rotation: float | None = Field(None, ge=-360, le=360)
    z_order: Literal[tuple(Z_ORDERS)] | None = Field(None, description=(  # type: ignore[valid-type]
        "Among the shapes on the page. front and back go all the way; forward and "
        "backward go one place. wrap says whether it is over or behind the text."))
    kind: str | None = Field(None, description="Change it to another kind of shape.")
    corner_radius: float | None = Field(None, ge=0, le=1)
    begin_arrow: Literal[tuple(ARROWS)] | None = None  # type: ignore[valid-type]
    end_arrow: Literal[tuple(ARROWS)] | None = None  # type: ignore[valid-type]

    @field_validator("kind")
    @classmethod
    def _known(cls, value):
        return None if value is None else kind_name(value)

    @model_validator(mode="after")
    def _a_change(self):
        spec = self.model_dump(exclude_none=True, exclude={"shape"})
        if not spec:
            raise ValueError("Give at least one thing to change.")
        if self.left is not None and self.dleft is not None:
            raise ValueError("Give left or dleft, not both.")
        if self.top is not None and self.dtop is not None:
            raise ValueError("Give top or dtop, not both.")
        return self


class DeleteShapeInput(ShapeRef):
    pass


def spec_of(params: Input, *leave_out: str) -> dict:
    return params.model_dump(exclude_none=True, exclude=set(leave_out))


# -- shared between the platforms ----------------------------------------------------

def position_name(value: float, names: dict):
    """A position as the tools give it: a number, or left, center..."""
    for name, code in names.items():
        if abs(value - code) < 1:
            return name
    return round(value, 2)


def _moved(current, delta, what: str) -> float:
    if isinstance(current, str):
        raise ValueError(f"It is at {current}, not at a number of points, so it cannot be "
                         f"moved by {what}. Give the position itself.")
    return current + delta


def _plain_hex(value: str) -> str:
    return "none" if value == "none" else "#" + value.lstrip("#").upper()


def line_endpoints(begin_x, begin_y, end_x, end_y) -> dict:
    return {"left": min(begin_x, end_x), "top": min(begin_y, end_y),
            "width": abs(end_x - begin_x), "height": abs(end_y - begin_y)}


def read_back_warnings(spec: dict, got: dict) -> list[str]:
    out = []
    for key in ("left", "top", "width", "height", "rotation", "line_weight"):
        wanted, have = spec.get(key), got.get(key)
        if wanted is None or have is None:
            continue
        same = wanted == have if isinstance(wanted, str) or isinstance(have, str) \
            else abs(wanted - have) < 0.6
        if not same:
            out.append(f"{key} reads back as {have}, not {wanted}.")
    for key in ("wrap", "horizontal_from", "vertical_from", "line_dash", "begin_arrow",
                "end_arrow", "orientation"):
        if key in spec and key in got and spec[key] != got[key]:
            out.append(f"{key} reads back as {got[key]}, not {spec[key]}.")
    for key in ("fill", "line_color"):
        if key in spec and key in got and _plain_hex(spec[key]) != got[key]:
            out.append(f"{key} reads back as {got[key]}, not {_plain_hex(spec[key])}.")
    return out


def off_page_warnings(got: dict, page: dict) -> list[str]:
    """Whether a shape placed from the page or the margins leaves its page."""
    out = []
    for start, size, origin, margin, length, name in (
            ("left", "width", "horizontal_from", "left_margin", "page_width", "wide"),
            ("top", "height", "vertical_from", "top_margin", "page_height", "tall")):
        at = got.get(start)
        if not isinstance(at, (int, float)) or got.get(origin) not in ("page", "margin"):
            continue
        at += page[margin] if got[origin] == "margin" else 0
        if at < -0.6 or at + got[size] > page[length] + 0.6:
            out.append(f"It runs off the page: {start} {at:g} pt and {size} {got[size]:g} pt "
                       f"on a page {page[length]:g} pt {name}. Move or resize it with "
                       "word_set_shape.")
    return out


CROPS = ("crop_left", "crop_right", "crop_top", "crop_bottom")
CROP_MEMBERS = {"crop_left": "CropLeft", "crop_right": "CropRight", "crop_top": "CropTop",
                "crop_bottom": "CropBottom"}
NOT_A_PICTURE = "Only a picture can be cropped; this is a {kind}."


def apply_crop(spec: dict, read, write, size) -> None:
    """Crop by points as the picture is shown now.

    Word counts a crop against the picture's own size, so one shown at twice
    that loses 10 points for a crop of 5 (measured on both platforms). The
    scale cannot be read anywhere, so it is measured: a small crop, the
    change in the frame, and the crop put back. Both directions are checked
    before either is cropped.
    """
    planned = []
    for axis, sides in (("width", CROPS[:2]), ("height", CROPS[2:])):
        given = [side for side in sides if side in spec]
        if not given:
            continue
        shown, first = size(axis), read(given[0])
        step = 5.0 if shown >= 50 else 1.0
        write(given[0], first + step)
        scale = round((shown - size(axis)) / step, 3)
        write(given[0], first)
        if scale <= 0:
            raise ValueError("Word did not change the picture's frame for a small crop, so "
                             "how far it is scaled cannot be measured. Nothing was cropped.")
        whole = size(axis) + sum(read(side) for side in sides) * scale
        wanted = sum(spec.get(side, read(side) * scale) for side in sides)
        if wanted >= whole - 1:
            raise ValueError(f"Those crops leave nothing of the picture: it is {whole:g} pt "
                             "across before cropping. Nothing was cropped.")
        planned += [(side, spec[side] / scale) for side in given]
    for side, value in planned:
        write(side, value)


NOT_FOUND = "There is no shape named '{name}'. word_list_shapes lists them."
TWO_NAMES = ("{count} shapes are named '{name}'. Use its number (word_list_shapes), and "
             "give it a name of its own with `name`.")
OUT_OF_RANGE = ("There is no shape {number}: the document has {count} (numbered from 1; "
                "word_list_shapes lists them).")
TAKEN = "A shape is already named '{name}'. Give another name, or use word_set_shape on it."
NO_TEXT_IN = ("A {kind} holds no text. Only a text box or a drawn shape does; "
              "use word_insert_text_box beside it.")


def _check_text_allowed(kind: str, spec: dict) -> None:
    if kind in NO_TEXT and any(key in spec for key in TEXT_KEYS):
        raise ValueError(NO_TEXT_IN.format(kind=kind))


def _finished(got: dict, spec: dict, page: dict, warnings: list[str],
              overflowing: bool) -> dict:
    warnings = [*warnings, *read_back_warnings(spec, got), *off_page_warnings(got, page)]
    if overflowing:
        warnings.append("The text does not fit in the shape: make the shape larger or the "
                        "text smaller.")
    out = {"success": True, **got}
    if warnings:
        out["warnings"] = warnings
    return out


# -- Windows -------------------------------------------------------------------------

def _all_shapes(doc) -> list:
    return [doc.Shapes.Item(i) for i in range(1, doc.Shapes.Count + 1)]


def _in_reading_order(doc) -> list:
    return sorted(_all_shapes(doc), key=lambda s: s.Anchor.Start)


def _check_name(doc, name) -> None:
    if name is not None and any(s.Name == name for s in _all_shapes(doc)):
        raise ValueError(TAKEN.format(name=name))


def _find(doc, ref):
    if isinstance(ref, int):
        found = _in_reading_order(doc)
        if ref > len(found):
            raise ValueError(OUT_OF_RANGE.format(number=ref, count=len(found)))
        return found[ref - 1]
    found = [s for s in _all_shapes(doc) if s.Name == ref]
    if not found:
        raise ValueError(NOT_FOUND.format(name=ref))
    if len(found) > 1:
        raise ValueError(TWO_NAMES.format(count=len(found), name=ref))
    return found[0]


def _number_of(doc, shape) -> int:
    wanted = shape.ID
    return next(n for n, s in enumerate(_in_reading_order(doc), 1) if s.ID == wanted)


def _kind_of(s) -> str:
    code = s.Type
    if code in (1, 2):  # a line callout is its own type, with its own shape
        return _KIND_NAMES.get(s.AutoShapeType, TYPES[code][0])
    return TYPES.get(code, (f"type_{code}",))[0]


def _name_for(table: dict, code):
    return next((name for name, value in table.items() if value == code), None)


def _page_of(s) -> dict:
    setup = s.Anchor.Sections.Item(1).PageSetup
    return {"page_width": setup.PageWidth, "page_height": setup.PageHeight,
            "left_margin": setup.LeftMargin, "top_margin": setup.TopMargin}


def _describe(doc, s, number: int) -> dict:
    kind = _kind_of(s)
    wrap = s.WrapFormat
    out = {"number": number, "name": s.Name, "kind": kind,
           "index": doc.Range(0, s.Anchor.Start + 1).Paragraphs.Count,
           "page": s.Anchor.Information(3),
           "left": position_name(s.Left, NAMED_LEFT), "top": position_name(s.Top, NAMED_TOP),
           "width": round(s.Width, 2), "height": round(s.Height, 2),
           "horizontal_from": _name_for(HORIZONTAL, s.RelativeHorizontalPosition),
           "vertical_from": _name_for(VERTICAL, s.RelativeVerticalPosition),
           "wrap": _name_for(WRAPS, wrap.Type), "z": s.ZOrderPosition,
           "rotation": round(s.Rotation, 2), "alt_text": s.AlternativeText or ""}
    if kind not in ("picture", "linked_picture", "group", "canvas", "smart_art", "chart"):
        out["line_color"] = "none" if not s.Line.Visible else P.hex_of(s.Line.ForeColor.RGB)
        out["line_weight"] = round(s.Line.Weight, 2)
        out["line_dash"] = _name_for(DASHES, s.Line.DashStyle)
    if kind == "picture" and any(getattr(s.PictureFormat, m) for m in CROP_MEMBERS.values()):
        out["cropped"] = True
    if kind == "line":
        out["begin_arrow"] = _name_for(ARROWS, s.Line.BeginArrowheadStyle)
        out["end_arrow"] = _name_for(ARROWS, s.Line.EndArrowheadStyle)
    elif kind not in NO_TEXT:
        out["fill"] = "none" if not s.Fill.Visible else P.hex_of(s.Fill.ForeColor.RGB)
        if s.Fill.Visible and s.Fill.Transparency:
            out["fill_transparency"] = round(s.Fill.Transparency, 2)
        out["text"] = clean_text(s.TextFrame.TextRange.Text).replace("\r", "\n")
        out["orientation"] = _name_for(ORIENTATIONS, s.TextFrame.Orientation)
    return out


TEXT_KEYS = ("text", "font", "alignment", "vertical_anchor", "orientation")


def _apply_text(s, spec: dict) -> list[str]:
    if not any(key in spec for key in TEXT_KEYS):
        return []  # a line has no text frame to touch: Word refuses even the look
    warnings = []
    frame = s.TextFrame
    if "text" in spec:
        frame.TextRange.Text = spec["text"].replace("\n", "\r")
    text = frame.TextRange
    if spec.get("font"):
        wanted = P.with_latin_extended(P.given(spec["font"], P.FONT_NAMES))
        refused = P.com_apply(text.Font, P.FONT_PROPS, wanted)
        warnings += P.mismatches(wanted, P.com_read(text.Font, P.FONT_PROPS), refused)
    if "alignment" in spec:
        text.ParagraphFormat.Alignment = ALIGNMENTS[spec["alignment"]]
    if "orientation" in spec:
        frame.Orientation = ORIENTATIONS[spec["orientation"]]
    if "vertical_anchor" in spec:
        frame.VerticalAnchor = VERTICAL_ANCHORS[spec["vertical_anchor"]]
    return warnings


def _apply_look(s, spec: dict) -> None:
    if "fill" in spec:
        if spec["fill"] == "none":
            s.Fill.Visible = 0
        else:
            s.Fill.Visible = -1
            s.Fill.Solid()
            s.Fill.ForeColor.RGB = P.bgr(spec["fill"])
    if "fill_transparency" in spec:
        s.Fill.Transparency = spec["fill_transparency"]
    if "line_color" in spec:
        if spec["line_color"] == "none":
            s.Line.Visible = 0
        else:
            s.Line.Visible = -1
            s.Line.ForeColor.RGB = P.bgr(spec["line_color"])
    if "line_weight" in spec:
        s.Line.Weight = spec["line_weight"]
    if "line_dash" in spec:
        s.Line.DashStyle = DASHES[spec["line_dash"]]
    if "begin_arrow" in spec:
        s.Line.BeginArrowheadStyle = ARROWS[spec["begin_arrow"]]
    if "end_arrow" in spec:
        s.Line.EndArrowheadStyle = ARROWS[spec["end_arrow"]]


def _place(s, spec: dict) -> None:
    """In this order: what a position is measured from changes the number."""
    if "wrap" in spec:
        s.WrapFormat.Type = WRAPS[spec["wrap"]]
    if "wrap_side" in spec:
        s.WrapFormat.Side = WRAP_SIDES[spec["wrap_side"]]
    if "wrap_distance" in spec:
        wrap = s.WrapFormat
        wrap.DistanceLeft = wrap.DistanceRight = spec["wrap_distance"]
        wrap.DistanceTop = wrap.DistanceBottom = spec["wrap_distance"]
    if "horizontal_from" in spec:
        s.RelativeHorizontalPosition = HORIZONTAL[spec["horizontal_from"]]
    if "vertical_from" in spec:
        s.RelativeVerticalPosition = VERTICAL[spec["vertical_from"]]
    if "width" in spec or "height" in spec:
        s.LockAspectRatio = 0
        if "width" in spec:
            s.Width = spec["width"]
        if "height" in spec:
            s.Height = spec["height"]
    if "left" in spec:
        s.Left = NAMED_LEFT.get(spec["left"], spec["left"])
    if "dleft" in spec:
        s.Left = _moved(position_name(s.Left, NAMED_LEFT), spec["dleft"], "dleft")
    if "top" in spec:
        s.Top = NAMED_TOP.get(spec["top"], spec["top"])
    if "dtop" in spec:
        s.Top = _moved(position_name(s.Top, NAMED_TOP), spec["dtop"], "dtop")


def crop_picture(target, spec: dict) -> None:
    """Crop an inline picture or a floating one."""
    fmt = target.PictureFormat
    apply_crop(spec, lambda side: getattr(fmt, CROP_MEMBERS[side]),
               lambda side, value: setattr(fmt, CROP_MEMBERS[side], value),
               lambda axis: target.Width if axis == "width" else target.Height)


def _crop(s, spec: dict, kind: str) -> None:
    if any(side in spec for side in CROPS):
        if kind not in ("picture", "linked_picture"):
            raise ValueError(NOT_A_PICTURE.format(kind=kind))
        crop_picture(s, spec)


def _apply(s, spec: dict, kind: str) -> list[str]:
    _check_text_allowed(kind, spec)
    _crop(s, spec, kind)  # first: if it is refused, nothing else has changed
    warnings = []
    if "kind" in spec:
        s.AutoShapeType = KINDS[spec["kind"]]
    _place(s, spec)
    if "rotation" in spec:
        s.Rotation = spec["rotation"]
    if "corner_radius" in spec:
        if s.AutoShapeType == ROUNDED:
            s.Adjustments[1] = spec["corner_radius"] * 0.5
        else:
            warnings.append("corner_radius is for a rounded_rectangle, so it was not set.")
    _apply_look(s, spec)
    warnings += _apply_text(s, spec)
    if "alt_text" in spec:
        s.AlternativeText = spec["alt_text"]
    if "z_order" in spec:
        s.ZOrder(Z_ORDERS[spec["z_order"]])
    if "name" in spec:
        s.Name = spec["name"]
    return warnings


def _result(doc, s, spec: dict, warnings: list[str]) -> dict:
    follow(doc, s.Anchor)
    # TextFrame.Overflowing is not asked: on Windows it answered True for an
    # empty shape and for two short lines with room to spare (docs/PROBES.md).
    return _finished(_describe(doc, s, _number_of(doc, s)), spec, _page_of(s), warnings, False)


def _finish_new(doc, s, spec: dict, kind: str) -> dict:
    """Give a new shape its place and look. If that fails no half-made shape stays."""
    try:
        warnings = _apply(s, spec, kind)
    except Exception:
        s.Delete()
        raise
    return _result(doc, s, spec, warnings)


_PLACED = {"left": 0.0, "top": 0.0, "horizontal_from": "margin", "vertical_from": "paragraph"}


def _new_spec(spec: dict, wrap: str) -> dict:
    """What a new shape is given: a place unless it is in line."""
    full = {"wrap": wrap, **spec}
    return full if full["wrap"] == "inline" else {**_PLACED, **full}


def _insert_shape_impl(kind, index, width, height, spec):
    doc = backend.word.document()
    _check_name(doc, spec.get("name"))
    anchor = paragraph(doc, index).Range
    s = doc.Shapes.AddShape(KINDS[kind], 0, 0, width, height, anchor)
    full = _new_spec(spec, "square")
    return _finish_new(doc, s, full, kind)


def _insert_text_box_impl(index, width, height, spec):
    doc = backend.word.document()
    _check_name(doc, spec.get("name"))
    anchor = paragraph(doc, index).Range
    s = doc.Shapes.AddTextbox(1, 0, 0, width, height, anchor)
    full = _new_spec(spec, "square")
    return _finish_new(doc, s, full, "text_box")


def _insert_line_impl(index, begin_x, begin_y, end_x, end_y, spec):
    doc = backend.word.document()
    _check_name(doc, spec.get("name"))
    anchor = paragraph(doc, index).Range
    s = doc.Shapes.AddLine(begin_x, begin_y, end_x, end_y, anchor)
    full = {**_new_spec(spec, "front"), **line_endpoints(begin_x, begin_y, end_x, end_y)}
    return _finish_new(doc, s, full, "line")


def _picture_spec(spec: dict) -> dict:
    """What a picture gets once it floats: its alt text is set while it is in line."""
    return _new_spec({k: v for k, v in spec.items() if k != "alt_text"}, "square")


def float_picture(doc, inline, spec: dict) -> dict:
    """An inline picture out of the text, at the place `spec` gives."""
    s = inline.ConvertToShape()
    return {**_finish_new(doc, s, _picture_spec(spec), "picture"), "now": "floating",
            "note": "It floats now: word_list_shapes lists it, and word_set_shape changes it."}


def _insert_floating_picture_impl(path, index, width, height, spec):
    from .references import _apply_picture

    if not os.path.isfile(path):
        raise ValueError(f"There is no file at '{path}'.")
    doc = backend.word.document()
    _check_name(doc, spec.get("name"))
    start = paragraph(doc, index).Range.Start
    inline = doc.InlineShapes.AddPicture(path, False, True, doc.Range(start, start))
    warnings = _apply_picture(doc, inline, width, height, None, spec.get("alt_text"))
    s = inline.ConvertToShape()
    out = _finish_new(doc, s, _picture_spec(spec), "picture")
    if warnings:
        out["warnings"] = [*warnings, *out.get("warnings", [])]
    return out


def _list_shapes_impl(start, limit):
    doc = backend.word.document()
    found = _in_reading_order(doc)
    out = {"success": True, "count": len(found), "shapes": [
        _describe(doc, s, n) for n, s in enumerate(found, 1)
        if start <= n < start + limit]}
    if start + limit <= len(found):
        out["more"] = (f"Shapes {start + limit} to {len(found)} follow; "
                       f"pass start={start + limit}.")
    return out


def _set_shape_impl(shape, spec):
    doc = backend.word.document()
    s = _find(doc, shape)
    if "name" in spec and spec["name"] != s.Name:
        _check_name(doc, spec["name"])
    kind = _kind_of(s)
    if kind in ("picture", "linked_picture") and spec.get("wrap") == "inline":
        return _picture_in_line(doc, s, spec)
    return _result(doc, s, spec, _apply(s, spec, kind))


def _picture_in_line(doc, s, spec: dict):
    """A floating picture back in the text, where it is an inline picture."""
    from .references import _describe_picture

    rest = {k: v for k, v in spec.items() if k not in ("wrap", "wrap_side", "wrap_distance")}
    warnings = _apply(s, rest, _kind_of(s))
    inline = s.ConvertToInlineShape()
    follow(doc, inline.Range)
    number = doc.Range(0, inline.Range.Start + 1).InlineShapes.Count
    out = {"success": True, "now": "inline", **_describe_picture(doc, inline, number),
           "note": "It is an inline picture now: word_list_pictures lists it, and "
                   "word_set_picture changes it."}
    if warnings:
        out["warnings"] = warnings
    return out


def _delete_shape_impl(shape):
    doc = backend.word.document()
    s = _find(doc, shape)
    name, anchor = s.Name, s.Anchor
    follow(doc, anchor)
    s.Delete()
    return {"success": True, "deleted": name, "left": doc.Shapes.Count}


def leading_anchors(doc, start: int, end: int, hidden: int) -> int:
    """How many anchors sit at the very start of a span of text.

    Each takes a position there that the span's text does not show, and
    replacing them deletes their shapes. `hidden` is the number of positions
    the text does not show: it caps the count, since the collection may also
    answer anchors from other parts of the document.
    """
    if not hidden or not doc.Shapes.Count:
        return 0
    at = {s.Anchor.Start for s in _all_shapes(doc) if start <= s.Anchor.Start < end}
    count = 0
    while start + count in at and count < hidden:
        count += 1
    return count


# -- macOS ---------------------------------------------------------------------------

def _mac_names(doc) -> list[str]:
    from appscript import k

    return [doc.shapes[i].name.get() for i in range(1, int(doc.count(each=k.shape)) + 1)]


def _mac_in_reading_order(doc) -> list[str]:
    """(anchor, place in the collection, name) of each shape, in reading order."""
    names = _mac_names(doc)
    found = [(doc.shapes[i].anchor.start_of_content.get(), i, name)
             for i, name in enumerate(names, 1)]
    return sorted(found)


def _mac_check_name(doc, name) -> None:
    if name is not None and name in _mac_names(doc):
        raise ValueError(TAKEN.format(name=name))


def _mac_find(doc, ref):
    """A reference to the shape, and whether its name is its own.

    By name wherever it can be: a place in the collection changes when the
    shape moves in z-order or out of line with the text.
    """
    names = _mac_names(doc)
    if isinstance(ref, int):
        found = _mac_in_reading_order(doc)
        if ref > len(found):
            raise ValueError(OUT_OF_RANGE.format(number=ref, count=len(found)))
        _, place, name = found[ref - 1]
    else:
        count = names.count(ref)
        if not count:
            raise ValueError(NOT_FOUND.format(name=ref))
        if count > 1:
            raise ValueError(TWO_NAMES.format(count=count, name=ref))
        place, name = names.index(ref) + 1, ref
    if names.count(name) == 1:
        return doc.shapes[name], True
    return doc.shapes[place], False


def _mac_number_of(doc, name: str) -> int:
    return next(n for n, (_, _, have) in enumerate(_mac_in_reading_order(doc), 1) if have == name)


def _mac_made(doc, before: list[str], make):
    """Make a shape and find it by the name that was not there."""
    make()
    new = [n for n in _mac_names(doc) if n not in before]
    if len(new) != 1:
        raise RuntimeError("Word answered without error, but no new shape can be found. "
                           "Read the document back with word_list_shapes.")
    return new[0]


def _mac_kind_of(s) -> str:
    from . import mac as m

    type_name = m._name(s.shape_type.get())
    if type_name in ("shape_type_auto", "shape_type_callout"):
        return _mac_kind_name(s.auto_shape_type.get()) or _MAC_TYPES[type_name]
    return _MAC_TYPES.get(type_name, type_name.removeprefix("shape_type_"))


def _mac_key(table: dict, value, prefix: str = ""):
    from . import mac as m

    name = m._name(value).removeprefix(prefix)
    return next((key for key, mac in table.items() if mac == name), name)


_MAC_HORIZONTAL = {**{k: k for k in HORIZONTAL}, "inside_margin": "inner_margin",
                   "outside_margin": "outer_margin"}
_MAC_VERTICAL = {**{k: k for k in VERTICAL}, "inside_margin": "inner_margin",
                 "outside_margin": "outer_margin"}
_H_PREFIX, _V_PREFIX = "relative_horizontal_position_", "relative_vertical_position_"


def _mac_page(doc) -> dict:
    setup = doc.page_setup
    return {"page_width": setup.page_width.get(), "page_height": setup.page_height.get(),
            "left_margin": setup.left_margin.get(), "top_margin": setup.top_margin.get()}


def _mac_describe(doc, s, number: int) -> dict:
    from appscript import k

    from . import mac as m

    kind = _mac_kind_of(s)
    anchor = s.anchor
    page = anchor.get_range_information(information_type=k.active_end_page_number)
    out = {"number": number, "name": s.name.get(), "kind": kind,
           "index": m._index_at(doc, anchor.start_of_content.get()),
           "page": int(page) if str(page).isdigit() else page,
           "left": position_name(s.left_position.get(), NAMED_LEFT),
           "top": position_name(s.top.get(), NAMED_TOP),
           "width": round(s.width.get(), 2), "height": round(s.height.get(), 2),
           "horizontal_from": _mac_key(_MAC_HORIZONTAL, s.relative_horizontal_position.get(),
                                       _H_PREFIX),
           "vertical_from": _mac_key(_MAC_VERTICAL, s.relative_vertical_position.get(), _V_PREFIX),
           "wrap": _mac_key(WRAPS_MAC, s.wrap_format.wrap_type.get()),
           "z": s.z_order_position.get(), "rotation": round(s.rotation_.get(), 2)}
    if kind not in ("picture", "linked_picture", "group", "canvas", "smart_art", "chart"):
        line = s.line_format
        out["line_color"] = "none" if not line.visible.get() else P._mac_hex(line.fore_color.get())
        out["line_weight"] = round(line.weight.get(), 2)
        out["line_dash"] = m._name(line.dash_style_.get()).removeprefix("line_dash_style_")
    if kind == "picture" and any(getattr(s.picture_format, side).get() for side in CROPS):
        out["cropped"] = True
    if kind == "line":
        out["begin_arrow"] = _mac_key(ARROWS_MAC, s.line_format.begin_arrowhead_style.get())
        out["end_arrow"] = _mac_key(ARROWS_MAC, s.line_format.end_arrowhead_style.get())
    elif kind not in NO_TEXT:
        fill = s.fill_format
        out["fill"] = "none" if not fill.visible.get() else P._mac_hex(fill.fore_color.get())
        if fill.visible.get() and fill.transparency.get():
            out["fill_transparency"] = round(fill.transparency.get(), 2)
        out["text"] = m._clean(s.text_frame.text_range.content.get()).replace("\r", "\n")
        out["orientation"] = _mac_key(ORIENTATIONS_MAC, s.text_frame.orientation.get())
    return out


def _mac_apply_text(s, spec: dict) -> list[str]:
    from appscript import k

    if not any(key in spec for key in TEXT_KEYS):
        return []
    warnings = []
    frame = s.text_frame
    if "text" in spec:
        frame.text_range.content.set(spec["text"].replace("\n", "\r"))
    text = frame.text_range
    if spec.get("font"):
        wanted = P.with_latin_extended(P.given(spec["font"], P.FONT_NAMES))
        refused = P.mac_apply(text.font_object, P.FONT_PROPS, wanted)
        warnings += P.mismatches(wanted, P.mac_read(text.font_object, P.FONT_PROPS), refused)
    if "alignment" in spec:
        text.paragraph_format.alignment.set(
            getattr(k, P.ALIGNMENTS_MAC[spec["alignment"]]))
    if "orientation" in spec:
        frame.orientation.set(getattr(k, ORIENTATIONS_MAC[spec["orientation"]]))
    if "vertical_anchor" in spec:
        warnings.append(NO_MIDDLE_ON_MAC)
    return warnings


def _mac_apply_look(s, spec: dict) -> None:
    from appscript import k

    fill, line = s.fill_format, s.line_format
    if "fill" in spec:
        if spec["fill"] == "none":
            fill.visible.set(False)
        else:
            fill.visible.set(True)
            fill.fore_color.set(P._mac_rgb(spec["fill"]))
    if "fill_transparency" in spec:
        fill.transparency.set(spec["fill_transparency"])
    if "line_color" in spec:
        if spec["line_color"] == "none":
            line.visible.set(False)
        else:
            line.visible.set(True)
            line.fore_color.set(P._mac_rgb(spec["line_color"]))
    if "line_weight" in spec:
        line.weight.set(spec["line_weight"])
    if "line_dash" in spec:
        line.dash_style_.set(getattr(k, "line_dash_style_" + spec["line_dash"]))
    if "begin_arrow" in spec:
        line.begin_arrowhead_style.set(getattr(k, ARROWS_MAC[spec["begin_arrow"]]))
    if "end_arrow" in spec:
        line.end_arrowhead_style.set(getattr(k, ARROWS_MAC[spec["end_arrow"]]))


def _mac_place(s, spec: dict) -> None:
    from appscript import k

    wrap = s.wrap_format
    if "wrap" in spec:
        wrap.wrap_type.set(getattr(k, WRAPS_MAC[spec["wrap"]]))
    if "wrap_side" in spec:
        wrap.wrap_side.set(getattr(k, "wrap_" + spec["wrap_side"]))
    if "wrap_distance" in spec:
        for side in ("left", "right", "top", "bottom"):
            getattr(wrap, f"distance_{side}").set(spec["wrap_distance"])
    if "horizontal_from" in spec:
        s.relative_horizontal_position.set(
            getattr(k, _H_PREFIX + _MAC_HORIZONTAL[spec["horizontal_from"]]))
    if "vertical_from" in spec:
        s.relative_vertical_position.set(
            getattr(k, _V_PREFIX + _MAC_VERTICAL[spec["vertical_from"]]))
    if "width" in spec or "height" in spec:
        s.lock_aspect_ratio.set(False)
        if "width" in spec:
            s.width.set(spec["width"])
        if "height" in spec:
            s.height.set(spec["height"])
    if "left" in spec:
        s.left_position.set(NAMED_LEFT.get(spec["left"], spec["left"]))
    if "dleft" in spec:
        s.left_position.set(_moved(position_name(s.left_position.get(), NAMED_LEFT),
                                   spec["dleft"], "dleft"))
    if "top" in spec:
        s.top.set(NAMED_TOP.get(spec["top"], spec["top"]))
    if "dtop" in spec:
        s.top.set(_moved(position_name(s.top.get(), NAMED_TOP), spec["dtop"], "dtop"))


def mac_crop_picture(target, spec: dict) -> None:
    """Crop an inline picture or a floating one."""
    fmt = target.picture_format
    apply_crop(spec, lambda side: getattr(fmt, side).get(),
               lambda side, value: getattr(fmt, side).set(value),
               lambda axis: getattr(target, axis).get())


def _mac_apply(s, spec: dict, kind: str) -> list[str]:
    from appscript import k

    _check_text_allowed(kind, spec)
    if any(side in spec for side in CROPS):  # first: if it is refused, nothing else has changed
        if kind not in ("picture", "linked_picture"):
            raise ValueError(NOT_A_PICTURE.format(kind=kind))
        mac_crop_picture(s, spec)
    warnings = []
    if "kind" in spec:
        s.auto_shape_type.set(_mac_kind(spec["kind"]))
    _mac_place(s, spec)
    if "rotation" in spec:
        s.rotation_.set(spec["rotation"])
    if "corner_radius" in spec:
        if int(s.count(each=k.adjustment)) and _mac_kind_of(s) == "rounded_rectangle":
            s.adjustments[1].adjustment_value.set(spec["corner_radius"] * 0.5)
        else:
            warnings.append("corner_radius is for a rounded_rectangle, so it was not set.")
    _mac_apply_look(s, spec)
    warnings += _mac_apply_text(s, spec)
    if "alt_text" in spec:
        warnings.append(NO_ALT_ON_MAC)
    if "z_order" in spec:
        s.z_order(z_order_command=getattr(k, Z_ORDERS_MAC[spec["z_order"]]))
    if "name" in spec:
        s.name.set(spec["name"])
    return warnings


def _mac_result(doc, name: str, spec: dict, warnings: list[str]) -> dict:
    from . import mac as m

    s = doc.shapes[name]
    at = s.anchor.start_of_content.get()
    m._scroll(doc, at, at + 1)
    kind = _mac_kind_of(s)
    # `has text` is true for an empty shape on the Mac, and `overflowing` too.
    overflowing = (kind not in NO_TEXT and s.text_frame.overflowing.get() is True
                   and bool(m._clean(s.text_frame.text_range.content.get())))
    return _finished(_mac_describe(doc, s, _mac_number_of(doc, name)), spec, _mac_page(doc),
                     warnings, overflowing)


def _mac_finish_new(doc, name: str, spec: dict, kind: str) -> dict:
    """Give a new shape its place and look. If that fails no half-made shape stays."""
    try:
        warnings = _mac_apply(doc.shapes[name], spec, kind)
    except Exception:
        doc.shapes[name].delete()
        raise
    return _mac_result(doc, spec.get("name", name), spec, warnings)


@mac_impl(_insert_shape_impl)
def _mac_insert_shape(kind, index, width, height, spec):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    m._para(doc, index)
    _mac_check_name(doc, spec.get("name"))
    name = _mac_made(doc, _mac_names(doc), lambda: doc.make(
        new=k.shape, at=doc, with_properties={
            k.auto_shape_type: _mac_kind(kind), k.left_position: 0, k.top: 0, k.width: width,
            k.height: height, k.anchor: doc.paragraphs[index].text_object}))
    return _mac_finish_new(doc, name, _new_spec(spec, "square"), kind)


@mac_impl(_insert_text_box_impl)
def _mac_insert_text_box(index, width, height, spec):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    m._para(doc, index)
    _mac_check_name(doc, spec.get("name"))
    name = _mac_made(doc, _mac_names(doc), lambda: doc.make(
        new=k.text_box, at=doc, with_properties={
            k.left_position: 0, k.top: 0, k.width: width, k.height: height,
            k.anchor: doc.paragraphs[index].text_object}))
    return _mac_finish_new(doc, name, _new_spec(spec, "square"), "text_box")


@mac_impl(_insert_line_impl)
def _mac_insert_line(index, begin_x, begin_y, end_x, end_y, spec):
    from appscript import k

    from . import mac as m

    doc = backend.word.document()
    m._para(doc, index)
    _mac_check_name(doc, spec.get("name"))
    name = _mac_made(doc, _mac_names(doc), lambda: doc.make(
        new=k.line_shape, at=doc, with_properties={
            k.begin_line_X: begin_x, k.begin_line_Y: begin_y, k.end_line_X: end_x,
            k.end_line_Y: end_y, k.anchor: doc.paragraphs[index].text_object}))
    full = {**_new_spec(spec, "front"), **line_endpoints(begin_x, begin_y, end_x, end_y)}
    return _mac_finish_new(doc, name, full, "line")


def mac_float_picture(doc, number: int, spec: dict) -> dict:
    """An inline picture out of the text, at the place `spec` gives."""
    name = _mac_made(doc, _mac_names(doc), doc.inline_shapes[number].convert_to_shape)
    return {**_mac_finish_new(doc, name, _picture_spec(spec), "picture"), "now": "floating",
            "note": "It floats now: word_list_shapes lists it, and word_set_shape changes it."}


@mac_impl(_insert_floating_picture_impl)
def _mac_insert_floating_picture(path, index, width, height, spec):
    from appscript import k

    from . import mac as m
    from .references import _mac_apply_picture, _mac_stage

    if not os.path.isfile(path):
        raise ValueError(f"There is no file at '{path}'.")
    doc = backend.word.document()
    m._para(doc, index)
    _mac_check_name(doc, spec.get("name"))
    start, _ = m._bounds(doc, index)
    count = int(doc.count(each=k.inline_shape))
    staged, copied = _mac_stage(path)
    try:
        doc.make(new=k.inline_picture, at=doc.create_range(start=start, end_=start),
                 with_properties={k.file_name: staged, k.link_to_file: False,
                                  k.save_with_document: True})
    finally:
        if copied:
            os.remove(staged)
    if int(doc.count(each=k.inline_shape)) != count + 1:
        raise RuntimeError("Word answered without error, but no picture was added.")
    number = next(n for n in range(1, count + 2)
                  if doc.inline_shapes[n].text_object.start_of_content.get() >= start)
    warnings = _mac_apply_picture(doc, doc.inline_shapes[number], width, height, None,
                                  spec.get("alt_text"))
    name = _mac_made(doc, _mac_names(doc), doc.inline_shapes[number].convert_to_shape)
    out = _mac_finish_new(doc, name, _picture_spec(spec), "picture")
    if warnings:
        out["warnings"] = [*warnings, *out.get("warnings", [])]
    return out


@mac_impl(_list_shapes_impl)
def _mac_list_shapes(start, limit):
    doc = backend.word.document()
    found = _mac_in_reading_order(doc)
    out = {"success": True, "count": len(found), "shapes": [
        _mac_describe(doc, doc.shapes[place], n)
        for n, (_, place, _name) in enumerate(found, 1) if start <= n < start + limit]}
    if start + limit <= len(found):
        out["more"] = (f"Shapes {start + limit} to {len(found)} follow; "
                       f"pass start={start + limit}.")
    return out


@mac_impl(_set_shape_impl)
def _mac_set_shape(shape, spec):
    doc = backend.word.document()
    s, own_name = _mac_find(doc, shape)
    name = s.name.get()
    if "name" in spec and spec["name"] != name:
        _mac_check_name(doc, spec["name"])
    if not own_name and set(spec) - {"name"}:
        raise ValueError(TWO_NAMES.format(count=_mac_names(doc).count(name), name=name))
    kind = _mac_kind_of(s)
    if kind in ("picture", "linked_picture") and spec.get("wrap") == "inline":
        return _mac_picture_in_line(doc, s, name, spec, kind)
    warnings = _mac_apply(s, spec, kind)
    return _mac_result(doc, spec.get("name", name), spec, warnings)


def _mac_picture_in_line(doc, s, name: str, spec: dict, kind: str):
    from appscript import k

    from .references import _mac_describe_picture

    rest = {key: v for key, v in spec.items() if key not in ("wrap", "wrap_side", "wrap_distance")}
    warnings = _mac_apply(s, rest, kind)
    name = rest.get("name", name)
    at = doc.shapes[name].anchor.start_of_content.get()
    doc.shapes[name].convert_to_inline_shape()
    count = int(doc.count(each=k.inline_shape))
    number = next(n for n in range(1, count + 1)
                  if doc.inline_shapes[n].text_object.start_of_content.get() >= at)
    out = {"success": True, "now": "inline", **_mac_describe_picture(doc, doc.inline_shapes[number],
                                                                     number),
           "note": "It is an inline picture now: word_list_pictures lists it, and "
                   "word_set_picture changes it."}
    if warnings:
        out["warnings"] = warnings
    return out


@mac_impl(_delete_shape_impl)
def _mac_delete_shape(shape):
    from appscript import k

    doc = backend.word.document()
    s, _ = _mac_find(doc, shape)
    name = s.name.get()
    s.delete()
    return {"success": True, "deleted": name, "left": int(doc.count(each=k.shape))}


def mac_leading_anchors(doc, start: int, end: int, hidden: int) -> int:
    """leading_anchors, on Word for Mac."""
    from appscript import k

    if not hidden or not int(doc.count(each=k.shape)):
        return 0
    at = {a for a, _, _ in _mac_in_reading_order(doc) if start <= a < end}
    count = 0
    while start + count in at and count < hidden:
        count += 1
    return count


# -- the tools -----------------------------------------------------------------------

def register(mcp) -> None:
    @mcp.tool(name="word_insert_shape", annotations=annotations("Insert a shape", WRITES))
    async def word_insert_shape(params: InsertShapeInput) -> str:
        """Draw a shape (rectangle, oval, arrow, star, callout, flowchart symbol...)
        anchored to a paragraph, optionally with text inside.

        It floats on the page. By default text flows beside it (wrap square) and it
        sits at the start of its paragraph, against the left margin; move it with
        left, top, horizontal_from and vertical_from. Use wrap behind for behind the
        text (背面) and front for over it. word_get_page_image shows the result.

        Text in a drawn shape starts white, for the theme's dark fill: on a pale
        fill, give font color too. word_undo does not take back a new shape on
        the Mac: use word_delete_shape.
        """
        p = params
        return await call(_insert_shape_impl, p.kind, p.index, p.width, p.height,
                          spec_of(p, "kind", "index", "width", "height"), undo="Insert shape")

    @mcp.tool(name="word_insert_text_box", annotations=annotations("Insert a text box", WRITES))
    async def word_insert_text_box(params: InsertTextBoxInput) -> str:
        """Add a text box anchored to a paragraph: text in a box of its own that
        floats on the page. orientation vertical makes it East Asian vertical text
        (縦書き).

        It takes the same placing as word_insert_shape. Its text is not in
        word_get_paragraphs or the outline: word_list_shapes reads it.
        """
        p = params
        return await call(_insert_text_box_impl, p.index, p.width, p.height,
                          spec_of(p, "index", "width", "height"), undo="Insert text box")

    @mcp.tool(name="word_insert_line", annotations=annotations("Insert a line", WRITES))
    async def word_insert_line(params: InsertLineInput) -> str:
        """Draw a straight line or an arrow from one point to another, anchored to a
        paragraph. begin and end are in points from horizontal_from and vertical_from
        (by default the left margin and the top of the paragraph). It goes over the
        text unless wrap says otherwise."""
        p = params
        return await call(_insert_line_impl, p.index, p.begin_x, p.begin_y, p.end_x, p.end_y,
                          spec_of(p, "index", "begin_x", "begin_y", "end_x", "end_y"),
                          undo="Insert line")

    @mcp.tool(name="word_insert_floating_picture",
              annotations=annotations("Insert a floating picture", WRITES))
    async def word_insert_floating_picture(params: InsertFloatingPictureInput) -> str:
        """Put a picture from a file on the page, anchored to a paragraph: text flows
        round it (wrap square), or it sits behind the text (wrap behind, 背面) or over
        it (wrap front). It takes the same placing as word_insert_shape. It goes in the
        paragraph you name and adds none; word_insert_picture is for a picture in
        line with the text. Answers its name, for word_set_shape."""
        p = params
        return await call(_insert_floating_picture_impl, p.path, p.index, p.width, p.height,
                          spec_of(p, "path", "index", "width", "height"),
                          undo="Insert floating picture")

    @mcp.tool(name="word_list_shapes", annotations=annotations("List shapes", READ_ONLY))
    async def word_list_shapes(params: ListShapesInput) -> str:
        """Everything floating in the document's body, in reading order: name, kind,
        anchor paragraph, page, position and size, wrap, look, and the text inside,
        which word_get_paragraphs does not show. Pictures in line with the text are
        in word_list_pictures instead. Shapes in headers and footers are not listed."""
        p = params
        return await call(_list_shapes_impl, p.start, p.limit)

    @mcp.tool(name="word_set_shape", annotations=annotations("Change a shape", WRITES))
    async def word_set_shape(params: SetShapeInput) -> str:
        """Change a floating shape, text box, line or picture: move it (left, top,
        or by dleft, dtop), resize it, change its wrap (behind, front, square,
        inline...), its order among shapes, rotation, fill, outline, text, font,
        kind or name; crop a picture. Only what is given changes, and it is read back.

        A picture set to wrap inline goes into the text as an inline picture. A
        shape stays anchored to its paragraph: to take it to another paragraph,
        delete it and insert it there."""
        p = params
        return await call(_set_shape_impl, p.shape, spec_of(p, "shape"), undo="Change shape")

    @mcp.tool(name="word_delete_shape",
              annotations=annotations("Delete a shape", DESTRUCTIVE))
    async def word_delete_shape(params: DeleteShapeInput) -> str:
        """Delete a floating shape, text box, line or picture. Its paragraph stays."""
        p = params
        return await call(_delete_shape_impl, p.shape, undo="Delete shape")
