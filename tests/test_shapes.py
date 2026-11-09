import json

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import _common, references, shapes, text


def answer(impl, *args):
    return json.loads(_common.run(impl, *args))


def add(index=1, kind="rectangle", width=120.0, height=60.0, **spec):
    return answer(shapes._insert_shape_impl, kind, index, width, height, spec)




def test_a_new_shape_says_where_it_went(make_doc):
    make_doc("One.", "Two.")
    out = add(2, left=10.0, top=5.0, name="Box")
    assert (out["name"], out["kind"], out["index"], out["left"], out["top"]) == (
        "Box", "rectangle", 2, 10.0, 5.0)
    assert (out["horizontal_from"], out["vertical_from"], out["wrap"]) == (
        "margin", "paragraph", "square")
    assert "warnings" not in out


def test_named_positions_are_answered_by_name(make_doc):
    doc = make_doc("One.")
    out = add(left="center", top="bottom", horizontal_from="page", vertical_from="page")
    assert (out["left"], out["top"]) == ("center", "bottom")
    assert doc.Shapes.Item(1).Left == -999995


def test_behind_the_text_and_over_it(make_doc):
    make_doc("One.")
    assert add(wrap="behind")["wrap"] == "behind"
    assert add(wrap="front")["wrap"] == "front"


def test_a_name_already_used_is_refused_before_anything_is_made(make_doc):
    doc = make_doc("One.")
    add(name="Box")
    assert "already named 'Box'" in add(name="Box")["error"]
    assert doc.Shapes.Count == 1


def test_colors_outline_and_text(make_doc):
    doc = make_doc("One.")
    out = add(fill="#FF8040", line_color="none", text="上\n下", alignment="center")
    assert (out["fill"], out["line_color"], out["text"]) == ("#FF8040", "none", "上\n下")
    assert doc.Shapes.Item(1).TextFrame.TextRange.ParagraphFormat.Alignment == 1


def test_no_fill(make_doc):
    make_doc("One.")
    assert add(fill="none")["fill"] == "none"


def test_text_that_does_not_fit_warns_when_word_says_so():
    got = {"left": 0.0, "top": 0.0, "width": 10.0, "height": 10.0}
    page = {"page_width": 100, "page_height": 100, "left_margin": 0, "top_margin": 0}
    assert "does not fit" in shapes._finished(got, {}, page, [], True)["warnings"][0]
    assert "warnings" not in shapes._finished(got, {}, page, [], False)


def test_overflowing_is_not_asked_on_windows(make_doc):
    doc = make_doc("One.")
    name = add()["name"]
    doc.Shapes.Item(1).TextFrame.Overflowing = True
    assert "warnings" not in answer(shapes._set_shape_impl, name, {"text": "Short."})


def test_a_shape_that_cannot_be_finished_does_not_stay(make_doc, monkeypatch):
    doc = make_doc("One.")

    def fails(*args):
        raise RuntimeError("Word refused.")

    monkeypatch.setattr(shapes, "_apply", fails)
    assert "Word refused" in add()["error"]
    assert doc.Shapes.Count == 0


def test_off_the_page_warns(make_doc):
    make_doc("One.")
    out = add(left=590.0, horizontal_from="page", width=100.0)
    assert "runs off the page" in out["warnings"][0]


def test_a_text_box_and_vertical_text(make_doc):
    doc = make_doc("One.")
    out = answer(shapes._insert_text_box_impl, 1, 60.0, 150.0,
                 {"text": "縦書き", "orientation": "vertical"})
    assert (out["kind"], out["text"], out["orientation"]) == ("text_box", "縦書き", "vertical")
    assert doc.Shapes.Item(1).TextFrame.Orientation == 4


def test_a_line_is_placed_by_its_ends_whichever_way_it_runs(make_doc):
    make_doc("One.")
    out = answer(shapes._insert_line_impl, 1, 300.0, 400.0, 50.0, 350.0,
                 {"end_arrow": "triangle", "line_color": "#112233"})
    assert (out["left"], out["top"], out["width"], out["height"]) == (50.0, 350.0, 250.0, 50.0)
    assert (out["kind"], out["wrap"], out["end_arrow"]) == ("line", "front", "triangle")
    assert "text" not in out


def test_a_line_callout_keeps_its_kind(make_doc):
    doc = make_doc("One.")
    out = add(kind="line_callout_one")
    assert out["kind"] == "line_callout_one"
    doc.Shapes.Item(1).Type = 2
    assert answer(shapes._list_shapes_impl, 1, 50)["shapes"][0]["kind"] == "line_callout_one"


def test_a_line_holds_no_text(make_doc):
    make_doc("One.")
    name = answer(shapes._insert_line_impl, 1, 0.0, 0.0, 10.0, 10.0, {})["name"]
    assert "A line holds no text" in answer(shapes._set_shape_impl, name, {"text": "x"})["error"]


def test_shapes_are_listed_in_reading_order_not_z_order(make_doc):
    make_doc("One.", "Two.", "Three.")
    add(3, name="Last")
    add(1, name="First")
    out = answer(shapes._list_shapes_impl, 1, 50)
    assert [s["name"] for s in out["shapes"]] == ["First", "Last"]
    assert [s["number"] for s in out["shapes"]] == [1, 2]


def test_a_long_list_says_where_to_go_on(make_doc):
    make_doc("One.")
    for _ in range(3):
        add()
    out = answer(shapes._list_shapes_impl, 1, 2)
    assert len(out["shapes"]) == 2 and out["more"] == "Shapes 3 to 3 follow; pass start=3."


def test_a_shape_by_name_or_number(make_doc):
    make_doc("One.")
    add(name="Box")
    assert answer(shapes._set_shape_impl, "Box", {"rotation": 30.0})["rotation"] == 30.0
    assert answer(shapes._set_shape_impl, 1, {"rotation": 45.0})["rotation"] == 45.0
    assert "no shape named 'Nope'" in answer(shapes._set_shape_impl, "Nope", {"rotation": 1.0})["error"]
    assert "no shape 2" in answer(shapes._set_shape_impl, 2, {"rotation": 1.0})["error"]


def test_two_shapes_with_one_name_are_not_guessed_between(make_doc):
    doc = make_doc("One.")
    add(name="Box")
    add(name="Other")
    doc.Shapes.Item(2).Name = "Box"
    assert "2 shapes are named 'Box'" in answer(shapes._set_shape_impl, "Box", {"rotation": 1.0})["error"]


def test_renaming_to_a_name_in_use_is_refused(make_doc):
    make_doc("One.")
    add(name="Box")
    add(name="Other")
    assert "already named" in answer(shapes._set_shape_impl, "Other", {"name": "Box"})["error"]
    assert answer(shapes._set_shape_impl, "Other", {"name": "Other"})["name"] == "Other"


def test_moving_it_from_where_it_is(make_doc):
    make_doc("One.")
    add(left=10.0, top=20.0, name="Box")
    out = answer(shapes._set_shape_impl, "Box", {"dleft": 5.0, "dtop": -3.0})
    assert (out["left"], out["top"]) == (15.0, 17.0)


def test_a_shape_at_a_named_place_cannot_be_moved_by_points(make_doc):
    make_doc("One.")
    add(left="center", name="Box")
    assert "It is at center" in answer(shapes._set_shape_impl, "Box", {"dleft": 5.0})["error"]


def test_order_among_shapes(make_doc):
    make_doc("One.")
    add(name="A")
    add(name="B")
    assert answer(shapes._set_shape_impl, "A", {"z_order": "front"})["z"] == 2
    assert answer(shapes._set_shape_impl, "A", {"z_order": "back"})["z"] == 1


def test_changing_the_kind_and_corners(make_doc):
    doc = make_doc("One.")
    add(name="A")
    assert answer(shapes._set_shape_impl, "A", {"kind": "rounded_rectangle",
                                                "corner_radius": 1.0})["kind"] == "rounded_rectangle"
    assert doc.Shapes.Item(1).Adjustments[1] == 0.5


def test_deleting_a_shape(make_doc):
    make_doc("One.")
    add(name="A")
    assert answer(shapes._delete_shape_impl, "A") == {"success": True, "deleted": "A", "left": 0}


def test_rewriting_a_paragraph_says_when_a_shape_went_with_it(make_doc):
    doc = make_doc("One.", "Two.")
    add(2, name="A")
    doc.Shapes.Item(1).Anchor.Start += 1  # inside the text, so it goes with it
    out = answer(text._set_paragraph_text_impl, 2, "New.")
    assert "1 floating shape(s)" in out["warnings"][0]


def test_deleting_a_paragraph_says_when_a_shape_went_with_it(make_doc):
    make_doc("One.", "Two.", "Three.")
    add(2, name="A")
    out = answer(text._delete_paragraphs_impl, 2, None)
    assert "1 floating shape(s)" in out["warnings"][0]


def test_a_paragraph_without_shapes_says_nothing_of_them(make_doc):
    make_doc("One.", "Two.")
    assert "warnings" not in answer(text._set_paragraph_text_impl, 1, "New.")
    assert "warnings" not in answer(text._delete_paragraphs_impl, 2, None)


def test_kinds_are_taken_loosely_and_refused_with_the_nearest():
    assert shapes.kind_name("Rounded Rectangle") == "rounded_rectangle"
    assert shapes.kind_name("right-arrow") == "right_arrow"
    assert shapes.kind_name("triangle") == "isosceles_triangle"
    assert shapes.kind_name("pentagon") == "regular_pentagon"
    assert shapes.kind_name("home_plate") == "home_plate"
    assert shapes.kind_name("star") == "five_point_star"
    assert shapes.kind_name("balloon") == "rounded_rectangular_callout"
    with pytest.raises(ValueError, match="Did you mean rounded_rectangle"):
        shapes.kind_name("rounded_rectangl")


def test_the_table_is_msos_numbers():
    assert shapes.KINDS["rectangle"] == 1
    assert shapes.KINDS["right_arrow"] == 33
    assert shapes.KINDS["home_plate"] == 51
    assert "balloon" not in shapes.KINDS
    assert shapes.KINDS["flowchart_offline_storage"] == 139
    assert shapes.KINDS["corner"] == 162
    assert shapes.KINDS["cloud"] == 179
    assert shapes.KINDS["line_inverse"] == 183
    assert "not_primitive" not in shapes.KINDS


def test_a_shape_in_line_has_no_position():
    with pytest.raises(ValidationError, match="in line with the text"):
        shapes.InsertShapeInput(kind="oval", index=1, width=1, height=1, wrap="inline", left=3)


def test_a_line_needs_two_places():
    with pytest.raises(ValidationError, match="different places"):
        shapes.InsertLineInput(index=1, begin_x=1, begin_y=1, end_x=1, end_y=1)
    with pytest.raises(ValidationError, match="leave out left and top"):
        shapes.InsertLineInput(index=1, begin_x=1, begin_y=1, end_x=2, end_y=2, left=3)


def test_a_change_changes_something():
    with pytest.raises(ValidationError, match="at least one"):
        shapes.SetShapeInput(shape="A")
    with pytest.raises(ValidationError, match="left or dleft"):
        shapes.SetShapeInput(shape="A", left=1, dleft=2)
    assert shapes.SetShapeInput(shape=2, wrap="behind").shape == 2
    with pytest.raises(ValidationError, match="numbered from 1"):
        shapes.SetShapeInput(shape=0, wrap="behind")


def test_corners_are_for_a_rounded_rectangle():
    with pytest.raises(ValidationError, match="rounded_rectangle"):
        shapes.InsertShapeInput(kind="oval", index=1, width=1, height=1, corner_radius=0.5)


@pytest.fixture
def picture(tmp_path):
    path = tmp_path / "a.png"
    path.write_bytes(b"not read by the fake")
    return str(path)


def floating(path, index=1, width=None, height=None, **spec):
    return answer(shapes._insert_floating_picture_impl, path, index, width, height, spec)


def test_a_floating_picture_goes_in_the_paragraph_named(make_doc, picture):
    doc = make_doc("One.", "Two.")
    out = floating(picture, 2, width=180.0, name="Pic", wrap="behind", left=10.0, alt_text="Box")
    assert (out["kind"], out["name"], out["index"], out["wrap"], out["left"], out["width"],
            out["height"], out["alt_text"]) == (
        "picture", "Pic", 2, "behind", 10.0, 180.0, 90.0, "Box")
    assert doc.Paragraphs.Count == 2 and doc.InlineShapes.Count == 0


def test_a_floating_picture_needs_its_file(make_doc):
    make_doc("One.")
    assert "There is no file" in floating("/no/such.png")["error"]


def test_a_picture_in_the_text_comes_out_of_it(make_doc, picture):
    doc = make_doc("One.")
    answer(references._insert_picture_impl, picture, None, None, None, None, None, "Box")
    out = answer(references._set_picture_impl, 1, None, None, None, None,
                 {"wrap": "behind", "left": 20.0, "top": 30.0, "horizontal_from": "page",
                  "vertical_from": "page", "name": "Moved"})
    assert (out["now"], out["kind"], out["name"], out["wrap"], out["left"], out["alt_text"]) == (
        "floating", "picture", "Moved", "behind", 20.0, "Box")
    assert doc.InlineShapes.Count == 0 and doc.Shapes.Count == 1


def test_a_floating_picture_goes_back_into_the_text(make_doc, picture):
    doc = make_doc("One.")
    floating(picture, name="Pic", alt_text="Box")
    out = answer(shapes._set_shape_impl, "Pic", {"wrap": "inline"})
    assert (out["now"], out["picture"], out["alt_text"]) == ("inline", 1, "Box")
    assert doc.Shapes.Count == 0 and doc.InlineShapes.Count == 1


def test_a_crop_counts_points_as_the_picture_is_shown(make_doc, picture):
    doc = make_doc("One.")
    answer(references._insert_picture_impl, picture, None, None, None, None, None, None)
    out = answer(references._set_picture_impl, 1, None, None, None, None, {"crop_left": 10.0})
    assert out["width"] == 80.0
    assert doc.InlineShapes.Item(1).PictureFormat.CropLeft == 5.0  # shown at twice its size
    assert answer(references._set_picture_impl, 1, None, None, None, None,
                  {"crop_left": 0.0})["width"] == 90.0


def test_a_floating_picture_is_cropped_and_says_so(make_doc, picture):
    make_doc("One.")
    name = floating(picture)["name"]
    out = answer(shapes._set_shape_impl, name, {"crop_top": 6.0})
    assert out["height"] == 39.0 and out["cropped"] is True


def test_crops_that_leave_nothing_are_refused(make_doc, picture):
    doc = make_doc("One.")
    name = floating(picture)["name"]
    out = answer(shapes._set_shape_impl, name, {"crop_left": 50.0, "crop_right": 50.0})
    assert "leave nothing" in out["error"]
    assert doc.Shapes.Item(1).Width == 90.0


def test_one_crop_too_many_leaves_the_other_undone(make_doc, picture):
    doc = make_doc("One.")
    name = floating(picture)["name"]
    out = answer(shapes._set_shape_impl, name, {"crop_left": 10.0, "crop_top": 100.0})
    assert "leave nothing" in out["error"]
    assert (doc.Shapes.Item(1).Width, doc.Shapes.Item(1).Height) == (90.0, 45.0)


def test_a_refused_crop_changes_nothing_else(make_doc, picture):
    doc = make_doc("One.")
    name = floating(picture, left=10.0, name="Pic")["name"]
    out = answer(shapes._set_shape_impl, name, {"crop_left": 500.0, "left": 40.0, "rotation": 9.0})
    assert "leave nothing" in out["error"]
    assert (doc.Shapes.Item(1).Left, doc.Shapes.Item(1).Rotation) == (10.0, 0.0)
    answer(references._insert_picture_impl, picture, None, None, None, None, None, None)
    out = answer(references._set_picture_impl, 1, 60.0, None, None, "Alt", {"crop_left": 500.0})
    assert "leave nothing" in out["error"]
    assert (doc.InlineShapes.Item(1).Width, doc.InlineShapes.Item(1).AlternativeText) == (90.0, "")


def test_a_picture_has_no_fill_to_change(make_doc, picture):
    make_doc("One.")
    name = floating(picture)["name"]
    assert "no fill to change" in answer(shapes._set_shape_impl, name,
                                         {"fill_transparency": 0.8})["error"]


def test_a_line_can_be_made_level(make_doc):
    assert shapes.SetShapeInput(shape="A", height=0).height == 0


def test_a_picture_is_faded_in_the_text_and_keeps_it_when_it_floats(make_doc, picture):
    doc = make_doc("One.")
    answer(references._insert_picture_impl, picture, None, None, None, None, None, None)
    out = answer(references._set_picture_impl, 1, None, None, None, None,
                 {"brightness": 0.85, "contrast": 0.15})
    assert (out["brightness"], out["contrast"]) == (0.85, 0.15)
    format_ = doc.InlineShapes.Item(1).PictureFormat
    assert (format_.Brightness, format_.Contrast) == (0.85, 0.15)


def test_a_floating_picture_is_faded(make_doc, picture):
    make_doc("One.")
    out = floating(picture, wrap="behind", brightness=0.85, contrast=0.15, name="Pic")
    assert (out["brightness"], out["contrast"]) == (0.85, 0.15)
    out = answer(shapes._set_shape_impl, "Pic", {"brightness": 0.5, "contrast": 0.5})
    assert "brightness" not in out


def test_only_a_picture_is_faded(make_doc):
    make_doc("One.")
    add(name="Box")
    assert "Only a picture" in answer(shapes._set_shape_impl, "Box", {"brightness": 0.8})["error"]


def test_the_text_in_a_shape_takes_paragraph_settings(make_doc):
    doc = make_doc("One.")
    out = answer(shapes._insert_text_box_impl, 1, 100.0, 60.0,
                 {"text": "A\nB", "paragraph": {"snap_to_grid": False, "space_after": 0.0}})
    assert "error" not in out
    assert doc.Shapes.Item(1).TextFrame.TextRange.ParagraphFormat.DisableLineHeightGrid is True


def test_a_line_takes_no_paragraph_settings(make_doc):
    make_doc("One.")
    name = answer(shapes._insert_line_impl, 1, 0.0, 0.0, 10.0, 10.0, {})["name"]
    out = answer(shapes._set_shape_impl, name, {"paragraph": {"space_after": 0.0}})
    assert "holds no text" in out["error"]


def test_borders_are_not_for_the_text_in_a_shape():
    with pytest.raises(ValidationError, match="not set on the text in a shape"):
        shapes.SetShapeInput(shape="A", paragraph={"shading": "#FFFFFF"})


def test_only_a_picture_is_cropped(make_doc):
    make_doc("One.")
    add(name="Box")
    assert "Only a picture" in answer(shapes._set_shape_impl, "Box", {"crop_left": 1.0})["error"]


def test_a_picture_that_is_not_in_the_text_is_pointed_at(make_doc):
    make_doc("One.")
    error = answer(references._set_picture_impl, 1, 10.0, None, None, None, {})["error"]
    assert "word_list_shapes" in error


def test_a_picture_is_placed_only_when_it_floats():
    assert references.SetPictureInput(picture=1, crop_left=1).crop_left == 1
    assert references.SetPictureInput(picture=1, wrap="behind", left=3).left == 3
    with pytest.raises(ValidationError, match="Give `wrap` too"):
        references.SetPictureInput(picture=1, left=3)
    with pytest.raises(ValidationError, match="in line with the text already"):
        references.SetPictureInput(picture=1, wrap="inline")
    with pytest.raises(ValidationError, match="at least one"):
        references.SetPictureInput(picture=1)


def test_a_floating_picture_takes_a_full_path_and_floats():
    with pytest.raises(ValidationError, match="full path"):
        shapes.InsertFloatingPictureInput(path="a.png", index=1)
    with pytest.raises(ValidationError, match="word_insert_picture"):
        shapes.InsertFloatingPictureInput(path="/a.png", index=1, wrap="inline")
