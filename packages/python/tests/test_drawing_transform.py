"""Shape flips, picture rotation and flips, and rotation stored as given."""

from __future__ import annotations

import unittest

from formulon import (
    AnchorKind,
    DrawingAnchorInput,
    DrawingObjectKind,
    FormulonError,
    ImageProperties,
    InvalidArgumentError,
    PictureProperties,
    ShapeInput,
    StatusCode,
    Workbook,
    WorkbookFormat,
)

PNG = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
    + (1).to_bytes(4, "big")
    + (1).to_bytes(4, "big")
    + b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
)
BOX = DrawingAnchorInput(from_row=1, from_col=1, to_row=5, to_col=4)
FORMATS = (WorkbookFormat.XLSX, WorkbookFormat.XLSB)


class ShapeTransformTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_flips_read_false_until_set_and_a_patch_to_false_clears_them(self) -> None:
        plain = self.ws.add_shape(ShapeInput(anchor=BOX))
        self.assertEqual((plain.get_properties().flip_h, plain.get_properties().flip_v), (False, False))
        shape = self.ws.add_shape(ShapeInput(anchor=BOX, flip_h=True))
        self.assertEqual((shape.get_properties().flip_h, shape.get_properties().flip_v), (True, False))
        shape.update(flip_v=True, flip_h=False)
        self.assertEqual((shape.get_properties().flip_h, shape.get_properties().flip_v), (False, True))
        shape.update(flip_v=True)
        self.assertTrue(shape.get_properties().flip_v)

    def test_a_connector_takes_flips(self) -> None:
        line = self.ws.add_shape(
            ShapeInput(
                anchor=DrawingAnchorInput(kind=AnchorKind.ONE_CELL, from_row=1, width_emu=914400, height_emu=457200),
                preset_geometry="straightConnector1",
                flip_v=True,
            )
        )
        props = line.get_properties()
        self.assertEqual(props.kind, DrawingObjectKind.CONNECTOR)
        self.assertTrue(props.flip_v)

    def test_none_is_a_type_error_on_update_and_a_non_bool_is_refused_everywhere(self) -> None:
        shape = self.ws.add_shape(ShapeInput(anchor=BOX))
        for name in ("flip_h", "flip_v"):
            with self.assertRaises(TypeError):
                shape.update(**{name: None})
            with self.assertRaises(TypeError):
                shape.update(**{name: 1})
            with self.assertRaises(TypeError):
                self.ws.add_shape(ShapeInput(anchor=BOX, **{name: "yes"}))
        self.assertFalse(shape.get_properties().flip_h)

    def test_rotation_is_stored_as_given_and_survives_both_formats(self) -> None:
        shape = self.ws.add_shape(ShapeInput(anchor=BOX, rotation_deg=360.0, flip_h=True))
        self.assertEqual(shape.get_properties().rotation_deg, 360.0)
        shape.update(rotation_deg=-90.0)
        self.assertEqual(shape.get_properties().rotation_deg, -90.0)
        shape.update(rotation_deg=400.0)
        self.assertEqual(shape.get_properties().rotation_deg, 400.0)
        shape.update(rotation_deg=-30.0, flip_v=True)
        for fmt in FORMATS:
            with self.subTest(format=fmt):
                with Workbook.load(self.wb.save(format=fmt)) as loaded:
                    props = loaded.sheet(0).shape(shape.object_id).get_properties()
                    self.assertEqual((props.rotation_deg, props.flip_h, props.flip_v), (-30.0, True, True))
        shape.update(rotation_deg=0.0)
        self.assertEqual(shape.get_properties().rotation_deg, 0.0)

    def test_a_rotation_that_does_not_fit_is_refused(self) -> None:
        shape = self.ws.add_shape(ShapeInput(anchor=BOX, rotation_deg=10.0))
        for bad in (35792.0, -35792.0, 1e12):
            with self.assertRaises(InvalidArgumentError):
                shape.update(rotation_deg=bad)
            with self.assertRaises(InvalidArgumentError):
                self.ws.add_shape(ShapeInput(anchor=BOX, rotation_deg=bad))
        for bad in (float("nan"), float("inf")):
            with self.assertRaises((TypeError, ValueError, FormulonError)):
                shape.update(rotation_deg=bad)
        self.assertEqual(shape.get_properties().rotation_deg, 10.0)
        shape.update(rotation_deg=35791.0)
        self.assertEqual(shape.get_properties().rotation_deg, 35791.0)


class PictureTransformTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_a_new_picture_is_unrotated_and_keeps_the_image_fields(self) -> None:
        props = self.ws.add_picture(PNG).get_properties()
        self.assertIsInstance(props, PictureProperties)
        self.assertIsInstance(props, ImageProperties)
        self.assertEqual((props.rotation_deg, props.flip_h, props.flip_v), (0.0, False, False))
        self.assertEqual((props.width_px, props.height_px, props.bytes), (1, 1, PNG))

    def test_update_sets_the_transform_beside_the_common_keys_and_keeps_the_anchor(self) -> None:
        picture = self.ws.add_picture(PNG, None)
        before = picture.get_drawing_properties().anchor
        picture.update(rotation_deg=-90.0, flip_h=True)
        self.assertEqual(picture.get_drawing_properties().anchor, before)
        props = picture.get_properties()
        self.assertEqual((props.rotation_deg, props.flip_h, props.flip_v), (-90.0, True, False))
        picture.update(rotation_deg=360.0, flip_h=False, flip_v=True, name="Turned")
        props = picture.get_properties()
        self.assertEqual((props.rotation_deg, props.flip_h, props.flip_v), (360.0, False, True))
        self.assertEqual(picture.name, "Turned")

    def test_the_transform_survives_both_formats(self) -> None:
        picture = self.ws.add_picture(PNG)
        picture.update(rotation_deg=45.0, flip_h=True, flip_v=True)
        for fmt in FORMATS:
            with self.subTest(format=fmt):
                with Workbook.load(self.wb.save(format=fmt)) as loaded:
                    props = loaded.sheet(0).picture(picture.object_id).get_properties()
                    self.assertEqual((props.rotation_deg, props.flip_h, props.flip_v), (45.0, True, True))

    def test_none_and_wrong_types_are_a_type_error(self) -> None:
        picture = self.ws.add_picture(PNG)
        for kwargs in ({"rotation_deg": None}, {"flip_h": None}, {"flip_v": "yes"}, {"rotation_deg": "90"}):
            with self.assertRaises(TypeError):
                picture.update(**kwargs)
        self.assertEqual(picture.get_properties().rotation_deg, 0.0)

    def test_a_rotation_that_does_not_fit_changes_nothing(self) -> None:
        picture = self.ws.add_picture(PNG)
        picture.update(rotation_deg=20.0)
        for bad in (35792.0, -35792.0):
            with self.assertRaises(InvalidArgumentError) as caught:
                picture.update(rotation_deg=bad, name="Never")
            self.assertEqual(caught.exception.code, StatusCode.INVALID_ARGUMENT)
        self.assertEqual(picture.get_properties().rotation_deg, 20.0)
        self.assertNotEqual(picture.name, "Never")


if __name__ == "__main__":
    unittest.main()
