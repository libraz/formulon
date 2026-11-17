"""Connectors glued to a shape follow it when the shape's anchor is patched."""

from __future__ import annotations

import io
import re
import unittest
import zipfile

from formulon import AnchorKind, DrawingAnchorInput, ShapeInput, Workbook

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def anchor(col: int, width: int, height: int) -> DrawingAnchorInput:
    return DrawingAnchorInput(kind=AnchorKind.ONE_CELL, from_row=1, from_col=col, width_emu=width, height_emu=height)


def glue(data: bytes, connector: int, start: tuple[int, int] | None, end: tuple[int, int] | None) -> bytes:
    """``data`` with connector ``connector`` glued at ``start`` / ``end`` (shape id, site index)."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            body = src.read(item.filename)
            if re.fullmatch(r"xl/drawings/drawing\d+\.xml", item.filename):
                text = body.decode()
                ends = "".join(
                    f'<a:{tag} xmlns:a="{NS_A}" id="{at[0]}" idx="{at[1]}"/>'
                    for tag, at in (("stCxn", start), ("endCxn", end))
                    if at is not None
                )
                pattern = rf'(<xdr:cNvPr id="{connector}"[^>]*/>\s*)<xdr:cNvCxnSpPr/>'
                text, count = re.subn(pattern, rf"\1<xdr:cNvCxnSpPr>{ends}</xdr:cNvCxnSpPr>", text)
                assert count == 1, text
                body = text.encode()
            dst.writestr(item, body)
    return out.getvalue()


class ConnectorGlueTests(unittest.TestCase):
    def test_patching_a_glued_shape_moves_its_connectors(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            start = ws.add_shape(ShapeInput(preset_geometry="rect", anchor=anchor(1, 914400, 457200)))
            end = ws.add_shape(ShapeInput(preset_geometry="rect", anchor=anchor(6, 914400, 457200)))
            both = ws.add_shape(ShapeInput(preset_geometry="straightConnector1", anchor=anchor(2, 914400, 1)))
            one = ws.add_shape(ShapeInput(preset_geometry="straightConnector1", anchor=anchor(3, 914400, 1)))
            data = glue(wb.save(), both.object_id, (start.object_id, 3), (end.object_id, 1))
            data = glue(data, one.object_id, None, (end.object_id, 1))
        with Workbook.load(data) as wb:
            ws = wb.sheet(0)
            target = ws.shape(end.object_id)
            target.update(anchor=DrawingAnchorInput(from_col_offset_emu=254000))
            routed = ws.drawing(both.object_id).get_drawing_properties().anchor
            moved = ws.drawing(one.object_id).get_drawing_properties().anchor
            target.update(anchor=DrawingAnchorInput(from_col_offset_emu=381000))
            # Glued at both ends, the connector runs between the two sites; glued at one, it moves whole.
            rerouted = ws.drawing(both.object_id).get_drawing_properties().anchor
            self.assertEqual(rerouted.width_emu, routed.width_emu + 127000)
            self.assertEqual(rerouted.height_emu, 0)
            self.assertEqual(rerouted.from_col_offset_emu, routed.from_col_offset_emu)
            translated = ws.drawing(one.object_id).get_drawing_properties().anchor
            self.assertEqual(translated.width_emu, moved.width_emu)
            self.assertEqual(translated.from_col_offset_emu, moved.from_col_offset_emu + 127000)
            props = ws.shape(both.object_id).get_properties()
            self.assertEqual(props.start_connection.object_id, start.object_id)
            self.assertEqual(props.end_connection.site_index, 1)


if __name__ == "__main__":
    unittest.main()
