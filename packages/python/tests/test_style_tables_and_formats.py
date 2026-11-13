"""Cell-style records, number formats, workbook formats, conditional-format fields and table-style defaults."""

from __future__ import annotations

import unittest

import formulon
from formulon import (
    CellStyle,
    CfDataBarInput,
    CfDataBarRuleInput,
    CfIconSetInput,
    CfIconSetRuleInput,
    CfValueObject,
    IconSetName,
    Workbook,
)


def _xf(num_fmt_id: int) -> formulon.CellXfInput:
    return formulon.CellXfInput(
        font_index=0,
        fill_index=0,
        border_index=0,
        num_fmt_id=num_fmt_id,
        horizontal=None,
        vertical=None,
        wrap_text=None,
    )


class _Base(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)


class StyleTableTests(_Base):
    def test_set_cell_style_xf(self) -> None:
        ooxml = self.wb.ooxml
        index = ooxml.add_cell_style_xf(_xf(2))
        ooxml.set_cell_style_xf(index, _xf(4))
        self.assertEqual(ooxml.get_cell_style_xf(index).num_fmt_id, 4)
        record = ooxml.get_cell_style_xf(index)
        ooxml.set_cell_style_xf(index, record)
        self.assertEqual(ooxml.get_cell_style_xf(index), record)
        with self.assertRaises(formulon.FormulonError) as caught:
            ooxml.set_cell_style_xf(index + 1000, _xf(4))
        self.assertEqual(caught.exception.code, formulon.StatusCode.INDEX_OUT_OF_RANGE)

    def test_list_number_formats(self) -> None:
        ooxml = self.wb.ooxml
        first = ooxml.add_number_format('0.000" kg"')
        second = ooxml.add_number_format('0.0" m"')
        listed = ooxml.list_number_formats()
        self.assertIsInstance(listed[0], formulon.NumberFormat)
        ids = [f.num_fmt_id for f in listed]
        self.assertEqual(ids, sorted(ids))
        by_id = {f.num_fmt_id: f.number_format for f in listed}
        self.assertEqual(by_id[first], '0.000" kg"')
        self.assertEqual(by_id[second], '0.0" m"')


class CellStyleTests(_Base):
    def test_cell_style_name_only_defaults(self) -> None:
        style = CellStyle(name="Mine")
        self.assertEqual((style.style_xf_index, style.outline_level, style.hidden), (0, 0, False))
        # Only the RowLevel_n / ColLevel_n built-ins (builtin_id 1 and 2) carry an outline level.
        self.wb.set_cell_style(CellStyle(name="RowLevel_3", builtin_id=1, outline_level=2))
        read = self.wb.get_cell_style("RowLevel_3")
        self.assertEqual(read.outline_level, 2)
        self.assertFalse(hasattr(read, "i_level"))
        self.wb.set_cell_style(style)
        self.assertEqual(self.wb.get_cell_style("Mine").style_xf_index, 0)


class WorkbookFormatTests(unittest.TestCase):
    def test_ordinals(self) -> None:
        fmt = formulon.WorkbookFormat
        self.assertEqual((fmt.XLSX, fmt.XLSB), (0, 1))
        self.assertFalse(hasattr(fmt, "UNKNOWN"))
        self.assertEqual([m.name for m in fmt], ["XLSX", "XLSB"])

    def test_save_default_and_binary_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_value("A1", 3)
            plain = wb.save()
            binary = wb.save(format=formulon.WorkbookFormat.XLSB)
        self.assertEqual(plain[:2], b"PK")
        for data in (plain, binary):
            with Workbook.load(data) as again:
                self.assertEqual(again.sheet(0).get_value("A1"), 3.0)


class CfTests(_Base):
    def test_match_icon_fields(self) -> None:
        for i, value in enumerate((1, 50, 100)):
            self.ws.set_value(i, 0, value)
        self.ws.add_cf_rule(
            CfIconSetRuleInput(
                ranges=["A1:A3"],
                icon_set=CfIconSetInput(0, [CfValueObject(1, "33"), CfValueObject(1, "67")]),
            )
        )
        cells = self.ws.evaluate_cf_range("A1:A3")
        icons = [m.icon for cell in cells for m in cell.matches if m.icon is not None]
        self.assertEqual(len(icons), 3)
        self.assertTrue(all(icon.icon_set == IconSetName.THREE_ARROWS for icon in icons))
        self.assertEqual(sorted(icon.icon_index for icon in icons), [0, 1, 2])
        self.assertFalse(hasattr(icons[0], "set_name"))
        self.assertFalse(hasattr(icons[0], "index"))

    def test_data_bar_min_max(self) -> None:
        bar = CfDataBarInput(
            min=CfValueObject(3),
            max=CfValueObject(4),
            fill=formulon.Color(formulon.ColorKind.RGB, argb=0xFF0000FF),
        )
        self.ws.add_cf_rule(CfDataBarRuleInput(ranges=["A1:A3"], data_bar=bar))
        (rule,) = self.ws.list_cf_rules()
        self.assertEqual((rule.data_bar.min, rule.data_bar.max), (CfValueObject(3), CfValueObject(4)))
        self.assertFalse(hasattr(rule.data_bar, "minimum"))
        self.assertFalse(hasattr(rule.data_bar, "maximum"))


class TableStyleDefaultsTests(_Base):
    def test_names(self) -> None:
        self.wb.update_table_style_defaults(table_style_name="TableStyleLight1", pivot_style_name="PivotStyleLight2")
        defaults = self.wb.get_table_style_defaults()
        self.assertEqual(
            (defaults.table_style_name, defaults.pivot_style_name), ("TableStyleLight1", "PivotStyleLight2")
        )
        self.assertFalse(hasattr(defaults, "table_style"))
        self.wb.update_table_style_defaults(pivot_style_name=None)
        defaults = self.wb.get_table_style_defaults()
        self.assertEqual((defaults.table_style_name, defaults.pivot_style_name), ("TableStyleLight1", None))


class FontPatchSchemeTests(_Base):
    """A patch naming a different face unlinks the font from the theme; size-only and same-name patches keep the link."""

    def _cell_font(self):
        return self.ws.get_cell_format("B2").font

    def test_a_new_name_without_scheme_unlinks_from_the_theme(self) -> None:
        wb = formulon.Workbook.create_for_profile("mac-365-ja_JP")
        ws = wb.sheet(0)
        self.assertEqual(ws.get_cell_format("B2").font.scheme, formulon.FontScheme.MINOR)
        ws.update_cell_format("B2", font={"size_pt": 14.0})
        self.assertEqual(ws.get_cell_format("B2").font.scheme, formulon.FontScheme.MINOR)
        ws.update_cell_format("B2", font={"name": ws.get_cell_format("B2").font.name})
        self.assertEqual(ws.get_cell_format("B2").font.scheme, formulon.FontScheme.MINOR)
        ws.update_cell_format("B2", font={"name": "Arial"})
        font = ws.get_cell_format("B2").font
        self.assertEqual((font.name, font.scheme), ("Arial", formulon.FontScheme.NONE))
        self.assertEqual(font.charset, 128)

    def test_a_stated_scheme_wins(self) -> None:
        wb = formulon.Workbook.create_for_profile("mac-365-ja_JP")
        ws = wb.sheet(0)
        ws.update_cell_format("B2", font={"name": "Arial", "scheme": formulon.FontScheme.MAJOR})
        self.assertEqual(ws.get_cell_format("B2").font.scheme, formulon.FontScheme.MAJOR)


class NormalFontInheritanceTests(_Base):
    def test_cleared_size_and_name_take_the_normal_font_and_cleared_colour_is_none(self) -> None:
        self.wb.set_default_font(formulon.Font(name="Arial", size_pt=14.0))
        self.assertEqual(self.wb.ooxml.get_font(0).size_pt, 14.0)
        self.ws.update_cell_format("B2", font={"name": "Courier New", "size_pt": 20.0, "color": "#FF0000"})
        self.assertEqual(self.ws.get_cell_format("B2").font.size_pt, 20.0)

        self.ws.update_cell_format("B2", font={"name": None, "size_pt": None, "color": None})
        font = self.ws.get_cell_format("B2").font
        self.assertEqual(font.size_pt, 14.0)
        self.assertEqual(font.name, "Arial")
        self.assertEqual(font.color.kind, formulon.ColorKind.NONE)

        added = self.wb.ooxml.add_font(formulon.Font(name="Courier New"))
        self.assertEqual(self.wb.ooxml.get_font(added).size_pt, 14.0)
        unnamed = self.wb.ooxml.add_font(formulon.Font(size_pt=9.0))
        self.assertEqual(self.wb.ooxml.get_font(unnamed).name, "Arial")

        self.wb.set_default_font(formulon.Font(name="Calibri"))
        self.assertEqual(self.wb.ooxml.get_font(0).size_pt, 14.0)


if __name__ == "__main__":
    unittest.main()
