"""Table and PivotTable style names are checked against the kinds a style applies to."""

from __future__ import annotations

import io
import unittest
import zipfile
from typing import Callable, Dict

import formulon
from formulon import (
    CellAddress,
    CellRect,
    FormulonError,
    InvalidArgumentError,
    PivotInput,
    PivotWorksheetSourceInput,
    StatusCode,
    TableInput,
    TableStyle,
    Workbook,
)


def _sales_book() -> Workbook:
    wb = Workbook.create_default()
    ws = wb.sheet(0)
    ws.set_text(0, 0, "Product")
    ws.set_text(0, 1, "Amount")
    ws.set_number(1, 1, 10.0)
    ws.set_number(2, 1, 20.0)
    ws.add_table(TableInput(range=CellRect(0, 0, 2, 1), name="Sales", columns=["Product", "Amount"]))
    return wb


def _add_pivot(wb: Workbook, name: str = "ByProduct") -> formulon.Pivot:
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
    cache.add_field("Product")
    cache.add_field("Amount")
    for product, amount in (("Widget", 10.0), ("Gadget", 20.0)):
        record = cache.add_record()
        cache.set_record_value(record, 0, product)
        cache.set_record_value(record, 1, amount)
    return wb.sheet(0).add_pivot(PivotInput(name=name, cache_id=cache.cache_id, anchor=CellAddress(9, 4)))


def _rewrite(data: bytes, edits: Dict[str, Callable[[str], str]]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for info in source.infolist():
            body = source.read(info.filename)
            if info.filename in edits:
                body = edits[info.filename](body.decode("utf-8")).encode("utf-8")
            target.writestr(info.filename, body)
    return out.getvalue()


def _flag_style(data: bytes, flag: str) -> bytes:
    return _rewrite(
        data, {"xl/styles.xml": lambda xml: xml.replace('<tableStyle name="Both"', f'<tableStyle name="Both" {flag}')}
    )


class TableStyleKindTests(unittest.TestCase):
    def assertPublicMessage(self, error: Exception) -> None:
        text = str(error)
        self.assertNotIn("None", text)
        self.assertNotIn("null", text)

    def test_add_table_refuses_an_excluded_or_slicer_style_and_creates_no_table(self) -> None:
        with _sales_book() as wb:
            ws = wb.sheet(0)
            wb.set_table_style(TableStyle(name="PivotOnly", table=False))
            for style_name in ("PivotOnly", "SlicerStyleLight1"):
                with self.assertRaises(InvalidArgumentError) as caught:
                    ws.add_table(
                        TableInput(range=CellRect(0, 3, 2, 4), name="Other", columns=["A", "B"], style_name=style_name)
                    )
                self.assertIn("Sheet.add_table", str(caught.exception))
                self.assertPublicMessage(caught.exception)
            self.assertEqual(len(ws.list_tables()), 1)

    def test_pivot_style_and_defaults_refuse_a_style_that_excludes_the_kind(self) -> None:
        with _sales_book() as wb:
            pivot = _add_pivot(wb)
            wb.set_table_style(TableStyle(name="PivotOnly", table=False))
            wb.set_table_style(TableStyle(name="TableOnly", pivot=False))
            with self.assertRaises(InvalidArgumentError) as caught:
                pivot.update_style_options(style_name="TableOnly")
            self.assertIn("Pivot.update_style_options", str(caught.exception))
            self.assertPublicMessage(caught.exception)
            with self.assertRaises(InvalidArgumentError):
                pivot.update_style_options(style_name="SlicerStyleDark2")
            with self.assertRaises(InvalidArgumentError) as defaults:
                wb.update_table_style_defaults(table_style_name="PivotOnly")
            self.assertIn("Workbook.update_table_style_defaults", str(defaults.exception))
            self.assertPublicMessage(defaults.exception)
            with self.assertRaises(InvalidArgumentError):
                wb.update_table_style_defaults(pivot_style_name="TableOnly")

    def test_a_loaded_default_that_excludes_its_kind_refuses_new_objects(self) -> None:
        with _sales_book() as wb:
            wb.set_table_style(TableStyle(name="Both"))
            wb.update_table_style_defaults(table_style_name="Both", pivot_style_name="Both")
            saved = wb.save()
        with Workbook.load(_flag_style(saved, 'table="0"')) as no_tables:
            with self.assertRaises(FormulonError) as caught:
                no_tables.sheet(0).add_table(TableInput(range=CellRect(0, 3, 2, 4), name="Other", columns=["A", "B"]))
            self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
            self.assertPublicMessage(caught.exception)
            self.assertEqual(len(no_tables.sheet(0).list_tables()), 1)
        with Workbook.load(_flag_style(saved, 'pivot="0"')) as no_pivots:
            with self.assertRaises(FormulonError) as caught:
                _add_pivot(no_pivots)
            self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
            self.assertPublicMessage(caught.exception)
            self.assertEqual(no_pivots.sheet(0).list_pivots(), [])

    def test_remove_table_style_repoints_every_user_to_its_own_default(self) -> None:
        with _sales_book() as wb:
            wb.set_table_style(TableStyle(name="Both"))
            wb.sheet(0).table(0).update_style_options(style_name="Both")
            _add_pivot(wb).update_style_options(style_name="Both")
            saved = wb.save()
        for flag in ('table="0"', 'pivot="0"'):
            with Workbook.load(_flag_style(saved, flag)) as loaded:
                self.assertEqual(loaded.sheet(0).table(0).get_style_options().style_name, "Both")
                self.assertEqual(loaded.sheet(0).pivot(0).get_style_options().style_name, "Both")
                loaded.remove_table_style("Both")
                self.assertEqual(loaded.sheet(0).table(0).get_style_options().style_name, "TableStyleMedium2")
                self.assertEqual(loaded.sheet(0).pivot(0).get_style_options().style_name, "PivotStyleLight16")

    def test_a_built_in_name_is_refused_for_a_custom_style_and_dropped_on_load(self) -> None:
        with _sales_book() as wb:
            for name in ("SlicerStyleLight1", "TableStyleMedium9", "PivotStyleDark3"):
                with self.assertRaises(InvalidArgumentError):
                    wb.set_table_style(TableStyle(name=name))
            wb.set_table_style(TableStyle(name="Mine"))
            wb.sheet(0).table(0).update_style_options(style_name="Mine")
            saved = wb.save()
        for builtin in ("SlicerStyleLight1", "TableStyleMedium9"):
            renamed = _rewrite(
                saved,
                {
                    "xl/styles.xml": lambda xml, b=builtin: xml.replace(
                        '<tableStyle name="Mine"', f'<tableStyle name="{b}"'
                    ),
                    "xl/tables/table1.xml": lambda xml, b=builtin: xml.replace('name="Mine"', f'name="{b}"'),
                },
            )
            with Workbook.load(renamed) as loaded:
                self.assertEqual(loaded.list_table_styles(), [])
                self.assertEqual(loaded.sheet(0).table(0).get_style_options().style_name, builtin)
                with Workbook.load(loaded.save()) as again:
                    self.assertEqual(again.sheet(0).table(0).get_style_options().style_name, builtin)


if __name__ == "__main__":
    unittest.main()
