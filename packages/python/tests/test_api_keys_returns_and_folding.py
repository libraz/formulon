"""Natural keys, returned keys, table totals, merge refusal, name folding and the I/O error family."""

import re
import unittest

import formulon
from formulon import (
    AlreadyExistsError,
    CellAddress,
    CellRect,
    CfExpressionRuleInput,
    DataTableInput,
    DefinedNameInput,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    InvalidFileError,
    IoError,
    NotFoundError,
    Person,
    ProtectedRangeInput,
    StatusCode,
    TableColumnPatch,
    TableInput,
    TableStyle,
    TableTotalsRowFunction,
    Workbook,
)

_GUID = re.compile(r"^\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$")
_ZERO_GUID = "{00000000-0000-0000-0000-000000000000}"
_PERSON = "{abcdef01-0000-0000-0000-000000000001}"


class IoErrorFamilyTests(unittest.TestCase):
    def test_an_unreadable_image_raises_the_io_class(self) -> None:
        with self.assertRaises(IoError) as caught:
            formulon.get_image_info(b"not an image")
        self.assertEqual(caught.exception.code, StatusCode.IO_IMAGE_UNSUPPORTED)

    def test_an_io_range_code_without_its_own_class_constructs_the_io_class(self) -> None:
        error = FormulonError(StatusCode.IO_IMAGE_UNSUPPORTED)
        self.assertIs(type(error), IoError)
        self.assertIsInstance(error, FormulonError)

    def test_invalid_file_derives_from_io_and_value_error(self) -> None:
        self.assertTrue(issubclass(InvalidFileError, IoError))
        self.assertTrue(issubclass(InvalidFileError, ValueError))
        self.assertTrue(issubclass(IoError, FormulonError))
        self.assertFalse(issubclass(IoError, ValueError))

    def test_junk_bytes_still_raise_invalid_file(self) -> None:
        with self.assertRaises(InvalidFileError) as caught:
            Workbook.load(b"junk")
        self.assertIsInstance(caught.exception, IoError)
        self.assertIsInstance(caught.exception, ValueError)

    def test_the_class_is_exported(self) -> None:
        self.assertIn("IoError", formulon.__all__)


class TableTotalsRowTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)
        for a1, value in (("A1", "Region"), ("B1", "Amt"), ("A2", "East"), ("B2", 10), ("A3", "West"), ("B3", 30)):
            self.ws.set_value(a1, value)
        self.table = self.ws.add_table(TableInput(range="A1:B3", name="Sales"))
        self.table.update(
            columns=[TableColumnPatch(), TableColumnPatch(totals_row_function=TableTotalsRowFunction.SUM)]
        )

    def _rect(self) -> CellRect:
        return self.table.get_properties().range.rect

    def test_turning_the_row_on_grows_the_table_and_keeps_the_data(self) -> None:
        self.table.update(totals_row=True)
        self.assertEqual(self._rect(), CellRect(0, 0, 3, 1))
        self.assertTrue(self.table.get_properties().totals_row)
        self.assertEqual((self.ws.get_value("A3"), self.ws.get_value("B3")), ("West", 30.0))
        self.assertEqual(self.ws.get_formula("B4"), "=SUBTOTAL(109,[Amt])")
        self.assertEqual(self.ws.get_value("B4"), 40.0)

    def test_turning_the_row_off_shrinks_the_table(self) -> None:
        self.table.update(totals_row=True)
        self.table.update(totals_row=False)
        self.assertEqual(self._rect(), CellRect(0, 0, 2, 1))
        self.assertFalse(self.table.get_properties().totals_row)
        self.assertEqual(self.ws.get_value("B3"), 30.0)
        self.assertIsNone(self.ws.get_value("B4"))

    def test_an_occupied_row_below_shifts_down_within_the_table_columns(self) -> None:
        self.ws.set_value("B4", 99)
        self.table.update(totals_row=True)
        self.assertEqual(self._rect(), CellRect(0, 0, 3, 1))
        self.assertTrue(self.table.get_properties().totals_row)
        self.assertEqual(self.ws.get_value("B4"), 40.0)
        self.assertEqual(self.ws.get_value("B5"), 99.0)

    def test_turning_the_row_off_is_refused_while_a_formula_references_it(self) -> None:
        self.table.update(totals_row=True)
        self.ws.set_formula("D1", "=B4")
        with self.assertRaises(FormulonError) as caught:
            self.table.update(totals_row=False)
        self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
        self.assertEqual(self._rect(), CellRect(0, 0, 3, 1))
        self.assertTrue(self.table.get_properties().totals_row)
        self.assertEqual(self.ws.get_value("B4"), 40.0)

    def test_a_table_on_the_last_grid_row_is_refused(self) -> None:
        edge = self.ws.add_table(TableInput(range="D1048575:E1048576", name="Edge"))
        with self.assertRaises(FormulonError) as caught:
            edge.update(totals_row=True)
        self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
        self.assertEqual(edge.get_properties().range.rect, CellRect(1048574, 3, 1048575, 4))
        self.assertFalse(edge.get_properties().totals_row)

    def test_a_refused_update_leaves_the_table_as_it_was(self) -> None:
        before = self.table.get_properties()
        with self.assertRaises(InvalidArgumentError):
            self.table.update(range="A1:C9", totals_row=True)
        self.assertEqual(self.table.get_properties(), before)


class MergeRefusalTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_an_overlapping_merge_is_already_exists(self) -> None:
        self.ws.add_merge("A1:B2")
        for overlap in ("B2:C3", "A1:B2", "A2:A5"):
            with self.subTest(overlap=overlap), self.assertRaises(AlreadyExistsError):
                self.ws.add_merge(overlap)
        self.assertEqual(self.ws.list_merges(), [CellRect(0, 0, 1, 1)])

    def test_an_adjacent_merge_is_accepted(self) -> None:
        self.ws.add_merge("A1:B2")
        self.ws.add_merge("C1:D2")
        self.assertEqual(self.ws.list_merges(), [CellRect(0, 0, 1, 1), CellRect(0, 2, 1, 3)])

    def test_a_merge_over_a_table_is_a_precondition_failure(self) -> None:
        self.ws.set_value("A1", "H")
        self.ws.add_table(TableInput(range="A1:A3", name="T"))
        with self.assertRaises(FormulonError) as caught:
            self.ws.add_merge("A2:B2")
        self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
        self.assertEqual(self.ws.list_merges(), [])


class ReturnedKeyTests(unittest.TestCase):
    def test_add_merge_returns_the_stored_rectangle(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.sheet(0).add_merge("B2:C3"), CellRect(1, 1, 2, 2))
            self.assertEqual(wb.sheet(0).add_merge(CellRect(5, 5, 6, 6)), CellRect(5, 5, 6, 6))

    def test_add_data_table_returns_the_top_left_cell(self) -> None:
        with Workbook.create_default() as wb:
            key = wb.sheet(0).add_data_table(DataTableInput("H2:H4", column_input="A1"))
            self.assertEqual(key, CellAddress(1, 7))

    def test_add_person_returns_the_stored_id(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.add_person(Person(id=_PERSON, display_name="A")), _PERSON.upper())

    def test_add_field_shared_item_returns_the_new_index(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.pivot_cache(wb.add_pivot_cache().cache_id)
            cache.add_field("Region")
            self.assertEqual(cache.add_field_shared_item(0, "East"), 0)
            self.assertEqual(cache.add_field_shared_item(0, "West"), 1)


class CfRuleKeyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def _add(self, **extra: str) -> str:
        return self.ws.add_cf_rule(CfExpressionRuleInput(ranges="A1:B2", formula1="TRUE", **extra))

    def test_add_returns_a_generated_guid_that_reads_back(self) -> None:
        rule_id = self._add()
        self.assertIsInstance(rule_id, str)
        self.assertRegex(rule_id, _GUID)
        self.assertEqual([r.id for r in self.ws.list_cf_rules()], [rule_id])

    def test_a_malformed_id_is_refused(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self._add(id="abc")
        self.assertEqual(self.ws.list_cf_rules(), [])

    def test_a_supplied_guid_is_kept_in_either_case(self) -> None:
        supplied = "{0a1b2c3d-0000-4000-8000-00000000000f}"
        stored = self._add(id=supplied)
        self.assertEqual(stored.upper(), supplied.upper())

    def test_a_used_id_is_replaced_by_a_fresh_one(self) -> None:
        first = self._add()
        second = self._add(id=first)
        self.assertNotEqual(first, second)
        self.assertRegex(second, _GUID)

    def test_remove_and_update_by_id(self) -> None:
        keep = self._add()
        gone = self._add()
        self.ws.update_cf_rule(keep, stop_if_true=True)
        self.assertTrue(next(r for r in self.ws.list_cf_rules() if r.id == keep).stop_if_true)
        self.ws.remove_cf_rule(gone)
        self.assertEqual([r.id for r in self.ws.list_cf_rules()], [keep])

    def test_an_unknown_id_is_not_found_and_a_number_is_a_position(self) -> None:
        self._add()
        with self.assertRaises(NotFoundError):
            self.ws.remove_cf_rule(_ZERO_GUID)
        with self.assertRaises(IndexOutOfRangeError):
            self.ws.remove_cf_rule(5)
        self.assertEqual(len(self.ws.list_cf_rules()), 1)
        self.ws.remove_cf_rule(0)
        self.assertEqual(self.ws.list_cf_rules(), [])


class ProtectedRangeKeyTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_add_returns_the_name(self) -> None:
        self.assertEqual(self.ws.add_protected_range(ProtectedRangeInput(name="P", ranges="A1")), "P")

    def test_update_and_remove_by_name(self) -> None:
        self.ws.add_protected_range(ProtectedRangeInput(name="P", ranges="A1"))
        self.ws.add_protected_range(ProtectedRangeInput(name="Q", ranges="B2"))
        self.ws.update_protected_range("Q", ranges="C3")
        self.assertEqual({p.name: p.ranges for p in self.ws.list_protected_ranges()}["Q"], [CellRect(2, 2, 2, 2)])
        self.ws.remove_protected_range("P")
        self.assertEqual([p.name for p in self.ws.list_protected_ranges()], ["Q"])

    def test_unknown_name_duplicate_name_and_position_past_the_end(self) -> None:
        self.ws.add_protected_range(ProtectedRangeInput(name="P", ranges="A1"))
        with self.assertRaises(NotFoundError):
            self.ws.remove_protected_range("Nope")
        with self.assertRaises(AlreadyExistsError):
            self.ws.add_protected_range(ProtectedRangeInput(name="P", ranges="B2"))
        with self.assertRaises(IndexOutOfRangeError):
            self.ws.remove_protected_range(5)
        self.assertEqual([p.name for p in self.ws.list_protected_ranges()], ["P"])


class NameFoldingTests(unittest.TestCase):
    def test_a_custom_table_style_matches_by_the_unicode_fold(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_table_style(TableStyle(name="Ä1"))
            found = wb.get_table_style("ä1")
            self.assertIsNotNone(found)
            self.assertEqual(found.name, "Ä1")
            wb.set_table_style(TableStyle(name="ä1", pivot=False))
            styles = wb.list_table_styles()
            self.assertEqual(len(styles), 1)
            self.assertFalse(styles[0].pivot)
            self.assertEqual(styles[0].name, "ä1")
            self.assertIsNone(wb.get_table_style("nope"))

    def test_a_defined_name_lookup_folds_non_ascii_case(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name(DefinedNameInput(name="Ä", formula="=1"))
            self.assertEqual(wb.get_defined_name("ä").name, "Ä")
            wb.set_defined_name(DefinedNameInput(name="ä", formula="=2"))
            self.assertEqual(wb.get_defined_name("Ä").name, "ä")
            self.assertIsNone(wb.get_defined_name("ö"))

    def test_a_sheet_whose_name_differs_only_by_width_is_already_exists(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(AlreadyExistsError):
                wb.add_sheet("\uff33\uff48\uff45\uff45\uff54\uff11")
            self.assertEqual(len(wb.list_sheets()), 1)

    def test_a_renamed_table_column_collides_by_the_fold(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for a1, value in (("A1", "Ä"), ("B1", "x"), ("A2", 1), ("B2", 2)):
                ws.set_value(a1, value)
            table = ws.add_table(TableInput(range="A1:B2", name="T"))
            stored = table.rename_column(1, "ä")
            self.assertNotEqual(stored.casefold(), "ä")
            self.assertTrue(stored.startswith("ä"))
            self.assertEqual([c.name for c in table.get_properties().columns], ["Ä", stored])


class BlankEvaluationTests(unittest.TestCase):
    def test_a_blank_cell_the_formula_reads_is_zero(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertEqual(ws.evaluate_formula("A2", "=B9"), 0)
            self.assertIsNone(ws.get_value("B9"))
            self.assertEqual(ws.evaluate_formula_array("A2", "=B9:B10"), [[0], [0]])


if __name__ == "__main__":
    unittest.main()
