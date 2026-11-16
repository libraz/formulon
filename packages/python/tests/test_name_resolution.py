"""A name given to a facade method resolves as the engine resolves sheet names.

Exact first, then case-only, then a unique width fold; two entries that collide only under the width fold resolve
nothing.
"""

from __future__ import annotations

import unittest
from typing import Any, Callable, List, Sequence

import formulon
from formulon import (
    CellAddress,
    CellRect,
    CfExpressionRuleInput,
    DrawingAnchorInput,
    NotFoundError,
    PivotCache,
    PivotInput,
    PivotWorksheetSourceInput,
    Sheet,
    SlicerInput,
    StatusCode,
    TableInput,
    Workbook,
)

# Full-width spellings: 'Sales' and its lower-case form fold to these under the width fold only.
WIDE_LOWER = "ｓａｌｅｓ"


class NameResolutionTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def add_cache(self, names: Sequence[str]) -> PivotCache:
        """A pivot cache over the header row written from ``names``, with a field per name."""
        for col, name in enumerate(names):
            self.ws.set_text(0, col, name)
            self.ws.set_number(1, col, col + 1)
        cache = self.wb.add_pivot_cache()
        cache.set_worksheet_source(
            PivotWorksheetSourceInput(range=f"A1:{chr(ord('A') + len(names) - 1)}2", sheet="Sheet1")
        )
        for name in names:
            cache.add_field(name)
        return cache

    def add_bare_cache(self, names: Sequence[str]) -> PivotCache:
        """A pivot cache with a field per name and no worksheet source, so the names are taken as given."""
        cache = self.wb.add_pivot_cache()
        for name in names:
            cache.add_field(name)
        return cache

    def assert_not_found(self, call: Callable[[], Any]) -> None:
        with self.assertRaises(NotFoundError) as caught:
            call()
        self.assertEqual(caught.exception.code, StatusCode.NOT_FOUND)


class PivotNameTests(NameResolutionTestCase):
    def pivots(self, names: List[str]) -> List[formulon.Pivot]:
        cache = self.add_cache(["Region"])
        return [
            self.ws.add_pivot(PivotInput(name=name, cache_id=cache.cache_id, anchor=CellAddress(5, i * 4)))
            for i, name in enumerate(names)
        ]

    def test_exact_case_only_and_unique_width_fold_find_the_pivot(self) -> None:
        (pivot,) = self.pivots(["Sales"])
        for name in ("Sales", "sALES", WIDE_LOWER):
            self.assertIs(self.ws.pivot(name), pivot, name)

    def test_a_miss_is_not_found(self) -> None:
        self.pivots(["Sales"])
        self.assert_not_found(lambda: self.ws.pivot("Nope"))

    def test_remove_pivot_takes_the_same_names(self) -> None:
        self.pivots(["Sales"])
        self.ws.remove_pivot(WIDE_LOWER)
        self.assertEqual(self.ws.list_pivots(), [])


class TableNameTests(NameResolutionTestCase):
    def tables(self, sheet: Sheet, names: List[str]) -> List[formulon.Table]:
        tables = []
        for i, name in enumerate(names):
            sheet.set_text(0, i * 3, "Key")
            sheet.set_number(1, i * 3, 1)
            tables.append(sheet.add_table(TableInput(range=CellRect(0, i * 3, 1, i * 3), name=name)))
        return tables

    def test_exact_case_only_and_unique_width_fold_find_the_table(self) -> None:
        (table,) = self.tables(self.ws, ["Sales"])
        for name in ("Sales", "sALES", WIDE_LOWER):
            self.assertIs(self.ws.table(name), table, name)

    def test_a_miss_is_not_found(self) -> None:
        self.tables(self.ws, ["Sales"])
        self.assert_not_found(lambda: self.ws.table("Nope"))

    def test_a_table_on_another_sheet_is_not_found_by_this_sheet(self) -> None:
        other = self.wb.add_sheet("Other")
        self.tables(other, ["Sales"])
        self.assert_not_found(lambda: self.ws.table("Sales"))
        self.assertEqual(self.wb.table("Sales").name, "Sales")

    def test_remove_table_takes_the_same_names(self) -> None:
        self.tables(self.ws, ["Sales"])
        self.ws.remove_table(WIDE_LOWER)
        self.assertEqual(self.ws.list_tables(), [])


class PivotFieldNameTests(NameResolutionTestCase):
    def shared(self, cache: PivotCache, field: str) -> int:
        return len(cache.list_field_shared_items(field))

    def test_exact_case_only_and_unique_width_fold_find_the_field(self) -> None:
        cache = self.add_bare_cache(["Sales", "Other"])
        cache.add_field_shared_item(0, "x")
        for name in ("Sales", "sALES", WIDE_LOWER):
            self.assertEqual(self.shared(cache, name), 1, name)
        self.assertEqual(self.shared(cache, "Other"), 0)

    def test_an_exact_name_wins_over_a_case_only_one(self) -> None:
        cache = self.add_bare_cache(["a", "A"])
        cache.add_field_shared_item(1, "x")
        self.assertEqual(self.shared(cache, "A"), 1)
        self.assertEqual(self.shared(cache, "a"), 0)

    def test_width_twins_resolve_nothing_and_a_miss_is_not_found(self) -> None:
        # Each twin is "sales" with one letter fullwidth, so "sales" reaches them only through the width fold.
        cache = self.add_bare_cache(["\uff53ales", "s\uff41les"])
        self.assert_not_found(lambda: cache.list_field_shared_items("sales"))
        self.assert_not_found(lambda: cache.list_field_shared_items("Nope"))

    def test_a_pivot_method_takes_the_same_names(self) -> None:
        cache = self.add_cache(["Sales"])
        pivot = self.ws.add_pivot(PivotInput(name="P", cache_id=cache.cache_id, anchor=CellAddress(5, 0)))
        self.assertEqual(pivot.get_field(WIDE_LOWER).cache_field_index, 0)
        self.assert_not_found(lambda: pivot.get_field("Nope"))


class SlicerColumnNameTests(NameResolutionTestCase):
    def column_of(self, headers: List[str], column: str) -> str:
        for col, name in enumerate(headers):
            self.ws.set_text(0, col, name)
            self.ws.set_number(1, col, col)
        table = self.ws.add_table(TableInput(range=CellRect(0, 0, 1, len(headers) - 1), name="T"))
        anchor = DrawingAnchorInput(from_row=20, from_col=8, to_row=30, to_col=10)
        return self.ws.add_slicer(table, column, SlicerInput(anchor=anchor)).get_properties().column_name

    def test_exact_case_only_and_unique_width_fold_find_the_column(self) -> None:
        for name in ("Sales", "sALES", WIDE_LOWER):
            with self.subTest(name=name):
                self.setUp()
                self.assertEqual(self.column_of(["Key", "Sales"], name), "Sales")

    def test_a_miss_is_not_found(self) -> None:
        self.assert_not_found(lambda: self.column_of(["Sales"], "Nope"))


class CfRuleIdTests(NameResolutionTestCase):
    def setUp(self) -> None:
        super().setUp()
        self.rule_id = self.ws.add_cf_rule(CfExpressionRuleInput(ranges="A1", formula1="A1>1"))

    def test_an_exact_id_and_its_other_case_spelling_find_the_rule(self) -> None:
        for spelling in (self.rule_id, self.rule_id.lower(), self.rule_id.upper()):
            self.ws.update_cf_rule(spelling, formula1="A1>2")
            self.assertEqual(self.ws.get_cf_rule(self.rule_id).formula1, "=A1>2", spelling)
            self.ws.update_cf_rule(self.rule_id, formula1="A1>1")

    def test_a_width_folded_spelling_and_an_unknown_id_are_not_found(self) -> None:
        wide = "".join(chr(ord(c) + 0xFEE0) if c.isascii() and c.isalnum() else c for c in self.rule_id)
        self.assert_not_found(lambda: self.ws.update_cf_rule(wide, formula1="A1>2"))
        self.assert_not_found(lambda: self.ws.update_cf_rule("{00000000-0000-4000-8000-000000000000}", formula1="A1>2"))


class WorksheetSourceSheetTests(NameResolutionTestCase):
    def test_a_name_that_resolves_finds_the_sheet(self) -> None:
        self.wb.add_sheet("Sales")
        cache = self.wb.add_pivot_cache()
        for name in ("Sales", "sALES", WIDE_LOWER):
            cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B3", sheet=name))
            source = cache.get_worksheet_source()
            assert source is not None and source.range is not None
            self.assertEqual(source.range.sheet, 1, name)

    def test_an_unresolved_name_is_accepted_and_reads_back_without_a_range(self) -> None:
        self.wb.add_sheet("Sales")
        cache = self.wb.add_pivot_cache()
        cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B3", sheet="Gone"))
        source = cache.get_worksheet_source()
        assert source is not None
        self.assertIsNone(source.range)

    def test_a_qualified_range_and_a_sheet_that_resolve_to_one_sheet_agree(self) -> None:
        self.wb.add_sheet("Sales")
        self.wb.add_sheet("Other")
        cache = self.wb.add_pivot_cache()
        cache.set_worksheet_source(PivotWorksheetSourceInput(range="Sales!A1:B3", sheet=WIDE_LOWER))
        source = cache.get_worksheet_source()
        assert source is not None and source.range is not None
        self.assertEqual(source.range.sheet, 1)
        with self.assertRaises(formulon.InvalidArgumentError):
            cache.set_worksheet_source(PivotWorksheetSourceInput(range="Sales!A1:B3", sheet="Other"))


if __name__ == "__main__":
    unittest.main()
