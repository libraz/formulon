"""Table.update(columns=...): per-column totals settings and calculated-column formulas merged element by element, the single-cell write
rules that keep a column's calculated mark and totals settings in step with its cells, and the save round trips."""

from __future__ import annotations

import unittest
from typing import Any, List, Tuple

from formulon import (
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    Table,
    TableColumnPatch,
    TableInput,
    TableTotalsRowFunction,
    Workbook,
    WorkbookFormat,
)

Fn = TableTotalsRowFunction


def _calc_table(wb: Workbook, totals_row: bool = False) -> Tuple[Any, Table]:
    """Headers a, b, c on row 2, data on rows 3..5 (a = 1..3, b = 10..30, c empty); table Tbl over A2:C5, or A2:C6
    with a totals row. The first data row is Excel row 3, so ``=A3*10`` is the first-row spelling."""
    ws = wb.sheet(0)
    for a1, value in (("A2", "a"), ("B2", "b"), ("C2", "c")):
        ws.set_text(a1, value)
    for row, (a, b) in zip((3, 4, 5), ((1, 10), (2, 20), (3, 30))):
        ws.set_number(f"A{row}", float(a))
        ws.set_number(f"B{row}", float(b))
    table = ws.add_table(TableInput(range="A2:C6" if totals_row else "A2:C5", name="Tbl", totals_row=totals_row))
    return ws, table


def _at(index: int, **fields: Any) -> List[TableColumnPatch]:
    """A ``columns`` argument whose element ``index`` carries ``fields``; the columns before it are left alone."""
    return [TableColumnPatch() for _ in range(index)] + [TableColumnPatch(**fields)]


def _calc(column: int, formula: Any) -> List[TableColumnPatch]:
    return _at(column, calculated_formula=formula)


def _col(table: Table, index: int) -> Any:
    return table.get_properties().columns[index]


# Pairwise rows over write kind x column x prior state.
TOTALS_STEPS = [
    ("blank", 2, "label"),
    ("constant_formula", 2, "function"),
    ("subtotal_own", 1, "none"),
    ("custom_formula", 0, "label"),
    ("number", 1, "custom"),
    ("subtotal_other_column", 2, "none"),
    ("constant_formula", 0, "custom"),
    ("custom_formula", 1, "function"),
    ("subtotal_own", 0, "function"),
    ("bool", 1, "label"),
    ("constant_formula", 1, "none"),
    ("number", 0, "none"),
    ("number", 2, "label"),
    ("blank", 0, "none"),
    ("bool", 0, "none"),
    ("subtotal_other_column", 1, "label"),
    ("constant_formula", 2, "label"),
    ("custom_formula", 2, "none"),
    ("text", 2, "none"),
    ("subtotal_own", 2, "custom"),
    ("bool", 2, "function"),
    ("blank", 1, "function"),
    ("subtotal_other_column", 0, "custom"),
    ("text", 0, "custom"),
    ("custom_formula", 0, "custom"),
    ("text", 1, "none"),
    ("number", 2, "function"),
    ("text", 1, "function"),
    ("blank", 0, "custom"),
    ("bool", 2, "custom"),
    ("subtotal_other_column", 0, "function"),
    ("subtotal_own", 0, "label"),
    ("text", 0, "label"),
]
NAMES = ("a", "b", "c")
COLS = ("A", "B", "C")
BUILT_INS = {
    109: Fn.SUM,
    105: Fn.MIN,
    104: Fn.MAX,
    101: Fn.AVERAGE,
    103: Fn.COUNT,
    102: Fn.COUNT_NUMBERS,
    107: Fn.STD_DEV,
    110: Fn.VAR,
}
PRIOR_WRITE = {"none": "blank", "label": "text", "function": "subtotal_own", "custom": "constant_formula"}
EXPECTED_LABEL = {"number": "5.5", "text": "Total", "bool": "TRUE"}


def _write_totals(ws: Any, col: int, kind: str) -> None:
    at = f"{COLS[col]}6"
    own = f"Tbl[{NAMES[col]}]"
    other = f"Tbl[{NAMES[(col + 1) % 3]}]"
    if kind == "number":
        ws.set_number(at, 5.5)
    elif kind == "text":
        ws.set_text(at, "Total")
    elif kind == "bool":
        ws.set_bool(at, True)
    elif kind == "subtotal_own":
        ws.set_formula(at, f"=SUBTOTAL(104,{own})")
    elif kind == "subtotal_other_column":
        ws.set_formula(at, f"=SUBTOTAL(109,{other})")
    elif kind == "constant_formula":
        ws.set_formula(at, "=5")
    elif kind == "custom_formula":
        ws.set_formula(at, f"=SUM({own})*2")
    elif kind == "blank":
        ws.set_blank(at)
    else:
        raise AssertionError(kind)


class ColumnPatchCalculatedFormulaTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_sets_the_formula_fills_the_column_and_overwrites_a_constant(self) -> None:
        ws, table = _calc_table(self.wb)
        self.assertIsNone(_col(table, 2).calculated_formula)
        ws.set_number("C4", 99.0)
        table.update(columns=_calc(2, "=A3*10"))
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual([ws.get_formula(a1) for a1 in ("C3", "C4", "C5")], ["=A3*10", "=A4*10", "=A5*10"])
        self.wb.recalc()
        self.assertEqual([ws.get_value(a1) for a1 in ("C3", "C4", "C5")], [10.0, 20.0, 30.0])
        self.assertIsNone(_col(table, 0).calculated_formula)
        self.assertIsNone(_col(table, 1).calculated_formula)
        self.assertEqual(ws.get_value("C2"), "c")

    def test_reads_back_in_the_first_row_spelling_whichever_form_was_given(self) -> None:
        cases = (
            ("=A3*10", "=A3*10"),
            ("A3*10", "=A3*10"),
            ("=[@a]*10", "=[@a]*10"),
            ("=Tbl[[#This Row],[a]]*10", "=[@a]*10"),
            ("=$A$3*10", "=$A$3*10"),
            ("=A$3*10+$B3", "=A$3*10+$B3"),
        )
        for given, read in cases:
            with self.subTest(given), Workbook.create_default() as wb:
                _, table = _calc_table(wb)
                table.update(columns=_calc(2, given))
                self.assertEqual(_col(table, 2).calculated_formula, read)
        ws, table = _calc_table(self.wb)
        table.update(columns=_calc(2, "=$A$3*10"))
        self.wb.recalc()
        self.assertEqual(ws.get_value("C5"), 10.0)
        table.update(columns=_calc(2, "=[@a]*10"))
        self.wb.recalc()
        self.assertEqual((ws.get_value("C3"), ws.get_value("C5")), (10.0, 30.0))

    def test_none_removes_the_mark_and_leaves_the_cells(self) -> None:
        ws, table = _calc_table(self.wb)
        table.update(columns=_calc(2, "=A3*10"))
        table.update(columns=_calc(2, None))
        self.assertIsNone(_col(table, 2).calculated_formula)
        self.assertEqual([ws.get_formula(a1) for a1 in ("C3", "C4", "C5")], ["=A3*10", "=A4*10", "=A5*10"])
        self.wb.recalc()
        self.assertEqual(ws.get_value("C5"), 30.0)

    def test_a_changed_formula_refills_and_the_same_formula_keeps_constants(self) -> None:
        ws, table = _calc_table(self.wb)
        table.update(columns=_calc(2, "=A3*10"))
        ws.set_number("C4", 99.0)
        table.update(columns=_calc(2, "=A3*10"))
        self.assertEqual(ws.get_value("C4"), 99.0)
        table.update(columns=_calc(2, "=A$3*10"))
        self.assertEqual(ws.get_formula("C4"), "=A$3*10")
        self.wb.recalc()
        self.assertEqual([ws.get_value(a1) for a1 in ("C3", "C4", "C5")], [10.0, 10.0, 10.0])
        table.update(columns=_calc(2, None))
        ws.set_number("C5", 7.0)
        table.update(columns=_calc(2, "=$A$3*10"))
        self.assertEqual(ws.get_formula("C5"), "=$A$3*10")

    def test_other_fields_leave_the_mark_and_an_empty_element_is_a_no_op(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(columns=_calc(2, "=A3*10"))
        table.update(columns=_at(2, totals_row_function=Fn.MAX))
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual(_col(table, 2).totals_row_function, Fn.MAX)
        table.update(columns=[TableColumnPatch(), TableColumnPatch(), TableColumnPatch()])
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual(_col(table, 2).totals_row_function, Fn.MAX)
        self.assertEqual(ws.get_formula("C4"), "=A4*10")

    def test_an_empty_formula_is_refused_and_keeps_the_mark(self) -> None:
        ws, table = _calc_table(self.wb)
        table.update(columns=_calc(2, "=A3*10"))
        with self.assertRaises(InvalidArgumentError):
            table.update(columns=_calc(2, ""))
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual(ws.get_formula("C4"), "=A4*10")

    def test_unusable_formulas_are_refused_and_change_nothing(self) -> None:
        ws, table = _calc_table(self.wb)
        ws.set_number("C4", 99.0)
        with self.assertRaises(FormulonError):
            table.update(columns=_calc(2, "=SUM(("))
        for formula in ("=SEQUENCE(3)", "=FILTER(A3:A5,A3:A5>1)", "=A1048576+1"):
            with self.subTest(formula), self.assertRaises(InvalidArgumentError):
                table.update(columns=_calc(2, formula))
        self.assertIsNone(_col(table, 2).calculated_formula)
        self.assertEqual(ws.get_value("C4"), 99.0)
        self.assertIsNone(ws.get_formula("C3"))

    def test_a_table_without_data_rows_takes_the_mark_only(self) -> None:
        ws = self.wb.sheet(0)
        for a1, value in (("A2", "a"), ("B2", "b"), ("C2", "c")):
            ws.set_text(a1, value)
        table = ws.add_table(TableInput(range="A2:C3", name="Tbl", totals_row=True))
        table.update(columns=_calc(2, "=A3*10"))
        self.assertIn("*10", _col(table, 2).calculated_formula)
        self.assertNotIn("*10", ws.get_formula("C3") or "")
        table.update(columns=_calc(2, None))
        self.assertIsNone(_col(table, 2).calculated_formula)


class ColumnPatchValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_an_array_longer_than_the_column_count_is_refused_and_nothing_applies(self) -> None:
        ws, table = _calc_table(self.wb, True)
        with self.assertRaises(IndexOutOfRangeError):
            table.update(
                columns=[
                    TableColumnPatch(),
                    TableColumnPatch(totals_row_function=Fn.SUM),
                    TableColumnPatch(calculated_formula="=A3*10"),
                    TableColumnPatch(),
                ]
            )
        self.assertIsNone(_col(table, 2).calculated_formula)
        self.assertEqual(_col(table, 1).totals_row_function, Fn.NONE)
        self.assertIsNone(ws.get_formula("C3"))
        self.assertIsNone(ws.get_formula("B6"))

    def test_a_shorter_array_patches_the_leading_columns_and_an_empty_one_changes_nothing(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(columns=[TableColumnPatch(totals_row_label="Total")])
        self.assertEqual(_col(table, 0).totals_row_label, "Total")
        self.assertIsNone(_col(table, 1).totals_row_label)
        table.update(
            columns=[
                TableColumnPatch(),
                TableColumnPatch(totals_row_function=Fn.SUM),
                TableColumnPatch(calculated_formula="=A3*10"),
            ]
        )
        table.update(columns=[])
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual(_col(table, 1).totals_row_function, Fn.SUM)
        self.assertEqual(ws.get_formula("B6"), "=SUBTOTAL(109,[b])")

    def test_the_columns_are_checked_before_the_range_or_flag_changes(self) -> None:
        _, table = _calc_table(self.wb)
        before = table.get_properties().range.rect
        with self.assertRaises(IndexOutOfRangeError):
            table.update(totals_row=True, columns=[TableColumnPatch()] * 4)
        self.assertFalse(table.get_properties().totals_row)
        self.assertEqual(table.get_properties().range.rect, before)
        for formula in ("=SUM((", "=SEQUENCE(3)", "=A1048576+1"):
            with self.subTest(formula), self.assertRaises(FormulonError):
                table.update(totals_row=True, columns=_calc(2, formula))
            self.assertFalse(table.get_properties().totals_row)
            self.assertEqual(table.get_properties().range.rect, before)

    def test_several_columns_apply_in_one_call_and_the_totals_row_is_not_filled(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(
            columns=[
                TableColumnPatch(totals_row_label="Total"),
                TableColumnPatch(totals_row_function=Fn.SUM),
                TableColumnPatch(calculated_formula="=A3*10"),
            ]
        )
        self.assertEqual(_col(table, 0).totals_row_label, "Total")
        self.assertEqual(_col(table, 1).totals_row_function, Fn.SUM)
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual(ws.get_formula("B6"), "=SUBTOTAL(109,[b])")
        self.assertEqual(ws.get_formula("C5"), "=A5*10")
        # The totals row holds the default total, not the calculated formula.
        self.assertNotIn("*10", ws.get_formula("C6") or "")

    def test_none_clears_each_field_and_a_non_custom_function_drops_the_formula(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(columns=_at(1, totals_row_function=Fn.SUM, totals_row_label="Sum"))
        self.assertEqual(ws.get_formula("B6"), "=SUBTOTAL(109,[b])")
        table.update(columns=_at(1, totals_row_function=None))
        self.assertEqual(_col(table, 1).totals_row_function, Fn.NONE)
        self.assertEqual(ws.get_value("B6"), "Sum")
        table.update(columns=_at(1, totals_row_label=None))
        self.assertIsNone(_col(table, 1).totals_row_label)
        self.assertIsNone(ws.get_value("B6"))
        table.update(columns=_at(1, totals_row_function=Fn.CUSTOM, totals_row_formula="=SUM(Tbl[b])*2"))
        self.assertEqual(_col(table, 1).totals_row_formula, "=SUM([b])*2")
        table.update(columns=_at(1, totals_row_function=Fn.MAX))
        self.assertIsNone(_col(table, 1).totals_row_formula)
        table.update(columns=_at(1, totals_row_function=Fn.CUSTOM, totals_row_formula="=5"))
        with self.assertRaises(InvalidArgumentError):
            table.update(columns=_at(1, totals_row_formula=None))
        self.assertEqual(_col(table, 1).totals_row_formula, "=5")

    def test_a_custom_function_without_a_formula_and_empty_strings_are_refused(self) -> None:
        _, table = _calc_table(self.wb, True)
        with self.assertRaises(InvalidArgumentError):
            table.update(columns=_at(1, totals_row_function=Fn.CUSTOM))
        with self.assertRaises(InvalidArgumentError):
            table.update(columns=_at(0, totals_row_label=""))
        with self.assertRaises(InvalidArgumentError):
            table.update(columns=_at(1, totals_row_formula=""))
        self.assertEqual(_col(table, 1).totals_row_function, Fn.NONE)

    def test_none_totals_fields_clear_a_shown_cell_and_hidden_totals_are_metadata_only(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(columns=_at(1, totals_row_function=Fn.SUM))
        self.assertEqual(ws.get_formula("B6"), "=SUBTOTAL(109,[b])")
        table.update(columns=_at(1, totals_row_function=None))
        self.assertEqual(_col(table, 1).totals_row_function, Fn.NONE)
        self.assertIsNone(ws.get_value("B6"))
        with Workbook.create_default() as hidden:
            hidden_ws, hidden_table = _calc_table(hidden)
            hidden_table.update(columns=_at(1, totals_row_function=Fn.MAX))
            self.assertEqual(_col(hidden_table, 1).totals_row_function, Fn.MAX)
            self.assertIsNone(hidden_ws.get_value("B6"))
            self.assertEqual(hidden_table.get_properties().range.rect.last_row, 4)

    def test_feeding_the_read_range_flags_and_columns_back_changes_nothing(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(
            columns=[
                TableColumnPatch(),
                TableColumnPatch(totals_row_function=Fn.SUM),
                TableColumnPatch(calculated_formula="=A3*10"),
            ]
        )
        ws.set_number("C4", 99.0)
        props = table.get_properties()
        table.update(
            range=props.range.rect, header_row=props.header_row, totals_row=props.totals_row, columns=props.columns
        )
        self.assertEqual(table.get_properties().range.rect, props.range.rect)
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        self.assertEqual(_col(table, 1).totals_row_function, Fn.SUM)
        self.assertEqual(ws.get_value("C4"), 99.0)
        self.assertEqual(ws.get_formula("B6"), "=SUBTOTAL(109,[b])")

    def test_a_malformed_element_is_a_type_error_and_changes_nothing(self) -> None:
        _, table = _calc_table(self.wb)
        bad = {
            "an unknown key": [{}, {}, {"calculated_formula": "=A3*10", "bogus": 1}],
            "the old totals nesting": [{}, {"totals": {"totals_row_function": Fn.SUM}}],
            "the old column key": [{"column": 2, "calculated_formula": "=A3*10"}],
            "a non-record element": [{}, "not a patch"],
        }
        for what, columns in bad.items():
            with self.subTest(what), self.assertRaises(TypeError):
                table.update(columns=columns)  # type: ignore[arg-type]
        self.assertIsNone(_col(table, 2).calculated_formula)
        self.assertEqual(_col(table, 1).totals_row_function, Fn.NONE)

    def test_a_refusal_names_the_element_and_the_field(self) -> None:
        _, table = _calc_table(self.wb)
        refusals = [
            ([{}, {"totals_row_function": "x"}], "Table.update columns[1].totals_row_function must be an integer"),
            ([{}, {"foo": 1}], "Table.update columns[1] has no field foo"),
            ([{}, 5], "Table.update columns must be a sequence of TableColumnPatch"),
        ]
        for columns, message in refusals:
            with self.subTest(message), self.assertRaises(TypeError) as ctx:
                table.update(columns=columns)  # type: ignore[arg-type]
            self.assertEqual(str(ctx.exception), message)


class CellWriteRuleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_a_single_cell_write_keeps_the_calculated_column(self) -> None:
        writes = {
            "a constant": lambda ws: ws.set_number("C4", 7.0),
            "a text": lambda ws: ws.set_text("C4", "x"),
            "a cleared cell": lambda ws: ws.set_blank("C4"),
            "another formula": lambda ws: ws.set_formula("C4", "=B4+1"),
        }
        for what, write in writes.items():
            with self.subTest(what), Workbook.create_default() as wb:
                ws, table = _calc_table(wb)
                table.update(columns=_calc(2, "=A3*10"))
                write(ws)
                self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
                self.assertEqual(ws.get_formula("C3"), "=A3*10")
                self.assertEqual(ws.get_formula("C5"), "=A5*10")

    def test_blanking_the_last_non_blank_cell_removes_the_mark(self) -> None:
        ws, table = _calc_table(self.wb)
        table.update(columns=_calc(2, "=A3*10"))
        ws.set_blank("C3")
        ws.set_blank("C5")
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
        ws.set_blank("C4")
        self.assertIsNone(_col(table, 2).calculated_formula)
        ws.set_blank("A3")
        self.assertIsNone(_col(table, 0).calculated_formula)

    def test_blanking_a_totals_cell_does_not_drop_the_calculated_column(self) -> None:
        ws, table = _calc_table(self.wb, True)
        table.update(columns=_calc(2, "=A3*10"))
        ws.set_text("C6", "x")
        ws.set_blank("C6")
        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")


class TotalsCellWriteTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_a_totals_cell_write_sets_the_column_totals_settings(self) -> None:
        ws, table = _calc_table(self.wb, True)

        ws.set_number("B6", 5.0)
        self.assertEqual((_col(table, 1).totals_row_function, _col(table, 1).totals_row_label), (Fn.NONE, "5"))
        self.assertEqual(ws.get_value("B6"), "5")
        ws.set_number("B6", 5.5)
        self.assertEqual(_col(table, 1).totals_row_label, "5.5")

        ws.set_text("A6", "Total")
        self.assertEqual((_col(table, 0).totals_row_function, _col(table, 0).totals_row_label), (Fn.NONE, "Total"))

        ws.set_formula("B6", "=SUBTOTAL(109,Tbl[b])")
        self.assertEqual(_col(table, 1).totals_row_function, Fn.SUM)
        self.assertIsNone(_col(table, 1).totals_row_formula)

        for formula, stored in (
            ("=5", "=5"),
            ("=SUBTOTAL(9,Tbl[b])", "=SUBTOTAL(9,[b])"),
            ("=SUM(Tbl[b])*2", "=SUM([b])*2"),
        ):
            with self.subTest(formula):
                ws.set_formula("B6", formula)
                self.assertEqual(_col(table, 1).totals_row_function, Fn.CUSTOM)
                self.assertEqual(_col(table, 1).totals_row_formula, stored)

        ws.set_bool("B6", True)
        self.assertEqual((_col(table, 1).totals_row_function, _col(table, 1).totals_row_label), (Fn.NONE, "TRUE"))

        ws.set_blank("B6")
        column = _col(table, 1)
        self.assertEqual(
            (column.totals_row_function, column.totals_row_label, column.totals_row_formula), (Fn.NONE, None, None)
        )

        # A data cell is not a totals cell.
        table.update(columns=_at(1, totals_row_function=Fn.SUM))
        ws.set_text("B5", "x")
        self.assertEqual(_col(table, 1).totals_row_function, Fn.SUM)

    def _assert_agrees(self, ws: Any, table: Table, col: int, where: str) -> None:
        read = _col(table, col)
        formula = ws.get_formula(f"{COLS[col]}6")
        value = ws.get_value(f"{COLS[col]}6")
        if formula is not None:
            expected = Fn.CUSTOM
            for code, fn in BUILT_INS.items():
                if formula == f"=SUBTOTAL({code},[{NAMES[col]}])":
                    expected = fn
            self.assertEqual(read.totals_row_function, expected, f"{where}: {formula}")
            self.assertEqual(read.totals_row_formula, formula if expected == Fn.CUSTOM else None, where)
        elif value is None:
            self.assertEqual(read.totals_row_function, Fn.NONE, where)
            self.assertIsNone(read.totals_row_label, where)
            self.assertIsNone(read.totals_row_formula, where)
        else:
            self.assertEqual(read.totals_row_function, Fn.NONE, where)
            self.assertIsInstance(value, str, f"{where}: a value in a totals cell is stored as its label text")
            self.assertEqual(read.totals_row_label, value, where)

    def test_the_totals_settings_agree_with_the_totals_cell_after_every_write(self) -> None:
        for i, (kind, col, prior) in enumerate(TOTALS_STEPS):
            with self.subTest(step=i), Workbook.create_default() as wb:
                ws, table = _calc_table(wb, True)
                _write_totals(ws, col, PRIOR_WRITE[prior])
                self._assert_agrees(ws, table, col, f"step {i} prior")
                _write_totals(ws, col, kind)
                self._assert_agrees(ws, table, col, f"step {i} write")
                if kind in EXPECTED_LABEL:
                    self.assertEqual(_col(table, col).totals_row_label, EXPECTED_LABEL[kind])

        # The same writes as one running sequence: every column stays consistent after every step.
        ws, table = _calc_table(self.wb, True)
        for i, (kind, col, _prior) in enumerate(TOTALS_STEPS):
            _write_totals(ws, col, kind)
            for c in range(3):
                self._assert_agrees(ws, table, c, f"sequence {i} column {c}")


class RoundTripTests(unittest.TestCase):
    def test_a_calculated_column_and_its_totals_survive_xlsx_and_xlsb(self) -> None:
        for fmt in (WorkbookFormat.XLSX, WorkbookFormat.XLSB):
            with self.subTest(fmt.name):
                with Workbook.create_default() as wb:
                    _, table = _calc_table(wb, True)
                    table.update(
                        columns=[
                            TableColumnPatch(totals_row_label="Total"),
                            TableColumnPatch(totals_row_function=Fn.SUM),
                            TableColumnPatch(calculated_formula="=A3*10"),
                        ]
                    )
                    data = wb.save(format=fmt)
                with Workbook.load(data) as reloaded:
                    ws = reloaded.sheet(0)
                    table = ws.table(0)
                    self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")
                    self.assertEqual(_col(table, 0).totals_row_label, "Total")
                    self.assertEqual(_col(table, 1).totals_row_function, Fn.SUM)
                    self.assertEqual([ws.get_formula(a1) for a1 in ("C3", "C4", "C5")], ["=A3*10", "=A4*10", "=A5*10"])
                    self.assertEqual(ws.get_formula("B6"), "=SUBTOTAL(109,[b])")
                    reloaded.recalc()
                    self.assertEqual(ws.get_value("C5"), 30.0)
                    self.assertEqual(ws.get_value("B6"), 60.0)
                    if fmt == WorkbookFormat.XLSX:
                        ws.set_number("C4", 7.0)
                        self.assertEqual(_col(table, 2).calculated_formula, "=A3*10")


if __name__ == "__main__":
    unittest.main()
