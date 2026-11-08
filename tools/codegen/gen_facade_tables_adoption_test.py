"""Tests for the `.ooxml` namespace tables, `RECORD_RANGE_FIELDS` and the ledger keys they and the adoption
rules read (`reads_values`, `resolved`, `color`, `namespace`, `a1_exempt`, `params` defaults).

Run: `python3 -m unittest tools/codegen/gen_facade_tables_adoption_test.py`
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "api"))

import facade_scope  # noqa: E402
import gen_facade_tables as gen  # noqa: E402
import validate_names  # noqa: E402

DTS = """
export interface CellRect { row: number; col: number; rowCount: number; colCount: number; }
export interface SheetRect { readonly sheet: number; readonly rect: CellRect; }
export interface CellAddress { row: number; col: number; }
export interface SortInput { ref: CellRect | string; }
export interface FilterInput {
  range: CellRect;
  sort?: SortInput | null;
  ranges: CellRect | string | readonly (CellRect | string)[];
  spans: readonly SheetRect[];
  anchor: CellAddress;
  title: string;
}
export interface ShapeInput { fromRow?: number; fromCol?: number; frame?: CellRect; }
export interface PlainInput { name: string; }
export interface FilterResult { readonly range: CellRect; }
"""


def _record(js: str, role: str = "input", py: str = "", **extra) -> dict:
    entry = {"kind": "record", "role": role, "js": js, "surfaces": ["wasm", "native"], **extra}
    if py:
        entry["py"] = py
        entry["surfaces"] = ["wasm", "native", "py"]
    return entry


def _ledger(entries: dict) -> facade_scope.Ledger:
    return facade_scope.Ledger({"schema": 1, "collections": {}, "entries": entries})


RECORDS = {
    "record.sort_input": _record("SortInput", py="SortInput"),
    "record.filter_input": _record("FilterInput", py="FilterInput"),
    "record.shape_input": _record(
        "ShapeInput", a1_exempt={"fields": ["fromRow", "fromCol", "frame"], "reason": "an EMU-offset anchor"}
    ),
    "record.plain_input": _record("PlainInput"),
    "record.filter_result": _record("FilterResult", role="result"),
}


class RangeTypeTest(unittest.TestCase):
    def test_scalar_array_and_single_or_array(self) -> None:
        known = gen.RANGE_TYPES
        self.assertEqual(gen._range_type("CellRect", known), "CellRect")
        self.assertEqual(gen._range_type("CellRect | string | null", known), "CellRect")
        self.assertEqual(gen._range_type("readonly CellRect[]", known), "CellRect[]")
        self.assertEqual(gen._range_type("CellRect | string | readonly (CellRect | string)[]", known), "CellRect[]")
        self.assertEqual(gen._range_type("ReadonlyArray<SheetRect>", known), "SheetRect[]")

    def test_other_types_are_not_ranges(self) -> None:
        known = gen.RANGE_TYPES
        self.assertIsNone(gen._range_type("string", known))
        self.assertIsNone(gen._range_type("number", known))
        self.assertIsNone(gen._range_type("CellRect | SheetRect", known))
        self.assertIsNone(gen._range_type("Foo", known))


class RecordRangeFieldsTest(unittest.TestCase):
    def test_input_records_nested_and_exempt(self) -> None:
        table = gen.record_range_fields(_ledger(RECORDS), DTS)
        self.assertEqual(
            table,
            {
                "FilterInput": {
                    "range": "CellRect",
                    "sort": "SortInput",
                    "ranges": "CellRect[]",
                    "spans": "SheetRect[]",
                    "anchor": "CellAddress",
                },
                "SortInput": {"ref": "CellRect"},
            },
        )

    def test_js_region(self) -> None:
        text = gen.range_fields_region("js", _ledger(RECORDS), DTS)
        self.assertEqual(
            text,
            "const RECORD_RANGE_FIELDS = Object.freeze({\n"
            "  FilterInput: {\n"
            "    range: 'CellRect',\n"
            "    sort: 'SortInput',\n"
            "    ranges: 'CellRect[]',\n"
            "    spans: 'SheetRect[]',\n"
            "    anchor: 'CellAddress',\n"
            "  },\n"
            "  SortInput: { ref: 'CellRect' },\n"
            "});\n",
        )

    def test_py_region_uses_python_names(self) -> None:
        text = gen.range_fields_region("py", _ledger(RECORDS), DTS)
        self.assertIn('"SortInput": {"ref": "CellRect"},', text)
        self.assertIn('"FilterInput": {\n', text)
        self.assertIn('        "sort": "SortInput",\n', text)

    def test_a_discriminated_union_has_every_alternatives_fields(self) -> None:
        dts = DTS + (
            "interface RuleBase { ranges: readonly (CellRect | string)[] | string; }\n"
            "interface ExprRuleInput extends RuleBase { type: 0; formula: string; }\n"
            "interface BarRuleInput extends RuleBase { type: 1; anchor: CellAddress | string; }\n"
            "export type RuleInput = ExprRuleInput | BarRuleInput;\n"
            "export type NotUnionInput = ExprRuleInput;\n"
        )
        ledger = _ledger({"record.rule_input": _record("RuleInput"), "record.not_union": _record("NotUnionInput")})
        self.assertEqual(
            gen.record_range_fields(ledger, dts), {"RuleInput": {"ranges": "CellRect[]", "anchor": "CellAddress"}}
        )

    def test_no_input_records_emit_an_empty_table(self) -> None:
        self.assertEqual(
            gen.range_fields_region("js", _ledger({}), DTS), "const RECORD_RANGE_FIELDS = Object.freeze({});\n"
        )


def _op(js: str, scope: str, **extra) -> dict:
    py = "".join("_" + c.lower() if c.isupper() else c for c in js)
    return {"kind": "op", "scope": scope, "js": js, "py": py, "surfaces": ["wasm", "native", "py"], **extra}


class OoxmlNamespaceTest(unittest.TestCase):
    ENTRIES = {
        "font.get": _op("getFont", "workbook", namespace="ooxml"),
        "default_font.set": _op("setDefaultFont", "workbook"),
        "cell_xf_index.get": _op("getCellXfIndex", "sheet", namespace="ooxml"),
        "value.get": _op("getValue", "sheet"),
    }

    def test_namespace_ops_leave_the_class_lists(self) -> None:
        names, _ = gen._tables("js", _ledger(self.ENTRIES))
        self.assertEqual(names["WORKBOOK_METHODS"], ["setDefaultFont"])
        self.assertEqual(names["SHEET_METHODS"], ["getValue"])
        self.assertEqual(names["OOXML_WORKBOOK_METHODS"], ["getFont"])
        self.assertEqual(names["OOXML_SHEET_METHODS"], ["getCellXfIndex"])
        py, _ = gen._tables("py", _ledger(self.ENTRIES))
        self.assertEqual(py["OOXML_SHEET_METHODS"], ["get_cell_xf_index"])

    def test_regions_carry_the_namespace_lists(self) -> None:
        js = gen.js_region(_ledger(self.ENTRIES))
        self.assertIn("const OOXML_WORKBOOK_METHODS = Object.freeze([\n  'getFont',\n]);\n", js)
        self.assertIn("const OOXML_SHEET_METHODS = Object.freeze([\n  'getCellXfIndex',\n]);\n", js)
        py = gen.py_region(_ledger(self.ENTRIES))
        self.assertIn('OOXML_WORKBOOK_METHODS = (\n    "get_font",\n)\n', py)

    def test_without_a_namespace_no_list_is_emitted(self) -> None:
        plain = {k: {f: v for f, v in e.items() if f != "namespace"} for k, e in self.ENTRIES.items()}
        self.assertNotIn("OOXML_", gen.js_region(_ledger(plain)))
        self.assertNotIn("OOXML_", gen.py_region(_ledger(plain)))


SIGNATURE_DTS = (
    DTS
    + """
export interface Workbook { listSheets(): Sheet[]; addFilter(input: FilterInput): void; }
export interface WorkbookOoxml { getFont(index: number): number; }
export interface Sheet {
  setNumber(row: number, col: number, value: number): void;
  setNumber(a1: string, value: number): void;
  setWidth(first: number, last: number, width: number): void;
  setWidth(span: string, width: number): void;
  setPrintArea(ranges: readonly (CellRect | string)[] | string | null): void;
  getValues(range: CellRect | string): number[][];
  sort(input: SortInput, options?: number): void;
}
export interface SheetOoxml { setXf(range: CellRect | string, xf: number): void; }
"""
)

SIGNATURE_OPS = {
    **RECORDS,
    "value.set": _op("setNumber", "sheet", cell_args=[[0, 1]]),
    "width.set": _op("setWidth", "sheet", span_args=[[0, "col", "span"]]),
}


class MethodSignaturesTest(unittest.TestCase):
    def test_arity_is_the_longest_overload_keyed_by_the_owning_class(self) -> None:
        arity, _typed = gen.method_signatures(_ledger(SIGNATURE_OPS), SIGNATURE_DTS)
        self.assertEqual(
            arity,
            {
                "Sheet.getValues": 1,
                "Sheet.setNumber": 3,
                "Sheet.setPrintArea": 1,
                "Sheet.setWidth": 3,
                "Sheet.setXf": 2,
                "Sheet.sort": 2,
                "Workbook.addFilter": 1,
                "Workbook.getFont": 1,
                "Workbook.listSheets": 0,
            },
        )

    def test_typed_args_name_records_and_range_arrays_but_not_a_plain_rect(self) -> None:
        _arity, typed = gen.method_signatures(_ledger(SIGNATURE_OPS), SIGNATURE_DTS)
        self.assertEqual(
            typed,
            {
                "Sheet.setPrintArea": [[0, "CellRect[]"]],
                "Sheet.sort": [[0, "SortInput"]],
                "Workbook.addFilter": [[0, "FilterInput"]],
            },
        )

    def test_a_cell_or_span_method_must_declare_its_string_form(self) -> None:
        dts = SIGNATURE_DTS.replace("  setNumber(a1: string, value: number): void;\n", "")
        with self.assertRaisesRegex(
            gen.GenError, r"Sheet\.setNumber: overloads take \[3\] arguments, expected \[2, 3\]"
        ):
            gen.method_signatures(_ledger(SIGNATURE_OPS), dts)

    def test_js_region(self) -> None:
        text = gen.signatures_region(_ledger(SIGNATURE_OPS), SIGNATURE_DTS)
        self.assertIn("const METHOD_ARITY = Object.freeze({\n  'Sheet.getValues': 1,\n", text)
        self.assertIn("  'Workbook.listSheets': 0,\n});\n", text)
        self.assertIn("const TYPED_ARGS = Object.freeze({\n  'Sheet.setPrintArea': [[0, 'CellRect[]']],\n", text)


class RegionGateTest(unittest.TestCase):
    PATTERN = gen._RANGE_REGION["js"]

    def test_empty_region_is_skipped_by_check_and_filled_by_write(self) -> None:
        current = "a\n// <record-range-fields>\n// </record-range-fields>\nb\n"
        filled = gen._splice(current, self.PATTERN, "const X = 1;\n", Path("common.mjs"))
        self.assertEqual(filled, "a\n// <record-range-fields>\nconst X = 1;\n// </record-range-fields>\nb\n")
        self.assertTrue(gen._skippable(gen.NPM_COMMON, current, filled))

    def test_a_stale_filled_region_is_not_skipped(self) -> None:
        current = "// <record-range-fields>\nconst X = 0;\n// </record-range-fields>\n"
        text = gen._splice(current, self.PATTERN, "const X = 1;\n", Path("common.mjs"))
        self.assertFalse(gen._skippable(gen.NPM_COMMON, current, text))

    def test_a_change_outside_the_empty_region_is_not_skipped(self) -> None:
        current = "a\n// <record-range-fields>\n// </record-range-fields>\n"
        text = "b\n// <record-range-fields>\nconst X = 1;\n// </record-range-fields>\n"
        self.assertFalse(gen._skippable(gen.NPM_COMMON, current, text))

    def test_today_the_outputs_are_current(self) -> None:
        for path, text in gen.outputs():
            if text is None:
                continue
            current = path.read_text(encoding="utf-8") if path.is_file() else None
            with self.subTest(path=str(path.relative_to(gen.REPO_ROOT))):
                self.assertTrue(current == text or gen._skippable(path, current, text))


def _schema(entry: dict) -> list:
    return validate_names._adoption_schema("e", entry)


class LedgerSchemaTest(unittest.TestCase):
    def test_reads_values_on_a_c_backed_op(self) -> None:
        self.assertEqual(_schema({"kind": "op", "c": "fm_sheet_get_value", "reads_values": True}), [])
        self.assertTrue(_schema({"kind": "op", "js": "getValue", "reads_values": True}))
        self.assertTrue(_schema({"kind": "op", "c": "fm_sheet_get_value", "reads_values": "yes"}))

    def test_resolved_and_color(self) -> None:
        self.assertEqual(_schema({"kind": "op", "resolved": True}), [])
        self.assertTrue(_schema({"kind": "record", "resolved": True}))
        self.assertEqual(_schema({"kind": "record", "color": "resolved"}), [])
        self.assertTrue(_schema({"kind": "record", "color": "rgb"}))

    def test_namespace(self) -> None:
        self.assertEqual(_schema({"kind": "op", "scope": "sheet", "namespace": "ooxml"}), [])
        self.assertTrue(_schema({"kind": "op", "scope": "sheet", "namespace": "advanced"}))
        self.assertTrue(_schema({"kind": "op", "scope": "process", "namespace": "ooxml"}))

    def test_a1_exempt(self) -> None:
        good = {"fields": ["fromRow", "fromCol"], "reason": "an EMU-offset anchor"}
        self.assertEqual(_schema({"kind": "record", "role": "input", "a1_exempt": good}), [])
        self.assertEqual(_schema({"kind": "typedef", "a1_exempt": good}), [])
        self.assertTrue(_schema({"kind": "record", "role": "result", "a1_exempt": good}))
        self.assertTrue(_schema({"kind": "record", "role": "input", "a1_exempt": {"fields": []}}))

    def test_param_defaults(self) -> None:
        check = validate_names._param_default_schema
        self.assertEqual(check("p", {"default": 1}), [])
        self.assertEqual(check("p", {"default": None}), [])
        self.assertEqual(check("p", {"type": "int", "default": 0}), [])
        self.assertEqual(check("p", {"default": "today"}), [])
        self.assertTrue(check("p", {"default": "tomorrow"}))
        self.assertTrue(check("p", {"type": "int"}))
        self.assertEqual(validate_names.slot_type({"type": "enum:x", "default": 0}), "enum:x")
        self.assertEqual(validate_names.slot_type("bool"), "bool")

    def test_the_ledger_validates(self) -> None:
        self.assertEqual(validate_names.main(["--ledger-only"]), 0)


if __name__ == "__main__":
    unittest.main()
