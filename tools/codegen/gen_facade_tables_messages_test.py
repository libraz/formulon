"""Tests for the tables the argument-error messages and the read-only keys read: `PARAM_NAMES`, `BOOLEAN_ARGS`,
`C_FIELD_NAMES`, `READ_ONLY_KEYS`, `binding_record_params.inc`, `binding_param_names.inc`, the `METHOD_ARITY` rows
of the statics and module functions, and the `single` span form.

Run: `python3 -m unittest tools/codegen/gen_facade_tables_messages_test.py`
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import facade_scope  # noqa: E402
import gen_facade_tables as gen  # noqa: E402

DTS = """
export interface CellAddress { row: number; col: number; }
export interface CivilInput { year: number; month: number; day: number; }
export interface DataTableInputCell { readonly row: number; readonly col: number; readonly deleted: boolean; }
export interface DataTable { readonly rowInput: DataTableInputCell | null; readonly twoDimensional: boolean; }
export interface DataTableInput { range: string; rowInput?: CellAddress | string | null; }
export interface PageBreakInput { first?: number; last?: number; manual?: boolean; }
export interface SheetView { readonly zoomPct: number; readonly splitXTwips: number; }
export interface SheetViewPatch { zoomPct?: number; }
export interface Workbook {
  setDate1904(enabled: boolean): void;
  setTitle(title: string | null): void;
  setCalcMode(mode: number, sticky?: boolean | null): void;
}
export interface WorkbookOoxml { getFont(index: number): number; }
export interface Sheet {
  addDataTable(spec: DataTableInput): void;
  setRowBreak(row: number, input: PageBreakInput | null): void;
  updateView(patch: SheetViewPatch): void;
  setRowHidden(first: number, last: number, hidden: boolean): void;
  setRowHidden(span: string, hidden: boolean): void;
  getColumnWidthPt(col: number, mode?: number): number;
  getColumnWidthPt(a1: string, mode?: number): number;
}
export interface SheetOoxml { getColumnXfIndex(col: number): number | null; getColumnXfIndex(a1: string): number | null; }
export interface Drawing { update(patch: SheetViewPatch): void; }
export interface Picture extends Drawing { readonly kind: number; }
export interface WorkbookCtor { createDefault(): Workbook; load(bytes: Uint8Array): Workbook; }
export interface Formulon {
  readonly Workbook: WorkbookCtor;
  evaluateFormula(formula: string): number;
  convertCivilToSerial(civil: CivilInput, options?: { date1904?: boolean }): number;
}
"""


def _op(js: str, scope: str, **fields: object) -> dict:
    return {
        "kind": "op",
        "scope": scope,
        "js": js,
        "py": gen._snake(js),
        "surfaces": ["wasm", "native", "py"],
        **fields,
    }


OPS = {
    "date1904.set": _op("setDate1904", "workbook"),
    "data_table.add": _op("addDataTable", "sheet"),
    "row_break.set": _op("setRowBreak", "sheet"),
    "view.update": _op("updateView", "sheet", patch=True, patch_record="SheetView"),
    "row_hidden.set": _op("setRowHidden", "sheet", span_args=[[0, "row", "span"]]),
    "column_width_pt.get": _op("getColumnWidthPt", "sheet", span_args=[[0, "col", "single"]]),
    "record.data_table_spec": {
        "kind": "record",
        "role": "input",
        "js": "DataTableInput",
        "py": "DataTableInput",
        "surfaces": ["wasm", "native", "py"],
        "read_only_keys": ["two_dimensional", "row_input.deleted"],
    },
    "record.page_break_input": {
        "kind": "record",
        "role": "input",
        "js": "PageBreakInput",
        "py": "PageBreakInput",
        "surfaces": ["wasm", "native", "py"],
        "read_only_keys": ["row", "col"],
    },
    "record.sheet_view_patch": {
        "kind": "record",
        "role": "patch",
        "js": "SheetViewPatch",
        "surfaces": ["wasm", "native"],
        "read_only_keys": ["split_x_twips"],
    },
    "field.sheet_view_split_x_twips": {
        "kind": "field",
        "record": "record.sheet_view",
        "c": "split_x_twips",
        "js": "splitXTwips",
        "py": "split_x_twips",
        "surfaces": ["c", "wasm", "native", "py"],
    },
    "field.sparkline_group_items": {
        "kind": "field",
        "record": "record.sparkline_group",
        "c": "items",
        "js": "sparklines",
        "py": "sparklines",
        "surfaces": ["c", "wasm", "native", "py"],
    },
    "field.page_margins_has_left_in": {
        "kind": "field",
        "record": "record.page_margins",
        "c": "has_left_in",
        "js": "leftInStated",
        "surfaces": ["c", "wasm", "native"],
    },
}


def _ledger(entries: dict) -> facade_scope.Ledger:
    return facade_scope.Ledger({"schema": 1, "collections": {}, "entries": entries})


class ReadOnlyKeysTest(unittest.TestCase):
    def test_input_keys_take_the_record_keys_and_its_read_only_keys(self) -> None:
        top, nested = gen.input_keys(_ledger(OPS), DTS)
        self.assertEqual(top["Sheet.setRowBreak"], [1, ["col", "first", "last", "manual", "row"]])
        self.assertEqual(top["Sheet.addDataTable"], [0, ["range", "rowInput", "twoDimensional"]])
        # `rowInput` also takes an A1 string, so only the dotted read-only key opens a path into it.
        self.assertEqual(nested, {"Sheet.addDataTable": {"rowInput": ["col", "deleted", "row"]}})

    def test_patch_keys_take_the_patch_keys_and_its_read_only_keys(self) -> None:
        keys = facade_scope.patch_keys(_ledger(OPS), DTS)
        self.assertEqual(keys, {"Sheet.updateView": [0, ["splitXTwips", "zoomPct"]]})

    def test_a_dotted_path_into_no_nested_record_is_refused(self) -> None:
        entries = {k: dict(v) for k, v in OPS.items()}
        entries["record.data_table_spec"]["read_only_keys"] = ["range.deleted"]
        with self.assertRaisesRegex(gen.GenError, "read_only_keys path range names no nested record"):
            gen.input_keys(_ledger(entries), DTS)

    def test_js_and_py_tables(self) -> None:
        self.assertEqual(
            gen.read_only_table("js", _ledger(OPS), DTS),
            {
                "Sheet.addDataTable": ["rowInput.deleted", "twoDimensional"],
                "Sheet.setRowBreak": ["col", "row"],
                "Sheet.updateView": ["splitXTwips"],
            },
        )
        self.assertEqual(
            gen.read_only_table("py", _ledger(OPS), DTS),
            {
                "Sheet.add_data_table": ["row_input.deleted", "two_dimensional"],
                "Sheet.set_row_break": ["col", "row"],
                "Sheet.update_view": ["split_x_twips"],
            },
        )

    def test_js_region(self) -> None:
        text = gen.input_keys_region(_ledger(OPS), DTS)
        self.assertIn(
            "const READ_ONLY_KEYS = Object.freeze({\n"
            "  'Sheet.addDataTable': ['rowInput.deleted', 'twoDimensional'],\n"
            "  'Sheet.setRowBreak': ['col', 'row'],\n",
            text,
        )

    def test_a_long_python_row_takes_one_key_per_line(self) -> None:
        row = gen._py_tuple_row("Sheet.update_page_setup", [f"key_{i:02}_stated" for i in range(8)])
        self.assertTrue(row.startswith('    "Sheet.update_page_setup": (\n        "key_00_stated",\n'))
        self.assertTrue(all(len(line) <= 120 for line in row.splitlines()))
        self.assertEqual(gen._py_tuple_row("Sheet.x", ["a"]), '    "Sheet.x": ("a",),\n')


class MessageTablesTest(unittest.TestCase):
    def test_param_names_are_keyed_by_the_called_class(self) -> None:
        names = gen.param_names(_ledger(OPS), DTS)
        self.assertEqual(names["Sheet.updateView"], ["patch"])
        self.assertEqual(names["Sheet.setRowHidden"], ["first", "last", "hidden"])
        self.assertEqual(names["SheetOoxml.getColumnXfIndex"], ["col"])
        self.assertEqual(names["WorkbookOoxml.getFont"], ["index"])
        self.assertEqual(names["Workbook.load"], ["bytes"])
        self.assertEqual(names["Formulon.evaluateFormula"], ["formula"])
        self.assertNotIn("Workbook.createDefault", names)

    def test_param_names_cover_the_handle_classes_and_their_inherited_methods(self) -> None:
        handles = {
            "schema": 1,
            "collections": {
                "drawing": {"key": "natural", "access": "handles", "handle_class": "Drawing", "plural": "drawings"},
                "image": {
                    "key": "natural",
                    "access": "handles",
                    "handle_class": "Picture",
                    "handle_base": "Drawing",
                    "plural": "pictures",
                },
            },
            "entries": {
                "drawing.update": {
                    "kind": "op",
                    "scope": "sheet",
                    "js": "updateDrawing",
                    "surfaces": ["wasm"],
                    "handle": {"class": ["Drawing", "Picture"], "js": "update", "supplies": ["sheet", "object_id"]},
                },
            },
        }
        names = gen.param_names(facade_scope.Ledger(handles), DTS)
        self.assertEqual(names["Drawing.update"], ["patch"])
        self.assertEqual(names["Picture.update"], ["patch"])

    def test_boolean_args_are_the_plain_boolean_parameters(self) -> None:
        self.assertEqual(
            gen.boolean_args(_ledger(OPS), DTS),
            {"Sheet.setRowHidden": [2], "Workbook.setDate1904": [0]},
        )

    def test_c_field_names_are_the_fields_spelled_otherwise(self) -> None:
        self.assertEqual(
            gen.c_field_names(_ledger(OPS)),
            {"has_left_in": {"js": "leftInStated", "py": None}, "items": {"js": "sparklines", "py": "sparklines"}},
        )

    def test_a_c_field_published_under_two_names_is_refused(self) -> None:
        entries = {k: dict(v) for k, v in OPS.items()}
        entries["field.slicer_items"] = {**entries["field.sparkline_group_items"], "js": "slicers", "py": "slicers"}
        with self.assertRaisesRegex(gen.GenError, "C field items"):
            gen.c_field_names(_ledger(entries))

    def test_js_regions(self) -> None:
        self.assertIn("  'Sheet.updateView': ['patch'],\n", gen.param_names_region(_ledger(OPS), DTS))
        self.assertIn("  'Workbook.setDate1904': [0],\n", gen.boolean_args_region(_ledger(OPS), DTS))
        self.assertEqual(
            gen.c_field_names_region(_ledger(OPS)),
            "const C_FIELD_NAMES = Object.freeze({\n"
            "  has_left_in: { js: 'leftInStated', py: null },\n"
            "  items: { js: 'sparklines', py: 'sparklines' },\n"
            "});\n",
        )

    def test_record_params_name_the_record_parameter_of_each_call(self) -> None:
        self.assertEqual(
            gen.record_params(_ledger(OPS), DTS),
            {
                "Formulon.convertCivilToSerial": "civil",
                "Sheet.addDataTable": "spec",
                "Sheet.setRowBreak": "input",
                "Sheet.updateView": "patch",
            },
        )
        inc = gen.record_params_inc(_ledger(OPS), DTS)
        self.assertIn("constexpr RecordParam kRecordParams[] = {\n", inc)
        self.assertIn('    {"Sheet.updateView", "patch"},\n};\n', inc)

    def test_only_record_params_writes_the_inc_alone(self) -> None:
        self.assertEqual([path for path, _ in gen.outputs("record-params")], [gen.RECORD_PARAMS_INC])

    def test_param_names_inc_spells_each_call_s_names_space_separated(self) -> None:
        inc = gen.param_names_inc(_ledger(OPS), DTS, {})
        self.assertIn("constexpr ParamNames kParamNames[] = {\n", inc)
        self.assertIn('    {"Sheet.setRowHidden", "first last hidden"},\n', inc)
        self.assertIn('    {"Formulon.convertCivilToSerial", "civil options"},\n', inc)
        self.assertNotIn('"Workbook.createDefault"', inc)

    def test_a_raw_layout_that_is_not_the_signature_loses_its_names(self) -> None:
        layouts = {
            # The facade supplies the sheet index; row, col and value line up.
            "Sheet.setRowHidden": [(1, {"u32": 0x7, "required": 0x8})],
            # A number where the d.ts takes a record: the raw layout is not the signature.
            "Sheet.setRowBreak": [(1, {"u32": 0x6})],
            # More checked positions than the d.ts declares.
            "Sheet.updateView": [(1, {"u32": 0x1, "object": 0x6})],
        }
        names = gen.binding_param_names(_ledger(OPS), DTS, layouts)
        self.assertEqual(names["Sheet.setRowHidden"], ["first", "last", "hidden"])
        self.assertNotIn("Sheet.setRowBreak", names)
        self.assertNotIn("Sheet.updateView", names)

    def test_an_object_union_parameter_roots_its_record(self) -> None:
        dts = (
            DTS.replace("convertCivilToSerial(civil: CivilInput,", "convertCivilToSerial(date: DateInput,")
            + "export type DateInput = CivilInput | Date;\nexport type BytesInput = Uint8Array | ArrayBuffer;\n"
        )
        params = gen.record_params(_ledger(OPS), dts)
        self.assertEqual(params["Formulon.convertCivilToSerial"], "date")

    def test_only_param_names_writes_the_inc_alone(self) -> None:
        self.assertEqual([path for path, _ in gen.outputs("param-names")], [gen.PARAM_NAMES_INC])


class SignatureRowsTest(unittest.TestCase):
    def test_arity_covers_the_statics_and_the_module_functions(self) -> None:
        arity, _typed = gen.method_signatures(_ledger(OPS), DTS)
        self.assertEqual(arity["Workbook.createDefault"], 0)
        self.assertEqual(arity["Workbook.load"], 1)
        self.assertEqual(arity["Formulon.evaluateFormula"], 1)
        self.assertEqual(arity["Formulon.convertCivilToSerial"], 2)
        self.assertNotIn("Formulon.Workbook", arity)

    def test_a_static_named_like_an_instance_method_is_refused(self) -> None:
        dts = DTS.replace("load(bytes: Uint8Array): Workbook;", "setDate1904(enabled: boolean): Workbook;")
        with self.assertRaisesRegex(gen.GenError, "Workbook.setDate1904 is both"):
            gen.method_signatures(_ledger(OPS), dts)

    def test_a_single_span_keeps_its_arity_and_its_coordinate_role(self) -> None:
        arity, _typed = gen.method_signatures(_ledger(OPS), DTS)
        self.assertEqual(arity["Sheet.getColumnWidthPt"], 2)
        header = (
            "FM_API fm_status_t fm_sheet_get_column_width_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t col,"
            " int32_t mode, double* out_pt);\n"
        )
        entries = {k: dict(v) for k, v in OPS.items()}
        entries["column_width_pt.get"]["c"] = "fm_sheet_get_column_width_pt"
        # The A1 overload has as many parameters; the coordinate form, declared first, still names them.
        roles = gen.arg_roles(_ledger(entries), DTS, header)
        self.assertEqual(roles, {"Sheet.getColumnWidthPt": [[0, "coordinate"], [1, "quantity"]]})

    def test_span_args_carry_the_single_form(self) -> None:
        self.assertEqual(
            gen.address_tables("js", _ledger(OPS)).split("const SPAN_ARGS")[1].split("});")[0],
            " = Object.freeze({\n"
            "  'Sheet.getColumnWidthPt': [[0, 'col', 'single']],\n"
            "  'Sheet.setRowHidden': [[0, 'row', 'span']],\n",
        )


class RegionGateTest(unittest.TestCase):
    def test_several_empty_regions_are_skipped_together(self) -> None:
        current = "a\n// <param-names>\n// </param-names>\n// <boolean-args>\n// </boolean-args>\nb\n"
        filled = gen._splice(current, gen._PARAM_NAMES_REGION, "const P = 1;\n", Path("common.mjs"))
        filled = gen._splice(filled, gen._BOOLEAN_ARGS_REGION, "const B = 1;\n", Path("common.mjs"))
        self.assertTrue(gen._skippable(gen.NPM_COMMON, current, filled))
        self.assertFalse(gen._skippable(gen.NPM_COMMON, current.replace("a\n", "c\n"), filled))


if __name__ == "__main__":
    unittest.main()
