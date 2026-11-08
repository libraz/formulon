"""Tests for the `INPUT_KEYS`, `INPUT_NESTED_KEYS` and `PATCH_NULLABLE` tables.

Run: `python3 -m unittest tools/codegen/gen_facade_tables_input_test.py`
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
export interface CellRect { row: number; col: number; rowCount: number; colCount: number; }
export interface ColumnInput { name: string; width?: number; }
export interface BoxInput { title: string; columns: readonly ColumnInput[]; frame?: ColumnInput | null; }
export interface Box { readonly id: number; title: string; }
export type Patch<T, Clearable extends keyof T = never> = { [K in keyof T]?: K extends Clearable ? T[K] | null : T[K] };
export interface ShapePatch { name?: string; hidden?: boolean; style?: number | null; note?: string | null; }
export interface Workbook {
  addBox(input: BoxInput): number;
  setBox(index: number, input: Box): void;
  setRange(range: CellRect): void;
}
export interface Sheet { updateShape(index: number, patch: ShapePatch): void; }
"""

OPS = {
    "box.add": {"kind": "op", "scope": "workbook", "js": "addBox", "surfaces": ["wasm", "native"]},
    "box.set": {
        "kind": "op",
        "scope": "workbook",
        "js": "setBox",
        "surfaces": ["wasm", "native"],
        "patch_record": "Box",
    },
    "range.set": {"kind": "op", "scope": "workbook", "js": "setRange", "surfaces": ["wasm", "native"]},
    "shape.update": {
        "kind": "op",
        "scope": "sheet",
        "js": "updateShape",
        "patch": True,
        "surfaces": ["wasm", "native"],
    },
    "record.shape_patch": {
        "kind": "record",
        "role": "patch",
        "js": "ShapePatch",
        "surfaces": ["wasm", "native"],
        "clearable": ["style", "note"],
    },
}


def _ledger(entries: dict) -> facade_scope.Ledger:
    return facade_scope.Ledger({"schema": 1, "collections": {}, "entries": entries})


class InputKeysTest(unittest.TestCase):
    def test_input_record_keys_and_nested_paths(self) -> None:
        top, nested = gen.input_keys(_ledger(OPS), DTS)
        # `BoxInput` also accepts the `Box` a read returns, so the read-only `id` is a known key.
        self.assertEqual(
            top,
            {"Workbook.addBox": [0, ["columns", "frame", "id", "title"]], "Workbook.setBox": [1, ["id", "title"]]},
        )
        self.assertEqual(nested, {"Workbook.addBox": {"columns[]": ["name", "width"], "frame": ["name", "width"]}})

    def test_js_region(self) -> None:
        text = gen.input_keys_region(_ledger(OPS), DTS)
        self.assertIn(
            "const INPUT_KEYS = Object.freeze({\n  'Workbook.addBox': [0, ['columns', 'frame', 'id', 'title']],", text
        )
        self.assertIn(
            "  'Workbook.addBox': {\n    'columns[]': ['name', 'width'],\n    frame: ['name', 'width'],\n  },\n", text
        )

    def test_no_input_ops_emit_empty_tables(self) -> None:
        text = gen.input_keys_region(_ledger({}), DTS)
        self.assertEqual(
            text, "const INPUT_KEYS = Object.freeze({\n});\nconst INPUT_NESTED_KEYS = Object.freeze({\n});\n"
        )


class PatchNullableTest(unittest.TestCase):
    def test_clearable_fields_with_their_typeof_name(self) -> None:
        self.assertEqual(
            gen.patch_nullable(_ledger(OPS), DTS), {"Sheet.updateShape": {"note": "string", "style": "number"}}
        )

    def test_js_region(self) -> None:
        self.assertEqual(
            gen.patch_nullable_region(_ledger(OPS), DTS),
            "const PATCH_NULLABLE = Object.freeze({\n  'Sheet.updateShape': { note: 'string', style: 'number' },\n});\n",
        )

    def test_absent_or_empty_clearable_has_no_entry(self) -> None:
        for clearable in ({}, {"clearable": []}, {"clearable": None}):
            entries = {k: dict(v) for k, v in OPS.items()}
            entries["record.shape_patch"].pop("clearable")
            entries["record.shape_patch"].update(clearable)
            self.assertEqual(gen.patch_nullable(_ledger(entries), DTS), {})

    def test_a_clearable_field_the_patch_lacks_is_refused(self) -> None:
        entries = {k: dict(v) for k, v in OPS.items()}
        entries["record.shape_patch"]["clearable"] = ["missing"]
        with self.assertRaises(ValueError):
            gen.patch_nullable(_ledger(entries), DTS)


if __name__ == "__main__":
    unittest.main()
