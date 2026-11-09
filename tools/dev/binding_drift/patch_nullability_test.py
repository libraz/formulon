#!/usr/bin/env python3
"""Fixture tests for the patch-nullability check.

Run from the repo root: `python3 -m tools.dev.binding_drift.patch_nullability_test`.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift.api_conventions import Ledger
from tools.dev.binding_drift.patch_nullability import check_patch_nullability, dts_null_keys

DTS = """
export type Patch<T, Clearable extends keyof T = never> = { [K in keyof T]?: K extends Clearable ? T[K] | null : T[K] };
/** doc */
export type WidgetPatch = Patch<
  { gapWidth: number; titleText: string },
  'titleText'
> & { extra?: string | null; plain?: number };
export type PlainPatch = Patch<Thing>;
export interface BoxPatch {
  readonly fill?: Fill | null;
  readonly size?: number;
}
"""

LEDGER = {
    "schema": 1,
    "entries": {
        "record.widget_patch": {
            "kind": "record",
            "role": "patch",
            "js": "WidgetPatch",
            "clearable": ["title_text", "extra"],
        },
        "widget.update": {
            "kind": "op",
            "scope": "sheet",
            "py": "update_widget",
            "patch": True,
            "patch_record": "Widget",
            "surfaces": ["py"],
        },
    },
}

WIDGET_DTS = (
    "export type WidgetPatch = Patch<{ gapWidth: number; titleText: string }, 'titleText'> & { extra?: string | null };"
)


def _package(source: str) -> TemporaryDirectory:
    tmp = TemporaryDirectory()
    Path(tmp.name, "_sheet.py").write_text(source, encoding="utf-8")
    return tmp


def _run(source: str, ledger: dict = LEDGER, dts: str = WIDGET_DTS) -> list:
    with _package(source) as tmp:
        return check_patch_nullability(Ledger(ledger), dts, Path(tmp))


GOOD = """
class Sheet:
    def update_widget(self, *, gap_width: Union[int, Unchanged] = U,
                      title_text: Union[str, None, Unchanged] = U,
                      extra: Optional[str] = None) -> None: ...
"""


class DtsNullKeysTest(unittest.TestCase):
    def test_patch_literals_and_brace_members_admit_null(self) -> None:
        self.assertEqual(dts_null_keys(DTS, "WidgetPatch"), {"titleText", "extra"})

    def test_a_patch_without_a_clearable_argument_admits_none(self) -> None:
        self.assertEqual(dts_null_keys(DTS, "PlainPatch"), set())

    def test_an_interface_admits_null_on_its_null_members(self) -> None:
        self.assertEqual(dts_null_keys(DTS, "BoxPatch"), {"fill"})

    def test_an_undeclared_type_is_none(self) -> None:
        self.assertIsNone(dts_null_keys(DTS, "MissingPatch"))


class PatchNullabilityTest(unittest.TestCase):
    def test_agreement_is_clean(self) -> None:
        self.assertEqual(_run(GOOD), [])

    def test_python_admitting_none_on_a_fixed_key_is_drift(self) -> None:
        problems = _run(GOOD.replace("gap_width: Union[int, Unchanged]", "gap_width: Union[int, None, Unchanged]"))
        self.assertEqual(len(problems), 1)
        self.assertIn("update_widget(gap_width) admits None", problems[0])

    def test_python_excluding_none_on_a_clearable_key_is_drift(self) -> None:
        problems = _run(GOOD.replace("title_text: Union[str, None, Unchanged]", "title_text: Union[str, Unchanged]"))
        self.assertEqual(len(problems), 1)
        self.assertIn("update_widget(title_text) excludes None", problems[0])

    def test_a_clearable_key_missing_from_python_is_drift(self) -> None:
        problems = _run(GOOD.replace("extra: Optional[str] = None", "other: int = 0"))
        self.assertTrue(any("no keyword `extra`" in p for p in problems))

    def test_dts_disagreeing_with_the_ledger_is_drift(self) -> None:
        problems = _run(
            GOOD, dts="export type WidgetPatch = Patch<{ gapWidth: number; titleText: string }, 'gapWidth'>;"
        )
        self.assertEqual(len(problems), 1)
        self.assertIn("d.ts WidgetPatch admits null on ['gap_width']", problems[0])

    def test_a_drawing_family_description_may_clear(self) -> None:
        ledger = {
            "schema": 1,
            "entries": {
                "drawing.update": {
                    "kind": "op",
                    "scope": "sheet",
                    "py": "update_drawing",
                    "patch": True,
                    "patch_record": "DrawingProperties",
                    "surfaces": ["py"],
                    "handle": {"class": "Drawing", "py": "update"},
                }
            },
        }
        source = "class Drawing:\n    def update(self, *, name: str = U, description: Union[str, None, Unchanged] = U): ...\n"
        self.assertEqual(_run(source, ledger, ""), [])
        self.assertEqual(len(_run(source.replace("name: str", "name: Optional[str]"), ledger, "")), 1)


if __name__ == "__main__":
    unittest.main()
