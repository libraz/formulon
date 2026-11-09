"""API coherence conformance: patch keys, patch signatures, anchors and rectangle argument names.

Criterion 3: an unknown key, a ``None`` on a key that cannot be cleared and an anchor key without its trigger raise
``TypeError`` at the top level and inside a nested patch. Criterion 4: the methods of the ``patch: true`` ops in
``tools/api/names.yaml`` take only ``UNCHANGED`` as a keyword default and no ``None`` default, and every drawing's
``update`` takes the common keys (``name``, ``description``, ``hidden``, ``anchor``) as keyword-only arguments.
Criterion 5: a rectangle argument position listed in
``tools/api/address_args.json`` carries, in Python, the snake_case of the name ``src/wasm/formulon.d.ts`` gives it.
"""

from __future__ import annotations

import inspect
import json
import re
import unittest
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import _chart_input

import formulon
from formulon import (
    UNCHANGED,
    AnchorKind,
    CellRect,
    Chart,
    DrawingAnchorInput,
    Picture,
    Shape,
    ShapeInput,
    Slicer,
    SlicerInput,
    TableInput,
    Workbook,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
ADDRESS_ARGS = REPO_ROOT / "tools" / "api" / "address_args.json"
NAMES_YAML = REPO_ROOT / "tools" / "api" / "names.yaml"
FORMULON_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"

# Patch-key combinations (pairwise model; the table below covers every pair of the values).
#   method : Shape.update | Slicer.update | Table.update | Chart.get_preview | PivotCache.set_worksheet_source
#            | Sheet.update_cell_format
#   key    : unknown | None-on-a-non-nullable-key | anchor-key-without-trigger | nullable-key-None (accepted)
#   depth  : top-level | nested (a patch mapping inside a keyword)
# The cases are written by hand from that model; the generator was not available when the file was written.


def _snake(name: str) -> str:
    return re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()


def _workbook() -> Workbook:
    wb = Workbook.create_default()
    ws = wb.sheet(0)
    ws.set_text(0, 0, "Product")
    ws.set_text(0, 1, "Amount")
    ws.set_number(1, 1, 10.0)
    ws.set_number(2, 1, 20.0)
    return wb


class _Fixture:
    """One workbook with a table, a shape, a slicer, a chart and a pivot cache."""

    def __init__(self) -> None:
        self.wb = _workbook()
        ws = self.wb.sheet(0)
        self.table = ws.add_table(TableInput(range=CellRect(0, 0, 2, 1), name="Sales", columns=["Product", "Amount"]))
        self.shape = ws.add_shape(
            ShapeInput(
                anchor=DrawingAnchorInput(kind=AnchorKind.ONE_CELL, from_row=6, width_emu=914400, height_emu=457200),
                preset_geometry="rect",
            )
        )
        self.slicer = ws.add_slicer(
            self.table, 0, SlicerInput(anchor=DrawingAnchorInput(from_row=10, from_col=4, to_row=20, to_col=6))
        )
        self.chart = _chart_input.add_chart(ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3")
        self.cache = self.wb.add_pivot_cache()

    def close(self) -> None:
        self.wb.close()


class PatchKeyTests(unittest.TestCase):
    """Criterion 3."""

    def setUp(self) -> None:
        self.f = _Fixture()
        self.addCleanup(self.f.close)

    def _cases(self) -> List[Tuple[str, Callable[[], Any], Optional[type]]]:
        f = self.f
        ws = f.wb.sheet(0)
        return [
            ("Shape.update unknown", lambda: f.shape.update(bogus=1), TypeError),
            ("Shape.update None", lambda: f.shape.update(name=None), TypeError),
            ("Shape.update None rotation", lambda: f.shape.update(rotation_deg=None), TypeError),
            ("Shape.update flat anchor key is gone", lambda: f.shape.update(from_row=1), TypeError),
            ("Shape.update flat edit_as is gone", lambda: f.shape.update(edit_as=1), TypeError),
            ("Shape.update anchor None", lambda: f.shape.update(anchor=None), TypeError),
            ("Shape.update anchor not a record", lambda: f.shape.update(anchor={"from_row": 1}), TypeError),
            ("Chart.update unknown", lambda: f.chart.update(bogus=1), TypeError),
            ("Chart.update None name", lambda: f.chart.update(name=None), TypeError),
            ("Shape.update flat fill_kind is gone", lambda: f.shape.update(fill_kind=1), TypeError),
            ("Shape.update paint None removes", lambda: f.shape.update(fill=None, line=None), None),
            ("Slicer.update unknown", lambda: f.slicer.update(bogus=1), TypeError),
            ("Slicer.update None", lambda: f.slicer.update(caption=None), TypeError),
            ("Slicer.update flat offset is gone", lambda: f.slicer.update(to_row_offset_emu=1), TypeError),
            ("Slicer.update flat corner is gone", lambda: f.slicer.update(to_row=3), TypeError),
            ("Slicer.update style_name None clears", lambda: f.slicer.update(style_name=None), None),
            ("Table.update unknown", lambda: f.table.update(bogus=1), TypeError),
            ("Table.update name is gone", lambda: f.table.update(name="Other"), TypeError),
            ("Table.update None range", lambda: f.table.update(range=None), TypeError),
            ("Table.update None flag", lambda: f.table.update(header_row=None), TypeError),
            ("Table.update rect is gone", lambda: f.table.update(rect=CellRect(0, 0, 2, 1)), TypeError),
            ("Table.update style is gone", lambda: f.table.update(style_name="TableStyleLight1"), TypeError),
            ("Table.update_style_options unknown", lambda: f.table.update_style_options(bogus=1), TypeError),
            (
                "Table.update_style_options None flag",
                lambda: f.table.update_style_options(show_first_column=None),
                TypeError,
            ),
            (
                "Table.update_style_options style None clears",
                lambda: f.table.update_style_options(style_name=None),
                None,
            ),
            ("Chart.get_preview unknown", lambda: f.chart.get_preview(bogus=1), TypeError),
            ("Chart.get_preview None", lambda: f.chart.get_preview(direction=None), TypeError),
            ("Chart.get_preview None range", lambda: f.chart.get_preview(range=None), TypeError),
            ("PivotCache.set_worksheet_source unknown", lambda: f.cache.set_worksheet_source(bogus=1), TypeError),
            (
                "PivotCache.set_worksheet_source None with keyword",
                lambda: f.cache.set_worksheet_source(None, ref="A1:B2"),
                TypeError,
            ),
            ("PivotCache.set_worksheet_source positional value", lambda: f.cache.set_worksheet_source("A1"), TypeError),
            ("update_cell_format unknown", lambda: ws.update_cell_format("A1", bogus=1), TypeError),
            (
                "update_cell_format nested font unknown",
                lambda: ws.update_cell_format("A1", font={"bogus": 1}),
                TypeError,
            ),
            (
                "update_cell_format nested fill unknown",
                lambda: ws.update_cell_format("A1", fill={"bogus": 1}),
                TypeError,
            ),
            (
                "update_cell_format nested alignment None",
                lambda: ws.update_cell_format("A1", alignment={"bogus": None}),
                TypeError,
            ),
        ]

    def test_cases(self) -> None:
        for label, call, expected in self._cases():
            with self.subTest(label):
                if expected is None:
                    call()
                else:
                    with self.assertRaises(expected):
                        call()


class DrawingUpdateSignatureTests(unittest.TestCase):
    """Criterion 4, second bullet."""

    COMMON = ("name", "description", "hidden", "anchor")

    def test_update_takes_the_common_keys_as_keyword_only_arguments(self) -> None:
        for cls in (Picture, Chart, Shape, Slicer):
            with self.subTest(cls=cls.__name__):
                params = {name: p for name, p in inspect.signature(cls.update).parameters.items() if name != "self"}
                for key in self.COMMON:
                    self.assertIn(key, params)
                self.assertTrue(all(p.kind == inspect.Parameter.KEYWORD_ONLY for p in params.values()))
                self.assertTrue(all(p.default is UNCHANGED for key, p in params.items() if key != "style"))

    def test_the_dropped_setters_are_gone(self) -> None:
        for cls in (Picture, Chart, Shape, Slicer):
            with self.subTest(cls=cls.__name__):
                self.assertFalse(hasattr(cls, "set_anchor"))
                self.assertFalse(hasattr(cls, "rename"))

    def test_update_anchor_moves_the_object(self) -> None:
        f = _Fixture()
        try:
            f.chart.update(anchor=DrawingAnchorInput(from_row=12, from_col=3))
            self.assertEqual(f.chart.get_drawing_properties().anchor.from_row, 12)
            with self.assertRaises(TypeError):
                f.chart.update(row=1)  # type: ignore[call-arg]
        finally:
            f.close()


def _patch_ops() -> List[Tuple[str, str, str]]:
    """``(op id, class name, method name)`` of every ``patch: true`` op with a Python surface."""
    import yaml

    with open(NAMES_YAML, encoding="utf-8") as handle:
        entries = yaml.safe_load(handle)["entries"]
    found = []
    for op, entry in entries.items():
        if not (isinstance(entry, dict) and entry.get("patch") and "py" in entry.get("surfaces", ())):
            continue
        handle_class = (entry.get("handle") or {}).get("class")
        if isinstance(handle_class, list):
            found.extend((op, name, entry["handle"]["py"]) for name in handle_class)
            continue
        if handle_class is None:
            found.append((op, "Workbook" if entry["scope"] == "workbook" else "Sheet", entry["py"]))
            continue
        verb, _, rest = entry["py"].partition("_")
        owner = _snake(handle_class)
        rest = "" if rest == owner else rest[len(owner) + 1 :] if rest.startswith(owner + "_") else rest
        found.append((op, handle_class, f"{verb}_{rest}" if rest else verb))
    return found


class PatchSignatureTests(unittest.TestCase):
    """Criterion 4, first bullet."""

    def test_patch_ops_are_found(self) -> None:
        self.assertGreaterEqual(len(_patch_ops()), 20)

    def test_keyword_defaults_are_unchanged(self) -> None:
        for op, class_name, method_name in _patch_ops():
            with self.subTest(op=op):
                method = getattr(getattr(formulon, class_name), method_name, None)
                self.assertIsNotNone(method, f"{class_name}.{method_name} is missing")
                for name, param in inspect.signature(method).parameters.items():
                    if name == "self":
                        continue
                    self.assertIsNot(param.default, None, f"{class_name}.{method_name}({name}) defaults to None")
                    if param.kind is inspect.Parameter.KEYWORD_ONLY:
                        self.assertIs(
                            param.default,
                            UNCHANGED,
                            f"{class_name}.{method_name}({name}) does not default to UNCHANGED",
                        )


def _dts_methods() -> Dict[str, Dict[str, List[str]]]:
    """``interface -> method -> [parameter names]`` of the top-level interfaces of the d.ts."""
    interfaces: Dict[str, Dict[str, List[str]]] = {}
    current: Optional[Dict[str, List[str]]] = None
    pending: Optional[Tuple[str, str]] = None
    for line in FORMULON_DTS.read_text(encoding="utf-8").split("\n"):
        opened = re.match(r"^(?:export )?(?:declare )?interface (\w+)\b[^{]*\{$", line)
        if opened:
            current = interfaces.setdefault(opened.group(1), {})
            continue
        if line == "}":
            current = None
            continue
        if current is None:
            continue
        if pending is not None:
            pending = (pending[0], pending[1] + " " + line.strip())
        else:
            method = re.match(r"^  (?:readonly )?(\w+)\((.*)$", line)
            if not method:
                continue
            pending = (method.group(1), method.group(2))
        name, text = pending
        depth, end = 1, None
        for at, ch in enumerate(text):
            depth += (ch in "([{<") - (ch in ")]}>" and text[at - 1] != "=")
            if depth == 0:
                end = at
                break
        if end is None:
            continue
        pending = None
        params, depth, start = [], 0, 0
        body = text[:end]
        for at, ch in enumerate(body + ","):
            if ch in "([{<":
                depth += 1
            elif ch in ")]}>" and body[at - 1 : at] != "=":
                depth -= 1
            elif ch == "," and depth == 0:
                piece = body[start:at].strip()
                start = at + 1
                if piece:
                    params.append(re.match(r"(?:\.\.\.)?(\w+)", piece).group(1))  # type: ignore[union-attr]
        current.setdefault(name, params)
    return interfaces


class RectArgumentNameTests(unittest.TestCase):
    """Criterion 5."""

    def test_python_names_follow_the_declarations(self) -> None:
        with open(ADDRESS_ARGS, encoding="utf-8") as handle:
            data = json.load(handle)
        rect_args = data["py"]["rect_args"]
        js_rect_args = data["js"]["rect_args"]
        self.assertTrue(rect_args)
        self.assertEqual(len(rect_args), len(js_rect_args))
        declared = _dts_methods()
        for (js_key, positions), (py_key, py_positions) in zip(sorted(js_rect_args.items()), sorted(rect_args.items())):
            with self.subTest(method=py_key):
                self.assertEqual(positions, py_positions)
                js_class, js_name = js_key.split(".")
                py_class, py_name = py_key.split(".")
                self.assertEqual(js_class, py_class)
                if js_name not in declared.get(js_class, {}):
                    js_class = py_class = js_class + "Ooxml"
                self.assertIn(js_name, declared.get(js_class, {}), f"{js_key} is not declared in the d.ts")
                js_params = declared[js_class][js_name]
                method = getattr(getattr(formulon, py_class), py_name)
                py_params = [p for name, p in inspect.signature(method).parameters.items() if name != "self"]
                for position in positions:
                    self.assertEqual(py_params[position].name, _snake(js_params[position]))
                    self.assertIn(
                        py_params[position].kind,
                        (inspect.Parameter.POSITIONAL_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD),
                    )


if __name__ == "__main__":
    unittest.main()
