"""Getter -> setter round trips driven by ``tools/api/roundtrip_cases.json`` (api-conventions rule 26).

Every row that lists the ``py`` surface and is not listed for ``py`` in ``tools/api/roundtrip_pending.json`` is run:
the object is built by the row's ``create`` ops, the getter's record is passed back through the setter, and

* ``patch``: every record field the setter names as a keyword goes back as that keyword (fields the setter does not
  take, such as ``*_stated`` flags and object ids, are read-only) and a re-read equals the first read, so state a
  dropped field carried shows up as a difference;
* ``readd``: the record is added again (with the row's ``clear`` field set to ``None`` first when it has one) and the
  re-added copy equals the original apart from its identity fields.

A row the driver cannot map (an op or a method it has no recipe for) fails by name rather than being skipped.
"""

from __future__ import annotations

import dataclasses
import datetime
import inspect
import json
import unittest
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import _chart_input

from formulon import (
    AnchorKind,
    CellAddress,
    CellRect,
    CfCellIsRuleInput,
    CfExpressionRuleInput,
    CfTop10RuleInput,
    CommentInput,
    ComparisonOperator,
    DataValidationInput,
    DrawingAnchorInput,
    HyperlinkInput,
    Person,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldInput,
    PivotInput,
    PivotWorksheetSourceInput,
    ProtectedRangeInput,
    ShapeInput,
    SlicerInput,
    SparklineGroupInput,
    SparklineInput,
    SparklineType,
    TableInput,
    ThreadedCommentInput,
    ValidationType,
    Workbook,
)

API_DIR = Path(__file__).resolve().parents[3] / "tools" / "api"
SURFACE = "py"
PERSON = "{6B3C2E1A-0000-4000-8000-00000000A11C}"
THREAD = "{00000000-0000-4000-8000-000000000001}"

Context = Dict[str, Any]


def _fill_source(ws: Any) -> None:
    ws.set_text(0, 0, "Region")
    ws.set_text(0, 1, "Amount")
    ws.set_text(1, 0, "East")
    ws.set_number(1, 1, 10.0)
    ws.set_text(2, 0, "West")
    ws.set_number(2, 1, 30.0)


def _table(ctx: Context) -> None:
    _fill_source(ctx["ws"])
    ctx["Table"] = ctx["ws"].add_table(TableInput(range="A1:B3", name="Tbl", columns=["Region", "Amount"]))


def _pivot_cache(ctx: Context) -> None:
    _fill_source(ctx["ws"])
    cache = ctx["wb"].add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
    cache.add_field("Region")
    cache.add_field("Amount")
    for region, amount in (("East", 10.0), ("West", 30.0)):
        record = cache.add_record()
        cache.set_record_value(record, 0, region)
        cache.set_record_value(record, 1, amount)
    ctx["PivotCache"] = cache


def _pivot(ctx: Context) -> None:
    pivot = ctx["ws"].add_pivot(PivotInput(name="ByRegion", cache_id=ctx["PivotCache"].cache_id, anchor="E1"))
    pivot.set_field_axis("Region", PivotAxis.ROW)
    pivot.set_field_axis("Amount", PivotAxis.VALUE)
    pivot.add_data_field(PivotDataFieldInput(field_index=1, name="Sum of Amount", aggregation=PivotAggregation.SUM))
    ctx["Pivot"] = pivot


def _chart(ctx: Context) -> None:
    _fill_source(ctx["ws"])
    ctx["Chart"] = _chart_input.add_chart(ctx["ws"], _chart_input.COLUMN, "Sheet1!$A$1:$B$3")


def _shape(ctx: Context) -> None:
    anchor = DrawingAnchorInput(kind=AnchorKind.ONE_CELL, from_row=2, width_emu=9144, height_emu=9144)
    ctx["Shape"] = ctx["ws"].add_shape(ShapeInput(preset_geometry="rect", anchor=anchor))


def _slicer(ctx: Context) -> None:
    anchor = DrawingAnchorInput(from_row=20, from_col=3, to_row=30, to_col=5)
    ctx["Slicer"] = ctx["ws"].add_slicer(ctx["Table"], 0, SlicerInput(anchor=anchor))


def _threaded_comment(ctx: Context) -> None:
    ctx["wb"].add_person(Person(id=PERSON, display_name="Alice", user_id="alice", provider_id="None"))
    ctx["ws"].add_threaded_comment(
        ThreadedCommentInput(
            id=THREAD,
            cell=CellAddress(4, 2),
            person_id=PERSON,
            created=datetime.datetime(2026, 1, 2, 3, 4, 5),
            text="Hi",
        )
    )


# Ledger create-op id -> recipe that adds what the op builds to the context.
CREATE_OPS: Dict[str, Callable[[Context], None]] = {
    "table.create": _table,
    "pivot_cache.create": _pivot_cache,
    "pivot.create": _pivot,
    "chart.add": _chart,
    "shape.add": _shape,
    "slicer.add": _slicer,
    "threaded_comment.add": _threaded_comment,
    "cf_rule.add": lambda c: c["ws"].add_cf_rule(CfExpressionRuleInput(ranges="A1:B2", formula1="A1>1")),
    "protected_range.add": lambda c: c["ws"].add_protected_range(
        ProtectedRangeInput(name="Totals", ranges=[CellRect(1, 1, 2, 2)])
    ),
    "sparkline_group.add": lambda c: c["ws"].add_sparkline_group(
        SparklineGroupInput(type=SparklineType.COLUMN, sparklines=[SparklineInput("Sheet1!A1:D1", "F1")])
    ),
    "validation.add": lambda c: c["ws"].add_validation(
        DataValidationInput(type=ValidationType.WHOLE, ranges="D1:D3", formula1="1", formula2="5")
    ),
}

# A handle row whose `create` list does not build its handle gets the ops that do.
HANDLE_DEFAULT_OPS: Dict[str, List[str]] = {
    "Pivot": ["pivot_cache.create", "pivot.create"],
    "Table": ["table.create"],
}

# A readd row whose `create` list is empty gets the op that builds the object it reads.
ROW_DEFAULT_OPS: Dict[str, List[str]] = {
    "cf_rule": ["cf_rule.add"],
    "validation": ["validation.add"],
}


def _seed_hyperlink(ctx: Context) -> None:
    ctx["ws"].set_text(0, 1, "Example")
    ctx["ws"].set_hyperlink(HyperlinkInput(range="B1", target="https://example.com/", display="Example"))


# Row id -> state set on the built object before the first read, where the default is not enough to tell a
# round trip from a no-op.
SEEDS: Dict[str, Callable[[Context], None]] = {
    "cell_format": lambda c: c["ws"].update_cell_format("B1", font={"bold": True}, number_format="0.00"),
    "header_footer": lambda c: c["ws"].update_header_footer(odd_header="&CTitle"),
    "page_setup": lambda c: c["ws"].update_page_setup(scale_pct=150),
    "page_margins": lambda c: c["ws"].update_page_margins(left_in=0.5),
    "print_options": lambda c: c["ws"].update_print_options(grid_lines=True),
    "pivot_style_info": lambda c: c["Pivot"].update_style_options(
        style_name="PivotStyleLight16", show_row_headers=True
    ),
    "sheet_protection": lambda c: c["ws"].update_sheet_protection(allow_sort=True, allow_format_cells=True),
    "comment": lambda c: c["ws"].set_comment(0, 1, CommentInput(text="Note", author="Alice")),
    "hyperlink": _seed_hyperlink,
}

# Row id -> (getter arguments, setter arguments) where the row's own address arguments do not fit the Python
# signature (A1 text for a range, an id, a field and item index).
ARGS: Dict[str, Tuple[tuple, tuple]] = {
    "cell_format": ((0, 1), ("B1",)),
    "threaded_comment": ((THREAD,), (THREAD,)),
    "pivot_field": ((0,), (0,)),
    "pivot_field_item": ((0, 0), (0, 0)),
    # A comment is re-added on another cell, a hyperlink in place (its record carries its own range).
    "comment": ((0, 1), (0, 2)),
    "hyperlink": ((0, 1), ()),
    # Position 0 of the sheet's list for the id-or-position addressed rows.
    "cf_rule_update": ((0,), (0,)),
    "protected_range": ((0,), (0,)),
    "sparkline_group": ((0,), (0,)),
    "validation_update": ((0,), (0,)),
    "cf_rule": ((0,), ()),
    "validation": ((0,), ()),
}

# readd rows: where the re-added copy is read back, and the identity fields that differ between it and the original.
AGAIN: Dict[str, Tuple[Callable[[Any, Context], Any], Tuple[str, ...]]] = {
    "cf_rule": (lambda ws, _c: ws.list_cf_rules()[1], ("id", "priority")),
    "validation": (lambda ws, _c: ws.list_validations()[1], ("id",)),
    "comment": (lambda ws, _c: ws.get_comment(0, 2), ()),
    "hyperlink": (lambda ws, _c: ws.get_hyperlink(0, 1), ()),
}


def _handle_snake(handle: str) -> str:
    return "".join("_" + ch.lower() if ch.isupper() else ch for ch in handle).lstrip("_")


def _method(owner: Any, name: str, handle: Optional[str]) -> Callable[..., Any]:
    """The bound method for ledger Python name ``name``; on a handle the handle's own noun is dropped from it."""
    if hasattr(owner, name):
        return getattr(owner, name)
    if handle is not None:
        stripped = name.replace("_" + _handle_snake(handle), "", 1)
        for candidate in (stripped, "get_properties" if stripped == "get" else stripped):
            if hasattr(owner, candidate):
                return getattr(owner, candidate)
    raise AssertionError(f"no method {name!r} on {type(owner).__name__}")


def _keywords(setter: Callable[..., Any], record: Any) -> Dict[str, Any]:
    """The record's fields that ``setter`` takes as keyword-only parameters."""
    names = {p.name for p in inspect.signature(setter).parameters.values() if p.kind is inspect.Parameter.KEYWORD_ONLY}
    return {f.name: getattr(record, f.name) for f in dataclasses.fields(record) if f.name in names}


def _without(record: Any, keys: Tuple[str, ...]) -> Any:
    return dataclasses.replace(record, **{key: None for key in keys}) if keys else record


class GetterSetterRoundTripTests(unittest.TestCase):
    def test_every_getter_record_goes_back_through_its_setter(self) -> None:
        cases = json.loads((API_DIR / "roundtrip_cases.json").read_text(encoding="utf-8"))["cases"]
        pending = {
            p["id"]: p["reason"]
            for p in json.loads((API_DIR / "roundtrip_pending.json").read_text(encoding="utf-8"))["pending"]
            if SURFACE in p["surfaces"]
        }
        for row in cases:
            if SURFACE not in row["surfaces"]:
                continue
            with self.subTest(row["id"]):
                if row["id"] in pending:
                    self.skipTest(pending[row["id"]])
                self._run(row)

    def _args(self, row: Dict[str, Any]) -> Tuple[tuple, tuple]:
        if row["id"] in ARGS:
            return ARGS[row["id"]]
        for side in ("get", "set"):
            for key in ("cell_args", "ref_args"):
                if key in row[side] and row["id"] not in ARGS:
                    self.fail(f"row {row['id']}: {side} carries {key} but the driver has no Python arguments for it")
        return (), ()

    def _run(self, row: Dict[str, Any]) -> None:
        rid, handle = row["id"], row.get("handle")
        ops = list(row["create"]) or ROW_DEFAULT_OPS.get(rid, [])
        with Workbook.create_default() as wb:
            ctx: Context = {"wb": wb, "ws": wb.sheet(0)}
            for op in ops:
                self.assertIn(op, CREATE_OPS, f"row {rid}: no driver recipe for create op {op}")
                CREATE_OPS[op](ctx)
            if handle is not None and handle not in ctx:
                for op in HANDLE_DEFAULT_OPS.get(handle, []):
                    CREATE_OPS[op](ctx)
                self.assertIn(handle, ctx, f"row {rid}: no driver recipe builds a {handle}")
            if rid in SEEDS:
                SEEDS[rid](ctx)
            target = ctx[handle] if handle is not None else (wb if row["scope"] == "workbook" else ctx["ws"])
            get_args, set_args = self._args(row)
            getter = _method(target, row["get"]["py"], handle)
            setter = _method(target, row["set"]["py"], handle)

            def read() -> Any:
                record = getter(*get_args)
                path = row["get"].get("path")
                return getattr(record, path) if path else record

            before = read()
            self.assertIsNotNone(before, f"row {rid}: the getter returned None")
            if row["kind"] == "patch":
                keywords = _keywords(setter, before)
                self.assertTrue(keywords, f"row {rid}: the setter takes none of the record's fields")
                setter(*set_args, **keywords)
                self.assertEqual(read(), before)
            elif row["kind"] == "readd":
                record = before
                if row.get("clear"):
                    record = dataclasses.replace(record, **{row["clear"]: None})
                setter(*set_args, record)
                self.assertIn(rid, AGAIN, f"row {rid}: no driver recipe reads back the re-added copy")
                again, identity = AGAIN[rid]
                self.assertEqual(_without(again(ctx["ws"], ctx), identity), _without(before, identity))
            else:
                self.fail(f"row {rid}: unknown kind {row['kind']!r}")

    def test_a_rule_that_does_not_use_an_argument_goes_back_unchanged(self) -> None:
        for rule in (
            CfTop10RuleInput(ranges="A1:B2", rank=3),
            CfCellIsRuleInput(ranges="A1:B2", operator=ComparisonOperator.GREATER_THAN, formula1="1"),
        ):
            with self.subTest(type(rule).__name__), Workbook.create_default() as wb:
                ws = wb.sheet(0)
                ws.add_cf_rule(rule)
                before = ws.get_cf_rule(0)
                ws.update_cf_rule(0, **_keywords(ws.update_cf_rule, before))
                self.assertEqual(ws.get_cf_rule(0), before)


if __name__ == "__main__":
    unittest.main()
