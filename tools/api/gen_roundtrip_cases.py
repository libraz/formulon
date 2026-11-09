#!/usr/bin/env python3
"""Generates `tools/api/roundtrip_cases.json` from the public-name ledger.

A round-trip case pairs a getter with the setter that takes back what the getter
returned, so the binding suites can check `set(get())` leaves the state unchanged.
Two kinds of row are derived:

* `patch`: every `patch: true` op of `tools/api/names.yaml` that names a
  `patch_record`, paired with the getter of that record.
* `readd`: the record-taking `add` / `set` ops listed in `READD` below, paired with
  their getter (`add(get(i))` must be accepted).

A getter or an op the ledger no longer carries, and a `patch: true` op that has no
getter, fail the generator, so the case table cannot silently drop a pair.

Usage:
    gen_roundtrip_cases.py           rewrite roundtrip_cases.json
    gen_roundtrip_cases.py --check   exit 1 if the file differs from the ledger
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

TOOLS_API = Path(__file__).resolve().parent
LEDGER = TOOLS_API / "names.yaml"
OUTPUT = TOOLS_API / "roundtrip_cases.json"
BINDING_SURFACES = ("wasm", "native", "py")

DESCRIPTION = (
    "Getter -> setter round trips, generated from tools/api/names.yaml by gen_roundtrip_cases.py. "
    "Each row: id; kind (patch: set(get()) passes the record back whole to a `patch: true` op; "
    "readd: add(get(i)) accepts the record the getter returned); scope (sheet | workbook); "
    "record (the record the getter returns); get / set / create ops as {op, js, py, cell_args, ref_args} "
    "taken from the ledger; get.path (a field of the getter's result that holds the record); "
    "create (ledger op ids run in order to build the object before the round trip; empty for a singleton); "
    "handle (the handle class when the op is a handle method); surfaces (binding surfaces that carry both ops); "
    "known_fix (a failure the row exposed, fixed by the named rule). "
    "A row listed in roundtrip_pending.json is skipped by the drivers with its reason."
)

# Overrides for patch ops, keyed by op id: the row `id` where `<stem>` is not a name, the getter where it
# is not `<stem>.get`, and the ops that build the object the patch applies to.
_PIVOT = ["pivot_cache.create", "pivot.create"]
PATCH_OVERRIDES: Dict[str, dict] = {
    "binding.set_view": {"id": "sheet_view", "get": "view.get"},
    "binding.update_pivot_field": {"id": "pivot_field", "get": "pivot_field.get", "create": _PIVOT},
    "binding.update_pivot_item": {"id": "pivot_item", "get": "pivot_field_item.get", "create": _PIVOT},
    "binding.set_cell_format": {"id": "cell_format", "get": "binding.get_cell_format"},
    "chart_legend.update": {"get": "chart.get", "path": "legend", "create": ["chart.add"]},
    # Chart edits other than the legend: their round trips live in the chart suites.
    "chart_display.set": {"skip": True},
    "chart_axis.set": {"skip": True},
    "chart_series.set": {"skip": True},
    "chart_point.update": {"skip": True},
    "chart_plot.set": {"skip": True},
    # The common drawing patch: its round trips live in the drawing suites.
    "drawing.update": {"skip": True},
    "table.update": {"create": ["table.create"]},
    "protected_range.update": {"get": "protected_range.get_at", "create": ["protected_range.add"]},
    "cf_rule.update": {"id": "cf_rule_update", "get": "cf_rule.get_at", "create": ["cf_rule.add"]},
    "validation.update": {"id": "validation_update", "get": "validation.get_at", "create": ["validation.add"]},
    "threaded_comment.update": {"create": ["threaded_comment.add"]},
    "shape.update": {"create": ["shape.add"]},
    "slicer.update": {"create": ["table.create", "slicer.add"]},
    "sparkline_group.update": {"get": "sparkline_group.get_at", "create": ["sparkline_group.add"]},
    "pivot_grand_totals.set": {"create": _PIVOT},
    "pivot_style_info.set": {"create": _PIVOT},
    "pivot_options.set": {"create": _PIVOT},
    "pivot_cache_options.set": {"create": ["pivot_cache.create"]},
}

# Record-taking adds: row id -> (getter op, adding op, record, known_fix).
READD: Dict[str, tuple] = {
    "cf_rule": (
        "cf_rule.get_at",
        "cf_rule.add",
        "ConditionalFormat",
        "an id equal to an existing rule's gets a new GUID",
    ),
    "validation": ("validation.get_at", "validation.add", "DataValidation", None),
    "hyperlink": ("hyperlink.get", "hyperlink.set", "Hyperlink", "location is `string | null | undefined`"),
    "comment": ("comment.get", "comment.set", "Comment", None),
}

# Failures a patch row exposed, fixed by the rule named here.
PATCH_KNOWN_FIX: Dict[str, str] = {
    "binding.set_view": "a view without panes whose paneTopLeft equals the default is not refused",
}


def _load_ledger() -> dict:
    import yaml  # type: ignore

    return yaml.safe_load(LEDGER.read_text(encoding="utf-8"))["entries"]


def _op(entries: dict, op_id: str, why: str) -> dict:
    entry = entries.get(op_id)
    if entry is None:
        raise SystemExit(f"gen_roundtrip_cases: {why}: ledger has no entry `{op_id}`")
    if not entry.get("js") or not entry.get("py"):
        raise SystemExit(f"gen_roundtrip_cases: {why}: `{op_id}` has no JS or Python name")
    return entry


def _ref(entries: dict, op_id: str, why: str) -> dict:
    entry = _op(entries, op_id, why)
    out: Dict[str, object] = {"op": op_id, "js": entry["js"], "py": entry["py"]}
    for key in ("cell_args", "ref_args"):
        if key in entry:
            out[key] = entry[key]
    return out


def _surfaces(*entries: dict) -> List[str]:
    return [s for s in BINDING_SURFACES if all(s in e.get("surfaces", []) for e in entries)]


def _row(entries: dict, row_id: str, kind: str, get_id: str, set_id: str, record: str, extra: dict) -> dict:
    getter = _op(entries, get_id, row_id)
    setter = _op(entries, set_id, row_id)
    create = extra.get("create", [])
    created = [_op(entries, c, row_id) for c in create]
    surfaces = _surfaces(getter, setter, *created)
    if not surfaces:
        raise SystemExit(f"gen_roundtrip_cases: {row_id}: no binding surface carries every op of the row")
    row: Dict[str, object] = {
        "id": row_id,
        "kind": kind,
        "scope": setter.get("scope"),
        "record": record,
        "get": _ref(entries, get_id, row_id),
        "set": _ref(entries, set_id, row_id),
        "create": create,
        "surfaces": surfaces,
    }
    if "path" in extra:
        row["get"]["path"] = extra["path"]  # type: ignore[index]
    handle = setter.get("handle")
    if isinstance(handle, dict):
        row["handle"] = handle["class"]
    if extra.get("known_fix"):
        row["known_fix"] = extra["known_fix"]
    return row


def build(entries: dict) -> dict:
    rows: List[dict] = []
    for op_id, entry in entries.items():
        if entry.get("patch") is not True or entry.get("patch_record") in (None, "none"):
            continue
        over = PATCH_OVERRIDES.get(op_id, {})
        if over.get("skip"):
            continue
        get_id = over.get("get", op_id.rsplit(".", 1)[0] + ".get")
        if get_id not in entries:
            raise SystemExit(
                f"gen_roundtrip_cases: patch op `{op_id}` has no getter `{get_id}`; add a PATCH_OVERRIDES entry"
            )
        extra = dict(over)
        if op_id in PATCH_KNOWN_FIX:
            extra["known_fix"] = PATCH_KNOWN_FIX[op_id]
        row_id = over.get("id", op_id.rsplit(".", 1)[0])
        rows.append(_row(entries, row_id, "patch", get_id, op_id, entry["patch_record"], extra))
    for row_id, (get_id, add_id, record, known_fix) in READD.items():
        rows.append(_row(entries, row_id, "readd", get_id, add_id, record, {"known_fix": known_fix}))
    ids = [r["id"] for r in rows]
    dupes = sorted({i for i in ids if ids.count(i) > 1})
    if dupes:
        raise SystemExit(f"gen_roundtrip_cases: duplicate row ids {dupes}")
    rows.sort(key=lambda r: (r["kind"], r["id"]))
    return {"description": DESCRIPTION, "cases": rows}


def render(entries: dict) -> str:
    return json.dumps(build(entries), indent=2, ensure_ascii=True) + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", action="store_true", help="fail if roundtrip_cases.json is not current")
    args = parser.parse_args(argv)
    text = render(_load_ledger())
    if args.check:
        current = OUTPUT.read_text(encoding="utf-8") if OUTPUT.exists() else ""
        if current != text:
            print(f"{OUTPUT.name} is stale against names.yaml; run tools/api/gen_roundtrip_cases.py", file=sys.stderr)
            return 1
        return 0
    OUTPUT.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
