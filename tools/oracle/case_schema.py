#!/usr/bin/env python3
"""Case YAML -> dataclass loader + golden JSON emitter.

The YAML files under `tests/oracle/cases/<category>.yaml` are the
human-authored source of truth. The C++ verifier (`tests/oracle/`) does not
read YAML at all — it reads the *golden JSON* that this module emits. The
split keeps YAML readable (shorthand values, comments) while the JSON is
rigid enough for a ~200 line hand-rolled C++ reader.

The canonical shape of an ingested case is documented on `Case` below.
Value records (setup cells + observed results) use the uniform
`{kind, value}` shape — Formulon's C++ side can then dispatch a single
switch over `kind`.
"""

from __future__ import annotations

import dataclasses
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import yaml


# Recognised value kinds, mirroring src/value.h's ValueKind minus Array/Ref/
# Lambda which the oracle pipeline doesn't exercise yet.
KINDS = {"blank", "number", "bool", "text", "error", "formula"}


@dataclass
class Tolerance:
    """Accepted numeric tolerance for a case or suite.

    Both fields are absolute-or-relative thresholds — the verifier passes
    when *either* condition is met. Zero means strict equality on that axis.
    """

    abs: float = 0.0
    rel: float = 0.0

    def to_dict(self) -> Dict[str, float]:
        return {"abs": self.abs, "rel": self.rel}


@dataclass
class Case:
    """A single oracle case.

    `setup` maps A1 addresses ("A1", "BC42") to value records; the Python
    generator writes these into the Excel workbook before triggering a
    calc. `formula` is the formula under test — always spelled with a
    leading `=` in the YAML to match how the Formulon tokenizer and Excel
    itself read it.

    `tolerance`, if set, overrides the suite default for this case only.
    """

    id: str
    formula: str
    setup: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    description: str = ""
    tolerance: Optional[Tolerance] = None
    # Parsed author-side expectation, kept for documentation only. The
    # verifier compares against the golden JSON (Excel's observed output),
    # not this field.
    author_expect: Optional[Dict[str, Any]] = None


@dataclass
class Suite:
    """The content of one `<category>.yaml` file."""

    name: str
    description: str
    locale: str
    tolerance: Tolerance
    options: Dict[str, Any]
    cases: List[Case]


def _normalise_value(raw: Any, *, where: str) -> Dict[str, Any]:
    """Canonicalise a shorthand setup value or author-expect into {kind,...}.

    Shorthand rules:
      * int/float  -> {"kind": "number", "value": float(raw)}
      * str starting with "="  -> {"kind": "formula", "formula": raw}
      * str otherwise  -> {"kind": "text", "value": raw}
      * bool  -> {"kind": "bool", "value": raw}
      * None  -> {"kind": "blank"}
      * dict with "kind"  -> validated pass-through

    `where` is a free-form label used in error messages to point at the
    YAML location (e.g. "case my-id, setup[A1]").
    """

    if raw is None:
        return {"kind": "blank"}
    if isinstance(raw, bool):
        # bool MUST be checked before int because bool is-a int in Python.
        return {"kind": "bool", "value": raw}
    if isinstance(raw, (int, float)):
        return {"kind": "number", "value": float(raw)}
    if isinstance(raw, str):
        if raw.startswith("="):
            return {"kind": "formula", "formula": raw}
        return {"kind": "text", "value": raw}
    if isinstance(raw, dict):
        if "kind" not in raw:
            raise ValueError(f"{where}: dict value missing 'kind'")
        kind = raw["kind"]
        if kind not in KINDS:
            raise ValueError(f"{where}: unknown kind '{kind}'")
        # Shallow copy so the YAML loader's mapping isn't mutated.
        return dict(raw)
    raise ValueError(f"{where}: unsupported value type {type(raw).__name__}")


def _load_tolerance(raw: Any, *, where: str) -> Tolerance:
    if raw is None:
        return Tolerance()
    if not isinstance(raw, dict):
        raise ValueError(f"{where}: tolerance must be a mapping")
    abs_t = float(raw.get("abs", 0.0))
    rel_t = float(raw.get("rel", 0.0))
    return Tolerance(abs=abs_t, rel=rel_t)


def load_suite(path: Path) -> Suite:
    """Parses a single `<category>.yaml` file into a :class:`Suite`.

    The file must have a top-level `suite` field (the category name) and a
    `cases` list. All other fields are optional.
    """

    with path.open("r", encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    if not isinstance(doc, dict):
        raise ValueError(f"{path}: top-level YAML must be a mapping")
    name = doc.get("suite")
    if not isinstance(name, str) or not name:
        raise ValueError(f"{path}: missing 'suite' string")

    description = doc.get("description", "") or ""
    locale = doc.get("locale", "ja-JP") or "ja-JP"
    tolerance = _load_tolerance(doc.get("tolerance"), where=f"{path}")
    options = doc.get("options") or {}
    if not isinstance(options, dict):
        raise ValueError(f"{path}: 'options' must be a mapping")

    raw_cases = doc.get("cases") or []
    if not isinstance(raw_cases, list):
        raise ValueError(f"{path}: 'cases' must be a list")

    cases: List[Case] = []
    seen_ids: set[str] = set()
    for i, raw in enumerate(raw_cases):
        if not isinstance(raw, dict):
            raise ValueError(f"{path}: case #{i} is not a mapping")
        cid = raw.get("id")
        if not isinstance(cid, str) or not cid:
            raise ValueError(f"{path}: case #{i} missing 'id'")
        if cid in seen_ids:
            raise ValueError(f"{path}: duplicate case id '{cid}'")
        seen_ids.add(cid)

        formula = raw.get("formula")
        if not isinstance(formula, str) or not formula.startswith("="):
            raise ValueError(
                f"{path}: case '{cid}' 'formula' must start with '='"
            )

        raw_setup = raw.get("setup") or {}
        if not isinstance(raw_setup, dict):
            raise ValueError(f"{path}: case '{cid}' 'setup' must be a mapping")
        setup: Dict[str, Dict[str, Any]] = {}
        for addr, value in raw_setup.items():
            if not isinstance(addr, str):
                raise ValueError(
                    f"{path}: case '{cid}' setup key must be a string A1 address"
                )
            setup[addr] = _normalise_value(
                value, where=f"case '{cid}', setup[{addr}]"
            )

        case_tol = _load_tolerance(
            raw.get("tolerance"), where=f"{path}: case '{cid}'"
        )

        author_expect: Optional[Dict[str, Any]] = None
        if "expect" in raw and raw["expect"] is not None:
            author_expect = _normalise_value(
                raw["expect"], where=f"case '{cid}', expect"
            )

        cases.append(
            Case(
                id=cid,
                formula=formula,
                setup=setup,
                description=raw.get("description", "") or "",
                tolerance=case_tol if raw.get("tolerance") is not None else None,
                author_expect=author_expect,
            )
        )

    return Suite(
        name=name,
        description=description,
        locale=locale,
        tolerance=tolerance,
        options=options,
        cases=cases,
    )


def discover_suites(cases_dir: Path) -> List[Tuple[Path, Suite]]:
    """Loads every `*.yaml` file under `cases_dir` (non-recursive).

    Returns a list of `(path, suite)` pairs sorted by path so the generator
    output is deterministic regardless of filesystem order.
    """

    if not cases_dir.exists() or not cases_dir.is_dir():
        return []
    out: List[Tuple[Path, Suite]] = []
    for path in sorted(cases_dir.iterdir()):
        if path.suffix.lower() not in {".yaml", ".yml"}:
            continue
        if path.name.startswith("."):
            continue
        out.append((path, load_suite(path)))
    return out
