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

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import yaml

# Recognised value kinds, mirroring src/value.h's ValueKind minus Array/Ref/
# Lambda which the oracle pipeline doesn't exercise yet.
KINDS = {"blank", "number", "bool", "text", "error", "formula"}

# Recognised compare modes for the C++ verifier. `None` / "exact" means
# strict byte-equality (the historical behaviour). Any other value selects
# a structured comparator on the verifier side; see `compare_value` in
# `tests/oracle/oracle_test.cpp`.
COMPARE_MODES = {
    "exact",
    "complex_text",
    "datevalue_roundtrip_readback",
    "empty_string_readback",
    "numeric_text",
}

# Excel's worksheet dimensions, used to bound an author-supplied formula
# placement. `drivers/base.py` carries its own copy on purpose: that module
# is documented as importing nothing but the standard library so drivers can
# be exercised without the YAML layer.
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384

# A bare, relative A1 cell address: column letters plus a row number, with
# no sheet qualifier, no `$` anchors, and no `:` range part.
_A1_ADDRESS_RE = re.compile(r"(?P<col>[A-Z]{1,3})(?P<row>[1-9][0-9]{0,6})")

# Ceiling on `samples` for a shape-captured case. Shape capture exists to
# read a small fixed set of cells instead of walking a spill, so the list
# stays short enough that the reader cannot become the walk it replaced.
MAX_SHAPE_SAMPLES = 64


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

    `setup` maps cell addresses to value records; the Python generator
    writes these into the Excel workbook before triggering a calc. The
    address can be either:

      * A bare A1 reference (``"A1"``, ``"BC42"``) — applies to the
        default sheet (``"Sheet1"``).
      * A sheet-qualified A1 reference (``"Sheet2!A1"``,
        ``"'My Sheet'!B5"``) — the driver creates the named sheet on
        first reference. The single-quoted form is required when the
        sheet name contains a space or other special character, matching
        Excel's own quoting convention.

    `formula` is the formula under test — always spelled with a leading
    `=` in the YAML to match how the Formulon tokenizer and Excel itself
    read it. It is placed at ``Sheet1!Z1`` by default (the formula cell
    convention; cross-sheet refs from the formula resolve through the
    setup-populated sheets).

    `formula_cell`, when set, moves the formula under test to that bare
    A1 address on the default sheet instead of ``Z1``. Drivers and the
    native verifier both honour it, so a case can probe behaviour that
    depends on where the formula sits — whether a whole-axis spill fits
    below the formula row, or whether the formula's own cell falls inside
    the range it references. Omitting the field keeps the historical
    ``Z1`` placement, so existing cases are unaffected.

    `capture` selects how the result is recorded. The default cell walk
    materialises every cell of a spill and is bounded by
    ``drivers.base.MAX_CAPTURE_CELLS``. ``capture: shape`` instead records
    the spill's shape plus the cells listed in `samples`, which is how a
    case whose true result is larger than that ceiling is verified at all
    -- the smallest whole-axis spill is 16,384 cells. Excel answers
    identically either way; only the recording differs.

    `samples` lists absolute A1 addresses inside the spill (they are read
    on the case's own sheet, so they are the formula cell and cells at
    positive offsets from it). Required with ``capture: shape`` and
    rejected otherwise.

    `merges`, when present, is a list of inclusive A1 ranges on the default
    sheet (for example ``["Z1:AA1"]``). Drivers and the native verifier
    apply these ranges before evaluating the formula so merged-cell spill
    collisions are observed consistently.

    `tolerance`, if set, overrides the suite default for this case only.

    `compare_mode`, if set, switches the C++ verifier to a structured
    comparator (e.g. `"complex_text"` parses both expected and actual text
    as Excel complex numbers and applies `tolerance` to each component).
    `None` or `"exact"` keeps the default strict byte equality.
    """

    id: str
    formula: str
    setup: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    # Optional bare A1 address of the formula under test on the default
    # sheet. `None` means the default `Z1` placement.
    formula_cell: Optional[str] = None
    # `"shape"` records the result as its dynamic-array shape plus the
    # `samples` cells instead of every cell; `None` is the default cell
    # walk. `samples` is non-empty exactly when capture is `"shape"`.
    capture: Optional[str] = None
    samples: List[str] = field(default_factory=list)
    # Optional inclusive A1 merge ranges applied to the default sheet before
    # setup cells and the formula-under-test are evaluated. Formula-oracle
    # cases keep these ranges on the case so the native verifier and Excel
    # drivers observe the same sheet metadata.
    merges: List[str] = field(default_factory=list)
    description: str = ""
    tolerance: Optional[Tolerance] = None
    compare_mode: Optional[str] = None
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


def _normalise_a1_address(raw: Any, *, where: str, field_name: str) -> Optional[str]:
    """Validate an optional bare, relative A1 cell address.

    No sheet qualifier and no ``$`` anchors, because every driver reads and
    writes these through ``sheet.range(addr)`` on the case's own worksheet.
    The address is upper-cased so the emitted golden is independent of how
    the YAML spelled it.
    """

    if raw is None:
        return None
    if not isinstance(raw, str) or not raw.strip():
        raise ValueError(f"{where}: {field_name} must be a non-empty A1 address string")
    addr = raw.strip().upper()
    match = _A1_ADDRESS_RE.fullmatch(addr)
    if match is None:
        raise ValueError(
            f"{where}: {field_name} '{raw}' is not a bare relative A1 address (no sheet qualifier, no '$', no range)"
        )
    column = 0
    for letter in match.group("col"):
        column = column * 26 + (ord(letter) - ord("A") + 1)
    row = int(match.group("row"))
    if row < 1 or row > EXCEL_MAX_ROWS or column < 1 or column > EXCEL_MAX_COLUMNS:
        raise ValueError(f"{where}: {field_name} '{raw}' is outside the Excel grid")
    return addr


def _normalise_formula_cell(raw: Any, *, where: str) -> Optional[str]:
    """Validate the optional formula-under-test placement address."""

    return _normalise_a1_address(raw, where=where, field_name="formula_cell")


def _normalise_capture(raw_capture: Any, raw_samples: Any, *, where: str) -> Tuple[Optional[str], List[str]]:
    """Validate the optional shape-capture mode and its sample addresses.

    Sample addresses use the same bare relative A1 form as `formula_cell`
    and are capped: the point of shape capture is to read a small fixed set
    of cells instead of walking a spill, so an unbounded list would defeat
    it.
    """

    if raw_capture is None or raw_capture == "cells":
        if raw_samples is not None:
            raise ValueError(f"{where}: samples requires capture: shape")
        return None, []
    if raw_capture != "shape":
        raise ValueError(f"{where}: unknown capture mode '{raw_capture}'; expected 'cells' or 'shape'")
    if not isinstance(raw_samples, list) or not raw_samples:
        raise ValueError(f"{where}: capture: shape requires a non-empty samples list")
    if len(raw_samples) > MAX_SHAPE_SAMPLES:
        raise ValueError(f"{where}: capture: shape accepts at most {MAX_SHAPE_SAMPLES} samples")
    samples: List[str] = []
    for index, sample in enumerate(raw_samples):
        address = _normalise_a1_address(sample, where=f"{where}/samples[{index}]", field_name="sample")
        if address is None:
            raise ValueError(f"{where}/samples[{index}]: expected an A1 address")
        if address in samples:
            raise ValueError(f"{where}/samples[{index}]: duplicate sample address '{address}'")
        samples.append(address)
    return "shape", samples


def _normalise_merges(raw: Any, *, where: str) -> List[str]:
    """Validate and copy optional inclusive A1 merge-range references."""

    if raw is None:
        return []
    if not isinstance(raw, list):
        raise ValueError(f"{where}: merges must be a list of A1 ranges")
    merges: List[str] = []
    for i, merge in enumerate(raw):
        if not isinstance(merge, str) or not merge.strip():
            raise ValueError(f"{where}/merges[{i}]: expected a non-empty A1 range string")
        merges.append(merge)
    return merges


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
            raise ValueError(f"{path}: case '{cid}' 'formula' must start with '='")

        raw_setup = raw.get("setup") or {}
        if not isinstance(raw_setup, dict):
            raise ValueError(f"{path}: case '{cid}' 'setup' must be a mapping")
        setup: Dict[str, Dict[str, Any]] = {}
        for addr, value in raw_setup.items():
            if not isinstance(addr, str):
                raise ValueError(f"{path}: case '{cid}' setup key must be a string A1 address")
            setup[addr] = _normalise_value(value, where=f"case '{cid}', setup[{addr}]")

        merges = _normalise_merges(raw.get("merges"), where=f"case '{cid}'")

        formula_cell = _normalise_formula_cell(raw.get("formula_cell"), where=f"{path}: case '{cid}'")

        capture, samples = _normalise_capture(raw.get("capture"), raw.get("samples"), where=f"{path}: case '{cid}'")

        case_tol = _load_tolerance(raw.get("tolerance"), where=f"{path}: case '{cid}'")

        compare_mode_raw = raw.get("compare_mode")
        compare_mode: Optional[str] = None
        if compare_mode_raw is not None:
            if not isinstance(compare_mode_raw, str):
                raise ValueError(f"{path}: case '{cid}' 'compare_mode' must be a string")
            if compare_mode_raw not in COMPARE_MODES:
                raise ValueError(
                    f"{path}: case '{cid}' has unknown compare_mode "
                    f"'{compare_mode_raw}'; expected one of {sorted(COMPARE_MODES)}"
                )
            # "exact" is the implicit default; normalise to None so the
            # generator omits the field for the historical strict path.
            if compare_mode_raw != "exact":
                compare_mode = compare_mode_raw

        author_expect: Optional[Dict[str, Any]] = None
        if "expect" in raw and raw["expect"] is not None:
            author_expect = _normalise_value(raw["expect"], where=f"case '{cid}', expect")

        cases.append(
            Case(
                id=cid,
                formula=formula,
                setup=setup,
                formula_cell=formula_cell,
                capture=capture,
                samples=samples,
                merges=merges,
                description=raw.get("description", "") or "",
                tolerance=case_tol if raw.get("tolerance") is not None else None,
                compare_mode=compare_mode,
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
