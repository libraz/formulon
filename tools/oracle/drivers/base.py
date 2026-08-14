#!/usr/bin/env python3
"""Driver-agnostic types and helpers shared across oracle backends.

This module deliberately has *no* dependency on xlwings, AppleScript, or
any platform-specific automation library. Concrete drivers (Mac Excel via
xlwings, future Windows COM, future WSL2 bridge) live in sibling modules
under `tools.oracle.drivers.<name>` and import from here.

Contents:
    - :class:`OracleDriver` — abstract base class every driver implements.
    - :class:`CaseResult`   — normalised observation for a single case.
    - :class:`EnvironmentInfo` — environment snapshot for goldens / docs.
    - :data:`_ERR_DISPLAY_NAMES` — recognised Excel error tokens.
    - :data:`_EXCEL_EPOCH`, :func:`_datetime_to_serial` — date helpers.

Driver authors: subclass :class:`OracleDriver`, implement the four
abstract methods, and yield :class:`CaseResult` records whose `kind`
matches the small enumeration documented on that class.
"""

from __future__ import annotations

import abc
import datetime as _dt
import math
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

# Excel's error values come back from automation bridges as CVErr /
# ErrorValue / str depending on Mac vs Windows and the property we read.
# We centralise the recognition here so callers only see the tokenised
# display name.
_ERR_DISPLAY_NAMES = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#CONNECT!",
    "#EXTERNAL!",
    "#BUSY!",
    "#PYTHON!",
    "#UNKNOWN!",
}

# Excel's worksheet dimensions are part of the oracle contract.  The spill
# shape probe packs two positive dimensions into one scalar using
# ``rows * _EXCEL_MAX_COLUMNS + columns``.  Keeping the constants here makes
# the decoder platform-neutral and, importantly, keeps the Mac and Windows
# adapters byte-for-byte aligned on the wire expression.
_EXCEL_MAX_ROWS = 1_048_576
_EXCEL_MAX_COLUMNS = 16_384
# Formula-oracle capture materialises every cell through xlwings and emits
# the flattened values into JSON.  Keep that bounded even when Excel reports
# a valid but adversarially large spill.  This matches the existing 4096-cell
# oracle import range budget and prevents an external result from turning the
# shape probe into an unbounded offset loop.
MAX_CAPTURE_CELLS = 4_096


# Where the formula under test is written on the case's own worksheet unless
# the case carries a `formula_cell` override. `Z1` keeps the formula clear of
# the A1-based setup block while staying in row 1.
DEFAULT_FORMULA_CELL = "Z1"


class SpillShapeProbeError(RuntimeError):
    """Raised when a non-invasive dynamic-array shape probe is invalid."""


def case_shape_samples(case: Dict[str, Any]) -> Optional[List[str]]:
    """Returns the sample addresses for a shape-captured case, else None.

    A case whose true result is larger than :data:`MAX_CAPTURE_CELLS` cannot
    be recorded cell by cell -- the smallest whole-axis spill alone is
    16,384 cells. Such a case sets ``capture: shape`` and lists the
    absolute A1 addresses it wants read back; the driver records the
    dynamic-array shape plus exactly those cells. Excel is answering
    normally in both modes; only our recording differs.
    """

    capture = case.get("capture")
    if capture is None or capture == "cells":
        return None
    if capture != "shape":
        raise ValueError(f"case {case.get('id')!r} has an unknown capture mode: {capture!r}")
    samples = case.get("samples")
    if (
        not isinstance(samples, list)
        or not samples
        or not all(isinstance(item, str) and item.strip() for item in samples)
    ):
        raise ValueError(f"case {case.get('id')!r} declares capture: shape without a non-empty samples list")
    return [item.strip().upper() for item in samples]


def case_formula_cell(case: Dict[str, Any]) -> str:
    """Returns the A1 address a case's formula under test is written to.

    `formula_cell` is optional; absent, the driver uses the historical
    :data:`DEFAULT_FORMULA_CELL`. A present-but-unusable value raises
    instead of falling back, so a malformed override surfaces as a skipped
    case with a reason rather than being observed at the wrong cell.
    """

    if "formula_cell" not in case or case["formula_cell"] is None:
        return DEFAULT_FORMULA_CELL
    addr = case["formula_cell"]
    if not isinstance(addr, str) or not addr.strip():
        raise ValueError(f"case {case.get('id')!r} has an unusable formula_cell: {addr!r}")
    return addr.strip().upper()


def decode_spill_shape_probe(value: Any) -> tuple[int, int]:
    """Decode and validate ``ROWS(anchor#)*16384+COLUMNS(anchor#)``.

    ``COLUMNS`` is allowed to equal the column ceiling, so a zero remainder
    is decoded as ``(rows - 1, 16384)``.  Every other non-finite, fractional,
    boolean, zero, negative, or out-of-grid value is rejected rather than
    silently falling back to a fake 1x1 result.
    """

    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SpillShapeProbeError(f"spill shape probe returned non-numeric value: {value!r}")
    numeric = float(value)
    if not math.isfinite(numeric) or not numeric.is_integer():
        raise SpillShapeProbeError(f"spill shape probe returned non-integral value: {value!r}")
    encoded = int(numeric)
    if encoded <= 0:
        raise SpillShapeProbeError(f"spill shape probe returned non-positive value: {value!r}")

    rows, columns = divmod(encoded, _EXCEL_MAX_COLUMNS)
    if columns == 0:
        rows -= 1
        columns = _EXCEL_MAX_COLUMNS
    if not (1 <= rows <= _EXCEL_MAX_ROWS and 1 <= columns <= _EXCEL_MAX_COLUMNS):
        raise SpillShapeProbeError(f"spill shape probe is outside the Excel grid: {value!r}")
    return rows, columns


def probe_spill_shape(evaluate, anchor, *, max_cells: Optional[int] = MAX_CAPTURE_CELLS) -> tuple[int, int]:
    """Evaluate the shared, non-invasive spill-shape expression once.

    ``evaluate`` is a platform adapter around Excel's Application.Evaluate;
    it receives the exact expression string and must not calculate, write, or
    clear a worksheet.  The anchor address is obtained with
    ``get_address(external=True)`` and is deliberately passed through
    unmodified so workbook/sheet apostrophes remain escaped by xlwings.

    Excel evaluates the postfix ``#`` expression for scalar/error/blank
    anchors as the valid 1x1 encoding, so every result must pass through the
    strict decoder.  A bridge failure or malformed value therefore remains a
    hard probe error; callers must surface it as a skipped case rather than
    inventing a 1x1 fallback.

    ``max_cells`` bounds the shape because the caller normally walks every
    cell afterwards.  A shape-captured case reads a fixed, case-declared
    sample list instead of walking, so it passes ``None`` -- there is no
    offset loop for an adversarially large shape to run away with.  The
    grid-fit check below is a correctness check and always applies.
    """

    try:
        external_address = anchor.get_address(external=True)
    except Exception as exc:
        raise SpillShapeProbeError("could not obtain an external anchor address") from exc
    if not isinstance(external_address, str) or not external_address.strip():
        raise SpillShapeProbeError(f"anchor returned an invalid external address: {external_address!r}")

    coordinates = {}
    for name, limit in (("row", _EXCEL_MAX_ROWS), ("column", _EXCEL_MAX_COLUMNS)):
        try:
            coordinate = getattr(anchor, name)
        except Exception as exc:
            raise SpillShapeProbeError(f"anchor has no usable {name} coordinate") from exc
        if type(coordinate) is not int or not 1 <= coordinate <= limit:
            raise SpillShapeProbeError(f"anchor {name} coordinate is outside the Excel grid: {coordinate!r}")
        coordinates[name] = coordinate

    expression = f"ROWS({external_address}#)*{_EXCEL_MAX_COLUMNS}+COLUMNS({external_address}#)"
    try:
        value = evaluate(expression)
    except Exception as exc:
        raise SpillShapeProbeError(f"Excel shape probe failed for {external_address!r}") from exc

    rows, columns = decode_spill_shape_probe(value)
    available_rows = _EXCEL_MAX_ROWS - coordinates["row"] + 1
    available_columns = _EXCEL_MAX_COLUMNS - coordinates["column"] + 1
    if rows > available_rows or columns > available_columns:
        raise SpillShapeProbeError(
            f"spill shape {rows}x{columns} does not fit from anchor ({coordinates['row']},{coordinates['column']})"
        )
    if max_cells is not None and rows * columns > max_cells:
        raise SpillShapeProbeError(
            f"spill shape {rows}x{columns} exceeds the oracle capture ceiling of {max_cells} cells"
        )
    return rows, columns


@dataclass
class CaseResult:
    """Normalised observation for one case.

    Mirrors the on-wire JSON shape the generator will emit. `kind` is one
    of {'blank','number','bool','text','error','array'}. `array_shape` is
    populated only for spill results; array values are flattened row-major.
    Scalar array cells are represented as JSON primitives where possible
    (blank -> None, number/bool/text -> primitive), with error cells using
    `{"kind": "error", "code": "#N/A"}` records.
    """

    id: str
    kind: str
    # Populated based on kind:
    value: Any = None
    error_code: Optional[str] = None
    array_shape: Optional[List[int]] = None


@dataclass
class EnvironmentInfo:
    """Environment snapshot recorded next to every golden run.

    The fields are intentionally small and string-typed so they can be
    re-emitted into JSON / Markdown without any further marshalling.
    """

    excel_version: str
    excel_locale: str
    date1904: bool
    iterative: bool


# Excel's epoch relative to which serial numbers are counted. 1900-01-01
# is serial 1, so the subtraction base is 1899-12-31. Dates on or after
# 1900-03-01 get a +1 bump to account for Excel's phantom Feb 29 1900
# (serial 60), which Python's datetime cannot represent.
_EXCEL_EPOCH = _dt.datetime(1899, 12, 31)


def _datetime_to_serial(v: _dt.datetime) -> float:
    """Converts a Python datetime to an Excel 1900-system serial number.

    The 1900 leap-bug means every serial >= 60 is off-by-one relative to
    a naive epoch subtraction — Excel's day 61 (March 1, 1900) is only
    61 real days after 1899-12-30, but Excel treats 1900 as a leap year
    and assigns serial 60 to a non-existent Feb 29. Python datetime
    skips that ghost day, so we add 1 to compensate for anything strictly
    after 1900-02-28.
    """

    delta = (v - _EXCEL_EPOCH).total_seconds() / 86400.0
    if v >= _dt.datetime(1900, 3, 1):
        delta += 1
    return delta


class OracleDriver(abc.ABC):
    """Abstract base class for oracle backends.

    Concrete drivers own a single live Excel process (or equivalent) for
    the lifetime of a `with`-block, evaluate batches of formulas, and
    return :class:`CaseResult` records that the generator turns into
    golden JSON. All drivers must support the same case-record shape so
    `oracle_gen` can swap them transparently.
    """

    @abc.abstractmethod
    def __enter__(self) -> "OracleDriver": ...

    @abc.abstractmethod
    def __exit__(self, exc_type, exc, tb) -> None: ...

    @abc.abstractmethod
    def probe_environment(self) -> EnvironmentInfo:
        """Returns the version / locale / option snapshot for logging."""

    @abc.abstractmethod
    def run_suite(
        self,
        suite_name: str,
        cases: List[Dict[str, Any]],
        *,
        date1904: bool = False,
        iterative: bool = False,
    ) -> List[CaseResult]:
        """Evaluates every case in `cases` and returns observed results.

        Each `cases[i]` is a dict with keys `id`, `formula`, and `setup`
        (mapping A1 -> `{kind, value, ...}` record), plus the optional
        `merges` list of inclusive A1 ranges on the default sheet and the
        optional `formula_cell` placement override (see
        :func:`case_formula_cell`). The driver does not normalise; it
        trusts upstream `case_schema` to have done so.
        """

    def assert_m365_or_abort(self) -> None:
        """Aborts when the attached Excel install is not Microsoft 365.

        Probes the live Excel instance with `=ARRAYTOTEXT(1)`. The function
        is post-2019 (introduced with the dynamic-array family), so an
        install older than M365 returns `#NAME?` -- silently, with no
        warning. Continuing past that point would bake `#NAME?` into the
        golden JSON for every post-2019 function and contaminate the
        oracle, which is the exact failure mode that produced the
        now-deleted `win-2019-ja_JP` archive.

        Raises:
            RuntimeError: when Excel returns `#NAME?` for ARRAYTOTEXT, or
                when the driver reports the sentinel suite as skipped
                (the write-errors path: a driver that cannot evaluate the
                sentinel cannot evaluate the real suite either, so we
                refuse rather than emit partial goldens).

        The exact expected result is the text value ``"1"``. Any other
        result shape is rejected; a product/version sentinel must not be
        treated as a best-effort hint.
        """

        message = (
            "Excel does not recognise ARRAYTOTEXT — this Excel install is "
            "pre-M365 (Office 2019 or earlier).\n"
            "\n"
            "Formulon's oracle requires Microsoft 365 (Excel build 16.0 "
            "with the post-2019 dynamic-array / LAMBDA function set). "
            "Generating goldens on Office 2019 silently bakes #NAME? "
            "results into the JSON for every post-2019 function — exactly "
            "the failure mode that produced the now-deleted win-2019 "
            "archive.\n"
            "\n"
            "Upgrade to Excel 365 (https://www.microsoft.com/microsoft-365) "
            "and retry, or run a Microsoft 365 install on a different host."
        )
        results = self.run_suite(
            "__formulon_m365_sentinel__",
            [{"id": "sentinel", "formula": "=ARRAYTOTEXT(1)", "setup": {}}],
        )
        if not results:
            raise RuntimeError("M365 sentinel probe returned no result; refusing unverified capture")
        result = results[0]
        if result.kind == "error" and result.error_code == "#NAME?":
            raise RuntimeError(message)
        if result.kind == "skipped":
            raise RuntimeError(message)
        if result.kind != "text" or result.value != "1":
            raise RuntimeError(
                "Microsoft 365 sentinel ARRAYTOTEXT(1) returned an unexpected "
                f"result: kind={result.kind!r}, value={result.value!r}"
            )

    def run_workbook_case(self, case: Dict[str, Any]) -> Dict[str, Any]:
        """Evaluates one declarative workbook case and returns the result.

        The workbook oracle track covers workbook-level features that are
        NOT formula results -- pivot tables and print areas. `case` is the
        declarative mini-workbook spec (`sheets`, `column_widths`,
        `row_heights`, and the optional `pivot` / `print` feature blocks)
        validated by `tools/oracle/workbook_case_schema.py`. The returned
        mapping is the observed pivot / print result, written verbatim
        into the golden's `expect` block.

        This base implementation is a stub. A later phase fills it in for
        each concrete driver via the helper steps `build_workbook`,
        `build_pivot`, and `apply_print`; the Windows COM driver lands
        first because reliable PivotTable automation needs it. Until then
        the workbook generator surfaces this as a clean
        "workbook driver not yet implemented" message.
        """

        raise NotImplementedError(
            "run_workbook_case is not implemented yet; the pivot/print driver lands in a later phase"
        )
