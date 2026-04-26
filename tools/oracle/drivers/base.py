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


@dataclass
class CaseResult:
    """Normalised observation for one case.

    Mirrors the on-wire JSON shape the generator will emit. `kind` is one
    of {'blank','number','bool','text','error','array'}. `array_shape` is
    populated only for spill results; the initial driver flattens arrays
    to a single value and signals `array_shape is not None` so callers
    can log.
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
    def __enter__(self) -> "OracleDriver":
        ...

    @abc.abstractmethod
    def __exit__(self, exc_type, exc, tb) -> None:
        ...

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
        (mapping A1 -> `{kind, value, ...}` record). The driver does not
        normalise; it trusts upstream `case_schema` to have done so.
        """
