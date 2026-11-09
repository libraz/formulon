"""Engine metadata: version, logging, date-serial conversion and the function-metadata seam.

:func:`get_version`, :func:`set_log_min_level`,
:func:`convert_civil_to_serial` and :func:`convert_serial_to_civil` are
process-level calls into the engine. The rest of this module is the
host-injected function-metadata provider seam.

The calculation engine ships no human-readable function documentation:
:func:`get_function_metadata` returns only structural data (name,
arity, availability) and leaves ``signature`` / ``description``
as ``None``. That content is a display concern, so a host supplies its own
metadata document and merges it over the engine result at display time.

This module holds the pure, side-effect-free merge helper and the provider
type shapes. The document contract lives in
``docs/function-metadata-schema.md``. The metadata is display-only: it never
affects formula parsing or evaluation, and the localized names here are not
used for formula-language input localization.
"""

from __future__ import annotations

import datetime
import math
import numbers
from dataclasses import dataclass
from typing import Dict, Mapping, Optional, TypedDict, Union

from . import _structs as S
from ._c import LIB, _require_bool, _sint
from ._c import decode_cstr as _decode_cstr
from ._workbook import FunctionAvailability, FunctionMetadata, LogLevel, _ArgScope, _check

__all__ = [
    "FunctionMetadataEntry",
    "FunctionMetadataLocalized",
    "FunctionMetadataProvider",
    "MergedFunctionMetadata",
    "convert_civil_to_serial",
    "convert_serial_to_civil",
    "get_version",
    "merge_function_metadata",
    "set_log_min_level",
]


def get_version() -> str:
    """Return the version string of the bundled engine.

    Always non-empty.
    """
    return _decode_cstr(LIB.fm_get_version())


def set_log_min_level(level: LogLevel) -> None:
    """Set the engine's minimum structured-log severity.

    This is **process-wide** state, not per :class:`Workbook`, so it is a
    module-level function rather than a method.

    The default is :attr:`LogLevel.OFF`, under which the engine writes
    nothing: an embedded library must not write to the host's stderr
    unless the host asks it to. Raising the threshold to
    :attr:`LogLevel.WARN` or below makes the engine emit one JSON record
    per line on stderr, including the per-cell XLSB downgrade diagnostics.

    There is no log sink: the wasmtime transport cannot bind a C function
    pointer (as for the iterative-progress callback), so this threshold is
    the only control over the log stream.

    Args:
      level: a :class:`LogLevel` member.

    Raises:
      FormulonError: when ``level`` is outside :class:`LogLevel`.
    """
    _check(LIB.fm_set_log_min_level(_sint(level, "level")))


def _check_date1904(date1904: object) -> None:
    _require_bool(date1904, "date1904")


def convert_civil_to_serial(date: Union[datetime.date, datetime.datetime], *, date1904: bool = False) -> float:
    """Convert a calendar date and time to an Excel serial number.

    The engine's own date rules apply, so the result equals
    ``=DATE(y,m,d)+TIME(h,mi,s)``. A ``datetime`` is rounded to the nearest
    second (half up); a plain ``date`` is midnight. The 1900 date system
    accepts only years 1900 to 9999 and the 1904 system only 1904 to 9999.

    Args:
      date: a naive ``datetime.datetime`` or a ``datetime.date``.
      date1904: ``True`` for the 1904 date system.

    Returns:
      The serial number, with the time of day as the fraction.

    Raises:
      TypeError: when ``date`` is not a ``date`` or ``datetime`` (a
        ``datetime.time`` included), or ``date1904`` is not a ``bool``.
      ValueError: when ``date`` carries a time zone, or rounding carries
        it past the largest ``datetime``.
      InvalidArgumentError: when the date lies outside the date system.

    Example::

        convert_civil_to_serial(datetime.datetime(2024, 1, 15, 12, 0))  # 45306.5
    """
    _check_date1904(date1904)
    if isinstance(date, datetime.datetime):
        if date.tzinfo is not None:
            raise ValueError("date must be naive; it is a wall-clock reading")
        rounded = date.replace(microsecond=0)
        if date.microsecond >= 500000:
            try:
                rounded += datetime.timedelta(seconds=1)
            except OverflowError:
                raise ValueError("date rounds past the largest datetime") from None
        civil = rounded
    elif isinstance(date, datetime.date):
        civil = datetime.datetime(date.year, date.month, date.day)
    else:
        raise TypeError(f"date must be a datetime.date or datetime.datetime, got {type(date).__name__}")
    with _ArgScope() as scope:
        civil_ptr = scope.struct(S.CIVIL_TIME)
        S.CIVIL_TIME.pack(
            LIB,
            civil_ptr,
            {
                "year": _sint(civil.year, "year"),
                "month": _sint(civil.month, "month"),
                "day": _sint(civil.day, "day"),
                "hour": _sint(civil.hour, "hour"),
                "minute": _sint(civil.minute, "minute"),
                "second": _sint(civil.second, "second"),
            },
        )
        out = scope.alloc(8)
        _check(LIB.fm_convert_civil_to_serial(civil_ptr, 1 if date1904 else 0, out))
        return LIB.read_f64(out)


def convert_serial_to_civil(serial: float, *, date1904: bool = False) -> datetime.datetime:
    """Convert an Excel serial number to a calendar date and time.

    The fraction is rounded to the nearest whole second (half up), carrying
    into the next day when it reaches 86400. In the 1900 date system serials
    ``[60, 61)`` are Excel's 1900-02-29, which ``datetime`` cannot represent.

    Args:
      serial: a finite serial number.
      date1904: ``True`` for the 1904 date system.

    Returns:
      A naive ``datetime.datetime``.

    Raises:
      TypeError: when ``serial`` is not a number or ``date1904`` is not a ``bool``.
      ValueError: when ``serial`` is NaN or infinite, or lies in ``[60, 61)`` of
        the 1900 date system.
      InvalidArgumentError: when ``serial`` lies outside the date system.

    Example::

        convert_serial_to_civil(45306.5)  # datetime.datetime(2024, 1, 15, 12, 0)
    """
    _check_date1904(date1904)
    if isinstance(serial, bool) or not isinstance(serial, numbers.Real):
        raise TypeError(f"serial must be a number, got {type(serial).__name__}")
    number = float(serial)
    if not math.isfinite(number):
        raise ValueError(f"serial must be finite, got {serial!r}")
    if not date1904 and 60.0 <= number < 61.0:
        raise ValueError("serial 60 is Excel's 1900-02-29 in the 1900 date system, which datetime cannot represent")
    with _ArgScope() as scope:
        out_ptr = scope.struct(S.CIVIL_TIME)
        _check(LIB.fm_convert_serial_to_civil(number, 1 if date1904 else 0, out_ptr))
        civil = S.CIVIL_TIME.unpack(LIB, out_ptr)
    return datetime.datetime(
        civil["year"], civil["month"], civil["day"], civil["hour"], civil["minute"], civil["second"]
    )


class FunctionMetadataLocalized(TypedDict, total=False):
    """Per-locale display overrides inside a :class:`FunctionMetadataEntry`."""

    signature: str
    description: str


class FunctionMetadataEntry(TypedDict, total=False):
    """One host-injected metadata entry.

    Keyed by canonical UPPERCASE function name inside a
    :data:`FunctionMetadataProvider`. Every field is optional.
    """

    #: Default (locale-agnostic) signature template.
    signature: str
    #: Default (locale-agnostic) description.
    description: str
    #: Map of BCP-47 locale tag -> localized display name.
    aliases: Dict[str, str]
    #: Map of BCP-47 locale tag -> per-locale signature/description overrides.
    localized: Dict[str, FunctionMetadataLocalized]


#: A whole host metadata document's ``functions`` map: canonical UPPERCASE
#: function name -> :class:`FunctionMetadataEntry`.
FunctionMetadataProvider = Dict[str, FunctionMetadataEntry]


@dataclass(frozen=True)
class MergedFunctionMetadata:
    """:class:`FunctionMetadata` with a resolved localized display name.

    Returned by :func:`merge_function_metadata`; carries the engine's
    structural fields plus the merged ``signature`` / ``description`` and the
    ``localized_name``.
    """

    name: str
    min_arity: int
    max_arity: Optional[int]
    availability: FunctionAvailability
    signature: Optional[str]
    description: Optional[str]
    #: ``entry.aliases[locale]`` when present, else the canonical ``name``.
    localized_name: str


def _first(*candidates: Optional[str]) -> Optional[str]:
    """Return the first non-``None`` candidate, or ``None``."""
    for value in candidates:
        if value is not None:
            return value
    return None


def merge_function_metadata(
    base: FunctionMetadata,
    entry: Optional[Mapping[str, object]],
    locale: str,
) -> MergedFunctionMetadata:
    """Merge a host-supplied ``entry`` over the engine result ``base``.

    Field precedence (first non-``None`` wins):

    - ``signature``: ``entry["localized"][locale]["signature"]`` ->
      ``entry["signature"]`` -> ``base.signature``
    - ``description``: ``entry["localized"][locale]["description"]`` ->
      ``entry["description"]`` -> ``base.description``
    - ``localized_name``: ``entry["aliases"][locale]`` -> ``base.name``

    Args:
      base: a :class:`FunctionMetadata` from
        :func:`get_function_metadata`.
      entry: the provider's ``functions[NAME]`` mapping, or ``None`` for the
        engine result alone (``signature`` / ``description`` as in ``base``,
        ``localized_name`` the canonical name).
      locale: a BCP-47 display locale tag (e.g. ``"fr-FR"``) matching the
        keys in ``aliases`` / ``localized``.

    Returns:
      A :class:`MergedFunctionMetadata`; with no ``entry`` it is ``base`` with
      ``localized_name`` equal to ``base.name``.
    """
    entry = entry or {}
    localized_map = entry.get("localized") or {}
    localized = localized_map.get(locale) or {}
    signature = _first(
        localized.get("signature"),
        entry.get("signature"),  # type: ignore[arg-type]
        base.signature,
    )
    description = _first(
        localized.get("description"),
        entry.get("description"),  # type: ignore[arg-type]
        base.description,
    )
    aliases = entry.get("aliases") or {}
    localized_name = aliases.get(locale, base.name)
    return MergedFunctionMetadata(
        name=base.name,
        min_arity=base.min_arity,
        max_arity=base.max_arity,
        availability=base.availability,
        signature=signature,
        description=description,
        localized_name=localized_name,
    )
