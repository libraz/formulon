# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Pythonic wrapper around the Formulon C ABI.

Public surface:
  * :class:`Workbook` -- context-manager around an ``fm_workbook_t*`` handle.
  * :class:`Value` -- frozen dataclass mirroring ``fm_value_t``.
  * :class:`FormulonError` -- exception raised on host-side failures.
  * :class:`Cell`, :class:`DefinedName`, :class:`Table` -- iteration tuples.

Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) are reported as
``Value`` instances with ``kind == ValueKind.ERROR``; only host-side
problems (NULL handle, parse failure during ``Workbook.load``, OOM)
raise :class:`FormulonError`.
"""

from __future__ import annotations

import ctypes
from dataclasses import dataclass
from typing import Iterator, NamedTuple, Optional, Union

from . import _c
from ._c import LIB, ValueKind, decode_cstr, fm_value_t, fm_workbook_t_p

__all__ = [
    "Cell",
    "DefinedName",
    "FormulonError",
    "PassthroughPart",
    "Table",
    "Value",
    "ValueKind",
    "Workbook",
]


# ---------------------------------------------------------------------------
# Error type
# ---------------------------------------------------------------------------


class FormulonError(Exception):
    """Raised for host-side failures of the Formulon C ABI.

    Attributes:
      status: numeric ``fm_status_t`` (matches
        ``formulon::FormulonErrorCode`` ordinals).
      status_name: symbolic name (e.g. ``"kInvalidArgument"``) returned
        by ``fm_status_string``.
      message: thread-local diagnostic message captured from
        ``fm_last_error_message`` at the time the failure was detected.
      context: optional context string captured from
        ``fm_last_error_context``.

    The message and context are copied out of the thread-local C buffers
    eagerly, so they remain valid even after subsequent C calls overwrite
    those buffers.
    """

    def __init__(self, status: int, *, op: str = "") -> None:
        self.status = int(status)
        self.status_name = decode_cstr(LIB.fm_status_string(self.status))
        self.message = decode_cstr(LIB.fm_last_error_message())
        self.context = decode_cstr(LIB.fm_last_error_context())
        prefix = f"{op}: " if op else ""
        text = f"{prefix}{self.status_name} ({self.status})"
        if self.message:
            text += f": {self.message}"
        if self.context:
            text += f" [{self.context}]"
        super().__init__(text)


def _check(status: int, op: str) -> None:
    """Raise :class:`FormulonError` if ``status`` is non-zero."""
    if status != 0:
        raise FormulonError(status, op=op)


# ---------------------------------------------------------------------------
# Value POD
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Value:
    """Pythonic mirror of ``fm_value_t``.

    Exactly one of ``number`` / ``boolean`` / ``text`` / ``error_code`` is
    populated according to ``kind``. The constructor enforces the
    invariant.
    """

    kind: ValueKind
    number: Optional[float] = None
    boolean: Optional[bool] = None
    text: Optional[str] = None
    error_code: Optional[int] = None

    @classmethod
    def _from_c(cls, raw: fm_value_t) -> "Value":
        """Translate a ``fm_value_t`` into a :class:`Value`.

        Text payloads are eagerly decoded into Python ``str`` so the
        result is independent of the workbook's text storage lifetime.
        Reserved kinds (ARRAY / REF / LAMBDA) currently surface as bare
        ``Value(kind=...)`` with no payload; they are reserved in the C
        ABI for a later bundle.
        """
        kind = ValueKind(raw.kind)
        if kind is ValueKind.NUMBER:
            return cls(kind=kind, number=float(raw.u.number))
        if kind is ValueKind.BOOL:
            return cls(kind=kind, boolean=bool(raw.u.boolean))
        if kind is ValueKind.TEXT:
            text = decode_cstr(raw.u.text)
            return cls(kind=kind, text=text)
        if kind is ValueKind.ERROR:
            return cls(kind=kind, error_code=int(raw.u.error_code))
        # BLANK / ARRAY / REF / LAMBDA: no scalar payload.
        return cls(kind=kind)

    def to_python(self) -> Union[None, float, bool, str, "Value"]:
        """Convert to a native Python type.

        - ``BLANK`` -> ``None``
        - ``NUMBER`` -> ``float``
        - ``BOOL`` -> ``bool``
        - ``TEXT`` -> ``str``
        - ``ERROR`` -> the :class:`Value` itself (Excel cell errors are
          values, not Python exceptions; callers inspect ``error_code``)
        - ``ARRAY`` / ``REF`` / ``LAMBDA`` -> the :class:`Value` itself
          (passthrough, payload not yet exposed by the C ABI)
        """
        if self.kind is ValueKind.BLANK:
            return None
        if self.kind is ValueKind.NUMBER:
            return self.number  # type: ignore[return-value]
        if self.kind is ValueKind.BOOL:
            return self.boolean  # type: ignore[return-value]
        if self.kind is ValueKind.TEXT:
            return self.text  # type: ignore[return-value]
        # ERROR / ARRAY / REF / LAMBDA: surface the wrapping Value.
        return self


# ---------------------------------------------------------------------------
# Iteration tuples
# ---------------------------------------------------------------------------


class Cell(NamedTuple):
    """One row of ``Workbook.iter_cells()``."""

    row: int
    col: int
    formula: Optional[str]
    value: Value


class DefinedName(NamedTuple):
    """One row of ``Workbook.iter_defined_names()``."""

    name: str
    formula: str


class Table(NamedTuple):
    """One row of ``Workbook.iter_tables()``."""

    name: str
    display_name: str
    ref: str
    sheet_index: int


class PassthroughPart(NamedTuple):
    """One row of ``Workbook.iter_passthrough()``."""

    path: str


# ---------------------------------------------------------------------------
# Workbook handle
# ---------------------------------------------------------------------------


def _encode_utf8(s: str) -> bytes:
    """Encode a string for the C ABI; reject non-string inputs early."""
    if not isinstance(s, str):
        raise TypeError(f"expected str, got {type(s).__name__}")
    return s.encode("utf-8")


class Workbook:
    """Owns an ``fm_workbook_t*`` handle.

    Construction:
      Use one of :meth:`create_default`, :meth:`create_empty`, or
      :meth:`load`. Direct ``Workbook()`` instantiation is supported as a
      stub for subclassing but does not allocate a handle.

    Lifetime:
      Always release via :meth:`close`, ideally through ``with`` syntax::

          with Workbook.create_default() as wb:
              wb.set_number(0, 0, 0, 42.0)
              wb.recalc()
              print(wb.get_value(0, 0, 0).to_python())

      :meth:`close` is idempotent. The destructor calls ``close`` as a
      defensive net, but relying on garbage-collection ordering is
      unsafe; prefer the context-manager form.

    Threading:
      Each ``Workbook`` instance is single-threaded. Concurrent calls on
      the same handle from multiple threads are undefined behaviour per
      the C ABI contract.
    """

    __slots__ = ("_handle",)

    def __init__(self) -> None:
        # Default-constructed instances start with a NULL handle. Use the
        # factory methods to actually allocate one. We deliberately do NOT
        # raise here so that subclasses / mocking are tractable.
        self._handle: fm_workbook_t_p = fm_workbook_t_p(0)

    # -- Factories ---------------------------------------------------------
    @classmethod
    def create_default(cls) -> "Workbook":
        """Build a default workbook (a single sheet named ``Sheet1``)."""
        wb = cls()
        out = fm_workbook_t_p(0)
        status = LIB.fm_workbook_create(ctypes.byref(out))
        _check(status, "fm_workbook_create")
        wb._handle = out
        return wb

    @classmethod
    def create_empty(cls) -> "Workbook":
        """Build an empty workbook (no sheets).

        Saving an empty workbook fails just as Excel rejects sheet-less
        archives; callers are expected to add at least one sheet before
        :meth:`save`.
        """
        wb = cls()
        out = fm_workbook_t_p(0)
        status = LIB.fm_workbook_create_empty(ctypes.byref(out))
        _check(status, "fm_workbook_create_empty")
        wb._handle = out
        return wb

    @classmethod
    def load(cls, data: Union[bytes, bytearray, memoryview]) -> "Workbook":
        """Load a workbook from an in-memory ``.xlsx`` byte buffer.

        Args:
          data: the raw OOXML archive bytes. Accepts any object that
            satisfies the buffer protocol.

        Raises:
          FormulonError: when the input cannot be parsed as a valid
            OOXML archive.
        """
        if not isinstance(data, (bytes, bytearray, memoryview)):
            raise TypeError(
                f"Workbook.load: expected bytes-like, got {type(data).__name__}"
            )
        # ctypes can convert bytes / bytearray directly. memoryview needs
        # to be flattened to bytes first because POINTER(c_uint8) does not
        # accept arbitrary buffer-protocol objects portably.
        buf = bytes(data)
        wb = cls()
        out = fm_workbook_t_p(0)
        # cast bytes to POINTER(c_uint8) via from_buffer_copy on a uint8 array.
        if len(buf) > 0:
            arr_t = ctypes.c_uint8 * len(buf)
            arr = arr_t.from_buffer_copy(buf)
            ptr = ctypes.cast(arr, ctypes.POINTER(ctypes.c_uint8))
        else:
            ptr = ctypes.cast(ctypes.c_void_p(0), ctypes.POINTER(ctypes.c_uint8))
        status = LIB.fm_workbook_load(ptr, len(buf), ctypes.byref(out))
        _check(status, "fm_workbook_load")
        wb._handle = out
        return wb

    # -- Lifecycle ---------------------------------------------------------
    def close(self) -> None:
        """Release the underlying handle. Idempotent."""
        h = self._handle
        if h and h.value:
            LIB.fm_workbook_destroy(h)
            self._handle = fm_workbook_t_p(0)

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # type: ignore[no-untyped-def]
        self.close()

    def __del__(self) -> None:  # pragma: no cover -- GC ordering best-effort
        # Defensive; prefer explicit close() via context manager.
        try:
            self.close()
        except Exception:  # noqa: BLE001 -- destructors must not raise
            pass

    @property
    def is_valid(self) -> bool:
        """``True`` while the handle is still alive."""
        return bool(self._handle and self._handle.value)

    def _require(self) -> fm_workbook_t_p:
        if not self.is_valid:
            raise FormulonError(
                # 7000-band: bindings / C API. 7000 is kBindingNullPointer
                # in the live error_codes table.
                7000,
                op="Workbook (handle is NULL or already closed)",
            )
        return self._handle

    # -- Sheets ------------------------------------------------------------
    def sheet_count(self) -> int:
        """Number of sheets in the workbook."""
        h = self._require()
        return int(LIB.fm_workbook_sheet_count(h))

    def sheet_name(self, index: int) -> str:
        """Display name (UTF-8) of the sheet at ``index``."""
        h = self._require()
        out = ctypes.c_char_p()
        status = LIB.fm_workbook_sheet_name(h, index, ctypes.byref(out))
        _check(status, "fm_workbook_sheet_name")
        # Decode immediately; pointer borrows from the workbook handle.
        return decode_cstr(out.value)

    def add_sheet(self, name: str) -> None:
        """Append a new sheet with the given UTF-8 display name."""
        h = self._require()
        status = LIB.fm_workbook_add_sheet(h, _encode_utf8(name))
        _check(status, "fm_workbook_add_sheet")

    # -- Cell mutation -----------------------------------------------------
    def set_number(self, sheet: int, row: int, col: int, value: float) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_number(h, sheet, row, col, float(value))
        _check(status, "fm_workbook_set_number")

    def set_bool(self, sheet: int, row: int, col: int, value: bool) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_bool(h, sheet, row, col, 1 if value else 0)
        _check(status, "fm_workbook_set_bool")

    def set_text(self, sheet: int, row: int, col: int, value: str) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_text(h, sheet, row, col, _encode_utf8(value))
        _check(status, "fm_workbook_set_text")

    def set_blank(self, sheet: int, row: int, col: int) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_blank(h, sheet, row, col)
        _check(status, "fm_workbook_set_blank")

    def set_formula(self, sheet: int, row: int, col: int, formula: str) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_formula(h, sheet, row, col, _encode_utf8(formula))
        _check(status, "fm_workbook_set_formula")

    # -- Cell read ---------------------------------------------------------
    def get_value(self, sheet: int, row: int, col: int) -> Value:
        """Read the cached value at ``(sheet, row, col)``.

        Reflects the most recent recalc; mutations made after the last
        :meth:`recalc` are not visible to dependents until you call
        :meth:`recalc` again.
        """
        h = self._require()
        raw = fm_value_t()
        status = LIB.fm_workbook_get_value(h, sheet, row, col, ctypes.byref(raw))
        _check(status, "fm_workbook_get_value")
        return Value._from_c(raw)

    # -- Recalc ------------------------------------------------------------
    def recalc(self) -> None:
        """Drive a full incremental recalc."""
        h = self._require()
        status = LIB.fm_workbook_recalc(h)
        _check(status, "fm_workbook_recalc")

    def set_iterative(
        self, enabled: bool, max_iterations: int, max_change: float
    ) -> None:
        """Configure iterative calculation."""
        h = self._require()
        status = LIB.fm_workbook_set_iterative(
            h, 1 if enabled else 0, int(max_iterations), float(max_change)
        )
        _check(status, "fm_workbook_set_iterative")

    # -- Save --------------------------------------------------------------
    def save(self) -> bytes:
        """Serialise the workbook to an in-memory ``.xlsx`` byte stream.

        The returned bytes are an independent copy; the underlying C
        buffer is freed before this method returns.
        """
        h = self._require()
        out_bytes = ctypes.POINTER(ctypes.c_uint8)()
        out_len = ctypes.c_size_t(0)
        status = LIB.fm_workbook_save(h, ctypes.byref(out_bytes), ctypes.byref(out_len))
        _check(status, "fm_workbook_save")
        try:
            n = int(out_len.value)
            if n == 0 or not out_bytes:
                return b""
            # Copy out to a pure-Python bytes object before freeing.
            return ctypes.string_at(out_bytes, n)
        finally:
            if out_bytes:
                LIB.fm_buffer_free(out_bytes)

    # -- Iteration ---------------------------------------------------------
    def iter_cells(self, sheet: int) -> Iterator[Cell]:
        """Iterate over every stored cell on ``sheet``.

        Iteration order is implementation-defined but stable for an
        unmutated workbook (sorted by ``(row, col)`` ascending per the C
        ABI). Borrowed text pointers are decoded eagerly per yielded
        item, so consumers may freely mutate the workbook between yields.
        """
        h = self._require()
        out_count = ctypes.c_size_t(0)
        status = LIB.fm_workbook_cell_count(h, sheet, ctypes.byref(out_count))
        _check(status, "fm_workbook_cell_count")
        n = int(out_count.value)
        for i in range(n):
            row = ctypes.c_uint32(0)
            col = ctypes.c_uint32(0)
            formula = ctypes.c_char_p()
            value = fm_value_t()
            status = LIB.fm_workbook_cell_at(
                h,
                sheet,
                i,
                ctypes.byref(row),
                ctypes.byref(col),
                ctypes.byref(formula),
                ctypes.byref(value),
            )
            _check(status, "fm_workbook_cell_at")
            yield Cell(
                row=int(row.value),
                col=int(col.value),
                formula=decode_cstr(formula.value) if formula.value else None,
                value=Value._from_c(value),
            )

    def iter_defined_names(self) -> Iterator[DefinedName]:
        """Iterate over every defined name in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_defined_name_count(h))
        for i in range(n):
            name = ctypes.c_char_p()
            formula = ctypes.c_char_p()
            status = LIB.fm_workbook_defined_name_at(
                h, i, ctypes.byref(name), ctypes.byref(formula)
            )
            _check(status, "fm_workbook_defined_name_at")
            yield DefinedName(
                name=decode_cstr(name.value),
                formula=decode_cstr(formula.value),
            )

    def iter_tables(self) -> Iterator[Table]:
        """Iterate over every table in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_table_count(h))
        for i in range(n):
            name = ctypes.c_char_p()
            display = ctypes.c_char_p()
            ref = ctypes.c_char_p()
            sheet = ctypes.c_size_t(0)
            status = LIB.fm_workbook_table_at(
                h,
                i,
                ctypes.byref(name),
                ctypes.byref(display),
                ctypes.byref(ref),
                ctypes.byref(sheet),
            )
            _check(status, "fm_workbook_table_at")
            yield Table(
                name=decode_cstr(name.value),
                display_name=decode_cstr(display.value),
                ref=decode_cstr(ref.value),
                sheet_index=int(sheet.value),
            )

    def iter_passthrough(self) -> Iterator[PassthroughPart]:
        """Iterate over passthrough OOXML part paths."""
        h = self._require()
        n = int(LIB.fm_workbook_passthrough_count(h))
        for i in range(n):
            path = ctypes.c_char_p()
            status = LIB.fm_workbook_passthrough_at(h, i, ctypes.byref(path))
            _check(status, "fm_workbook_passthrough_at")
            yield PassthroughPart(path=decode_cstr(path.value))


# Re-export the ValueKind enum here too so callers don't need to dip into
# the private _c module.
_ = _c  # silence unused-import warnings on minimal type checkers
