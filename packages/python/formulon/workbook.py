# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Pythonic wrapper around the Formulon C ABI (WASM transport).

Public surface:
  * :class:`Workbook` -- context-manager around an ``fm_workbook_t*`` handle.
  * :class:`Value` -- frozen dataclass mirroring ``fm_value_t``.
  * :class:`FormulonError` -- exception raised on host-side failures.
  * :class:`Cell`, :class:`DefinedName`, :class:`Table` -- iteration tuples.

Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) are reported as
``Value`` instances with ``kind == ValueKind.ERROR``; only host-side
problems (NULL handle, parse failure during ``Workbook.load``, OOM)
raise :class:`FormulonError`.

The transport is WebAssembly (``formulon_capi.wasm``) loaded via
``wasmtime``; see :mod:`formulon._c` for the low-level memory and call
plumbing.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass
from typing import Iterator, NamedTuple, Optional, Union

from ._c import LIB, ValueKind, fm_value_t_size

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

    The message and context are copied out of the thread-local WASM
    buffers eagerly, so they remain valid even after subsequent WASM
    calls overwrite those buffers.
    """

    def __init__(self, status: int, *, op: str = "") -> None:
        self.status = int(status)
        self.status_name = LIB.read_cstr(LIB.fm_status_string(self.status))
        self.message = LIB.read_cstr(LIB.fm_last_error_message())
        self.context = LIB.read_cstr(LIB.fm_last_error_context())
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
    def _from_wasm(cls, ptr: int) -> "Value":
        """Decode an ``fm_value_t`` stored at WASM offset ``ptr``.

        Layout (wasm32): int32 kind | 4 pad | union (8 bytes).

        Reserved kinds (ARRAY / REF / LAMBDA) currently surface as bare
        ``Value(kind=...)`` with no payload; they are reserved in the C
        ABI for a later bundle.
        """
        buf = LIB.read_bytes(ptr, fm_value_t_size)
        kind_int = struct.unpack_from("<i", buf, 0)[0]
        kind = ValueKind(kind_int)
        if kind is ValueKind.NUMBER:
            number = struct.unpack_from("<d", buf, 8)[0]
            return cls(kind=kind, number=number)
        if kind is ValueKind.BOOL:
            boolean = struct.unpack_from("<i", buf, 8)[0]
            return cls(kind=kind, boolean=bool(boolean))
        if kind is ValueKind.TEXT:
            text_ptr = struct.unpack_from("<I", buf, 8)[0]
            return cls(kind=kind, text=LIB.read_cstr(text_ptr))
        if kind is ValueKind.ERROR:
            error_code = struct.unpack_from("<i", buf, 8)[0]
            return cls(kind=kind, error_code=int(error_code))
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


def _alloc_out_ptr() -> int:
    """Allocate a 4-byte WASM scratch slot for an out-i32 / out-ptr."""
    ptr = LIB.alloc(4)
    LIB.write_bytes(ptr, b"\x00\x00\x00\x00")
    return ptr


class Workbook:
    """Owns an ``fm_workbook_t*`` handle (a WASM i32 offset).

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
      Each ``Workbook`` instance is single-threaded. The underlying
      wasmtime store is serialised process-wide by an internal lock in
      :mod:`formulon._c`, so cross-thread misuse cannot corrupt WASM
      state -- it just blocks. The parallel recalc scheduler degrades
      to serial under WASM (single-threaded build).
    """

    __slots__ = ("_handle",)

    def __init__(self) -> None:
        # 0 == NULL pointer. Use the factory methods to allocate.
        self._handle: int = 0

    # -- Factories ---------------------------------------------------------
    @classmethod
    def create_default(cls) -> "Workbook":
        """Build a default workbook (a single sheet named ``Sheet1``)."""
        wb = cls()
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_create(out_ptr)
            _check(status, "fm_workbook_create")
            wb._handle = LIB.read_u32(out_ptr)
        finally:
            LIB.free(out_ptr)
        return wb

    @classmethod
    def create_empty(cls) -> "Workbook":
        """Build an empty workbook (no sheets).

        Saving an empty workbook fails just as Excel rejects sheet-less
        archives; callers are expected to add at least one sheet before
        :meth:`save`.
        """
        wb = cls()
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_create_empty(out_ptr)
            _check(status, "fm_workbook_create_empty")
            wb._handle = LIB.read_u32(out_ptr)
        finally:
            LIB.free(out_ptr)
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
        buf = bytes(data)
        wb = cls()
        data_ptr = LIB.alloc_bytes(buf) if len(buf) > 0 else 0
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_load(data_ptr, len(buf), out_ptr)
            _check(status, "fm_workbook_load")
            wb._handle = LIB.read_u32(out_ptr)
        finally:
            LIB.free(out_ptr)
            if data_ptr:
                LIB.free(data_ptr)
        return wb

    # -- Lifecycle ---------------------------------------------------------
    def close(self) -> None:
        """Release the underlying handle. Idempotent."""
        if self._handle:
            LIB.fm_workbook_destroy(self._handle)
            self._handle = 0

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # type: ignore[no-untyped-def]
        self.close()

    def __del__(self) -> None:  # pragma: no cover -- GC ordering best-effort
        try:
            self.close()
        except Exception:  # noqa: BLE001 -- destructors must not raise
            pass

    @property
    def is_valid(self) -> bool:
        """``True`` while the handle is still alive."""
        return self._handle != 0

    def _require(self) -> int:
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
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_sheet_name(h, index, out_ptr)
            _check(status, "fm_workbook_sheet_name")
            return LIB.read_cstr(LIB.read_u32(out_ptr))
        finally:
            LIB.free(out_ptr)

    def add_sheet(self, name: str) -> None:
        """Append a new sheet with the given UTF-8 display name."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            status = LIB.fm_workbook_add_sheet(h, name_ptr)
            _check(status, "fm_workbook_add_sheet")
        finally:
            LIB.free(name_ptr)

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
        text_ptr, _ = LIB.alloc_utf8(value)
        try:
            status = LIB.fm_workbook_set_text(h, sheet, row, col, text_ptr)
            _check(status, "fm_workbook_set_text")
        finally:
            LIB.free(text_ptr)

    def set_blank(self, sheet: int, row: int, col: int) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_blank(h, sheet, row, col)
        _check(status, "fm_workbook_set_blank")

    def set_formula(self, sheet: int, row: int, col: int, formula: str) -> None:
        h = self._require()
        formula_ptr, _ = LIB.alloc_utf8(formula)
        try:
            status = LIB.fm_workbook_set_formula(h, sheet, row, col, formula_ptr)
            _check(status, "fm_workbook_set_formula")
        finally:
            LIB.free(formula_ptr)

    # -- Cell read ---------------------------------------------------------
    def get_value(self, sheet: int, row: int, col: int) -> Value:
        """Read the cached value at ``(sheet, row, col)``.

        Reflects the most recent recalc; mutations made after the last
        :meth:`recalc` are not visible to dependents until you call
        :meth:`recalc` again.
        """
        h = self._require()
        value_ptr = LIB.alloc(fm_value_t_size)
        try:
            status = LIB.fm_workbook_get_value(h, sheet, row, col, value_ptr)
            _check(status, "fm_workbook_get_value")
            return Value._from_wasm(value_ptr)
        finally:
            LIB.free(value_ptr)

    # -- Recalc ------------------------------------------------------------
    def recalc(self) -> None:
        """Drive a full incremental recalc.

        Under WASM this is always serial -- the parallel scheduler is
        unavailable because the standalone reactor build is compiled
        without ``-pthread``. Throughput is therefore lower than the
        native CLI but result fidelity is identical.
        """
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

        The returned bytes are an independent copy; the underlying WASM
        buffer is freed before this method returns.
        """
        h = self._require()
        out_ptr_ptr = LIB.alloc(4)
        out_len_ptr = LIB.alloc(4)
        LIB.write_bytes(out_ptr_ptr, b"\x00\x00\x00\x00")
        LIB.write_bytes(out_len_ptr, b"\x00\x00\x00\x00")
        try:
            status = LIB.fm_workbook_save(h, out_ptr_ptr, out_len_ptr)
            _check(status, "fm_workbook_save")
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            try:
                if data_len == 0 or data_ptr == 0:
                    return b""
                return LIB.read_bytes(data_ptr, data_len)
            finally:
                if data_ptr:
                    LIB.fm_buffer_free(data_ptr)
        finally:
            LIB.free(out_ptr_ptr)
            LIB.free(out_len_ptr)

    # -- Iteration ---------------------------------------------------------
    def iter_cells(self, sheet: int) -> Iterator[Cell]:
        """Iterate over every stored cell on ``sheet``.

        Iteration order is implementation-defined but stable for an
        unmutated workbook (sorted by ``(row, col)`` ascending per the C
        ABI). Borrowed text pointers are decoded eagerly per yielded
        item, so consumers may freely mutate the workbook between yields.
        """
        h = self._require()
        out_count = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_cell_count(h, sheet, out_count)
            _check(status, "fm_workbook_cell_count")
            n = LIB.read_u32(out_count)
        finally:
            LIB.free(out_count)
        # Scratch: row(4) + col(4) + formula_ptr(4) + value(16) = 28 bytes.
        # Allocate them as separate slots so the ABI's
        # individual out-parameter contract is respected.
        for i in range(n):
            row_ptr = LIB.alloc(4)
            col_ptr = LIB.alloc(4)
            formula_ptr = LIB.alloc(4)
            value_ptr = LIB.alloc(fm_value_t_size)
            LIB.write_bytes(row_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(col_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(formula_ptr, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_cell_at(
                    h, sheet, i, row_ptr, col_ptr, formula_ptr, value_ptr
                )
                _check(status, "fm_workbook_cell_at")
                row = LIB.read_u32(row_ptr)
                col = LIB.read_u32(col_ptr)
                formula_addr = LIB.read_u32(formula_ptr)
                formula = LIB.read_cstr(formula_addr) if formula_addr else None
                value = Value._from_wasm(value_ptr)
            finally:
                LIB.free(row_ptr)
                LIB.free(col_ptr)
                LIB.free(formula_ptr)
                LIB.free(value_ptr)
            yield Cell(row=row, col=col, formula=formula, value=value)

    def iter_defined_names(self) -> Iterator[DefinedName]:
        """Iterate over every defined name in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_defined_name_count(h))
        for i in range(n):
            name_ptr = LIB.alloc(4)
            formula_ptr = LIB.alloc(4)
            LIB.write_bytes(name_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(formula_ptr, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_defined_name_at(h, i, name_ptr, formula_ptr)
                _check(status, "fm_workbook_defined_name_at")
                name = LIB.read_cstr(LIB.read_u32(name_ptr))
                formula = LIB.read_cstr(LIB.read_u32(formula_ptr))
            finally:
                LIB.free(name_ptr)
                LIB.free(formula_ptr)
            yield DefinedName(name=name, formula=formula)

    def iter_tables(self) -> Iterator[Table]:
        """Iterate over every table in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_table_count(h))
        for i in range(n):
            name_ptr = LIB.alloc(4)
            display_ptr = LIB.alloc(4)
            ref_ptr = LIB.alloc(4)
            sheet_ptr = LIB.alloc(4)
            for p in (name_ptr, display_ptr, ref_ptr, sheet_ptr):
                LIB.write_bytes(p, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_table_at(
                    h, i, name_ptr, display_ptr, ref_ptr, sheet_ptr
                )
                _check(status, "fm_workbook_table_at")
                name = LIB.read_cstr(LIB.read_u32(name_ptr))
                display = LIB.read_cstr(LIB.read_u32(display_ptr))
                ref = LIB.read_cstr(LIB.read_u32(ref_ptr))
                sheet = LIB.read_u32(sheet_ptr)
            finally:
                LIB.free(name_ptr)
                LIB.free(display_ptr)
                LIB.free(ref_ptr)
                LIB.free(sheet_ptr)
            yield Table(name=name, display_name=display, ref=ref, sheet_index=sheet)

    def iter_passthrough(self) -> Iterator[PassthroughPart]:
        """Iterate over passthrough OOXML part paths."""
        h = self._require()
        n = int(LIB.fm_workbook_passthrough_count(h))
        for i in range(n):
            path_ptr = LIB.alloc(4)
            LIB.write_bytes(path_ptr, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_passthrough_at(h, i, path_ptr)
                _check(status, "fm_workbook_passthrough_at")
                path = LIB.read_cstr(LIB.read_u32(path_ptr))
            finally:
                LIB.free(path_ptr)
            yield PassthroughPart(path=path)
