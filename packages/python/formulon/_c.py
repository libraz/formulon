"""Low-level WASM binding for the Formulon C ABI.

Public users should import the ``Workbook`` / ``Value`` / ``FormulonError``
symbols re-exported by the top-level :mod:`formulon` package instead.

Architecture
------------

The binding loads ``formulon_capi.wasm`` (a standalone reactor-style
WebAssembly module that exports the ``fm_*`` C ABI from
``src/c_api/formulon_c.h``) via the ``wasmtime`` runtime. A single
module instance is created lazily at first use and shared across every
:class:`formulon.Workbook` in the process; each ``Workbook`` instance
owns an opaque ``fm_workbook_t*`` (i32 in WASM) handle.

Pointers are 32-bit offsets into the WASM linear memory. Strings cross
the boundary as ``malloc``-allocated UTF-8 NUL-terminated buffers in
that memory; the caller is responsible for ``free``-ing them after the
WASM function returns. Borrowed return pointers (e.g. text from
``fm_workbook_get_value``) are read directly from WASM memory and
decoded eagerly into Python ``str`` so the result outlives any
subsequent WASM mutation.

The ``fm_value_t`` POD has the wasm32 layout::

    offset 0: int32  kind
    offset 4: int32  (padding for 8-byte alignment of the union)
    offset 8: union  { double number ; int32 boolean ; int32 error_code ; ptr text }
    total: 16 bytes

The union member is selected by ``kind``; reading any other member is
undefined per the C ABI contract.
"""

from __future__ import annotations

import struct
import threading
from enum import IntEnum
from pathlib import Path
from typing import Optional, Tuple

import wasmtime

__all__ = [
    "LIB",
    "FormulonError",
    "ValueKind",
    "decode_cstr",
    "fm_value_t_size",
]

# fm_value_t layout: int32 kind + 4 pad + 8 union = 16 bytes.
fm_value_t_size = 16

# `formulon::FormulonErrorCode::kInternalError` (src/utils/error.h, the
# "General" 0-999 band). Used only for a host-side failure the WASM trampoline
# below observes but the C ABI itself never reported a status for -- a
# wasmtime Trap, which aborts the call before any `fm_status_t` is returned.
_STATUS_INTERNAL_ERROR = 8


def _uint(value: int, name: str, bits: int = 32) -> int:
    """Validate ``value`` as an unsigned WASM argument of ``bits`` width.

    Every integer the C ABI takes by value crosses into WebAssembly as an
    ``i32``, and ``wasmtime`` marshals a Python ``int`` into that slot
    modulo 2**32 without complaint. An index the caller never intended --
    ``2**32 + 5``, or a negative produced by an arithmetic slip -- would
    therefore arrive on the engine side as a perfectly plausible
    coordinate (``5``) and silently overwrite an unrelated cell. The
    engine's own grid bounds cannot catch that: they see only the wrapped
    value.

    The same applies to an integer packed into a struct the ABI reads
    through a pointer, with a different failure: ``struct.pack`` rejects
    the value outright and raises ``struct.error``, which is not the
    exception this API documents and not one a caller catches. Both routes
    go through here so the diagnosis is the same either way.

    Args:
      value: the caller-supplied index, count, or length.
      name: the C ABI parameter or struct field name, used in the message.
      bits: width of the destination (32 for ``uint32_t`` / ``size_t``,
        16 and 8 for the narrower struct fields, which overflow well
        before an ``i32`` does).

    Returns:
      ``value`` as an ``int``, unchanged.

    Raises:
      ValueError: when ``value`` is negative or does not fit ``bits``.
    """
    ivalue = int(value)
    if ivalue < 0 or ivalue >= (1 << bits):
        raise ValueError(f"formulon: {name} out of range for uint{bits}: {ivalue}")
    return ivalue


def _sint(value: int, name: str, bits: int = 32) -> int:
    """Validate ``value`` as a signed WASM argument of ``bits`` width.

    Signed counterpart of :func:`_uint`, for the ``int32_t`` parameters
    and struct fields (enum selectors, iteration caps, the ``-1``
    sentinels of the pivot date-group and show-as APIs) where a negative
    value is meaningful.

    Raises:
      ValueError: when ``value`` does not fit a signed ``bits``-wide int.
    """
    ivalue = int(value)
    limit = 1 << (bits - 1)
    if ivalue < -limit or ivalue >= limit:
        raise ValueError(f"formulon: {name} out of range for int{bits}: {ivalue}")
    return ivalue


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

    Defined here rather than in :mod:`formulon.workbook` (which
    re-exports it) so the low-level export trampoline below can raise it
    directly on a trap, without importing back into the module that
    imports ``LIB`` from here.
    """

    def __init__(
        self,
        status: int,
        *,
        op: str = "",
        _diagnostic_override: Optional[Tuple[str, str]] = None,
        _status_name_override: Optional[str] = None,
    ) -> None:
        self.status = int(status)
        if _status_name_override is not None:
            # A trap leaves the module instance poisoned (see `_wrapped`
            # below); calling back into it for `fm_status_string` would
            # retry the very call that just failed, so this path never
            # touches WASM memory.
            self.status_name = _status_name_override
            self.message, self.context = _diagnostic_override or ("", "")
        else:
            # Take the Python-side snapshot before calling fm_status_string.
            # The latter is a non-status pointer-returning export and must
            # not become a diagnostic read or overwrite the pending
            # snapshot. An override still drains a pending C diagnostic so
            # it cannot leak into a later exception on this Python thread.
            message, context = LIB.last_diagnostic(self.status)
            if _diagnostic_override is not None:
                message, context = _diagnostic_override
            self.status_name = LIB.read_cstr(LIB.fm_status_string(_sint(self.status, "status")))
            self.message, self.context = message, context
        prefix = f"{op}: " if op else ""
        text = f"{prefix}{self.status_name} ({self.status})"
        if self.message:
            text += f": {self.message}"
        if self.context:
            text += f" [{self.context}]"
        super().__init__(text)


# Generated from the current working-tree C header and WASM export manifest:
# {src/c_api/formulon_c.h declarations returning fm_status_t} intersection
# {tools/wasm/capi_exports.txt}. Keep this literal so the drift checker can
# validate the binding without importing this module (or wasmtime).
_STATUS_RETURNING_EXPORT_NAMES = (
    "fm_cell_get_xf_index",
    "fm_cell_nodes_at",
    "fm_cell_set_xf_index",
    "fm_cf_results_cell_at",
    "fm_cf_results_match_at",
    "fm_function_canonicalize",
    "fm_function_localize",
    "fm_function_metadata",
    "fm_function_name_at",
    "fm_pagination_horizontal_break_at",
    "fm_pagination_print_area_at",
    "fm_pagination_vertical_break_at",
    "fm_pivot_cells_at",
    "fm_pivot_cells_bounds",
    "fm_set_log_min_level",
    "fm_set_log_sink",
    "fm_sheet_add_hyperlink",
    "fm_sheet_add_merge",
    "fm_sheet_add_validation",
    "fm_sheet_cf_add_rule",
    "fm_sheet_cf_clear",
    "fm_sheet_cf_count",
    "fm_sheet_cf_get_at",
    "fm_sheet_cf_remove_at",
    "fm_sheet_add_col_break",
    "fm_sheet_add_row_break",
    "fm_sheet_clear_breaks",
    "fm_sheet_clear_hyperlinks",
    "fm_sheet_clear_merges",
    "fm_sheet_clear_validations",
    "fm_sheet_col_break_at",
    "fm_sheet_get_auto_filter_xml",
    "fm_sheet_get_column",
    "fm_sheet_get_column_count",
    "fm_sheet_get_comment_at",
    "fm_sheet_get_comment_at_index",
    "fm_sheet_get_comment_count",
    "fm_sheet_get_header_footer_xml",
    "fm_sheet_get_hyperlink_at",
    "fm_sheet_get_hyperlink_count",
    "fm_sheet_get_merge_at",
    "fm_sheet_get_merge_count",
    "fm_sheet_get_page_margins",
    "fm_sheet_get_page_margins_xml",
    "fm_sheet_get_page_setup",
    "fm_sheet_get_page_setup_xml",
    "fm_sheet_get_print_area",
    "fm_sheet_get_print_options_xml",
    "fm_sheet_get_print_titles",
    "fm_sheet_get_protection",
    "fm_sheet_get_row_override",
    "fm_sheet_get_row_override_count",
    "fm_sheet_get_sheet_pr_xml",
    "fm_sheet_get_validation_at",
    "fm_sheet_get_validation_count",
    "fm_sheet_get_view",
    "fm_sheet_remove_col_break",
    "fm_sheet_remove_hyperlink",
    "fm_sheet_remove_hyperlink_at",
    "fm_sheet_remove_merge",
    "fm_sheet_remove_merge_at",
    "fm_sheet_remove_row_break",
    "fm_sheet_remove_validation_at",
    "fm_sheet_row_break_at",
    "fm_sheet_set_auto_filter_xml",
    "fm_sheet_set_column_hidden",
    "fm_sheet_set_column_outline",
    "fm_sheet_set_column_width",
    "fm_sheet_set_comment",
    "fm_sheet_set_fit_to_page",
    "fm_sheet_set_freeze",
    "fm_sheet_set_header_footer",
    "fm_sheet_set_header_footer_xml",
    "fm_sheet_set_page_margins",
    "fm_sheet_set_page_margins_xml",
    "fm_sheet_set_page_setup",
    "fm_sheet_set_page_setup_xml",
    "fm_sheet_set_print_area",
    "fm_sheet_set_print_options",
    "fm_sheet_set_print_options_xml",
    "fm_sheet_set_print_titles",
    "fm_sheet_set_protection",
    "fm_sheet_set_range_xf_index",
    "fm_sheet_set_right_to_left",
    "fm_sheet_set_row_height",
    "fm_sheet_set_row_hidden",
    "fm_sheet_set_row_outline",
    "fm_sheet_set_sheet_pr_xml",
    "fm_sheet_set_show_grid_lines",
    "fm_sheet_set_show_row_col_headers",
    "fm_sheet_set_show_zeros",
    "fm_sheet_set_tab_hidden",
    "fm_sheet_set_tab_selected",
    "fm_sheet_set_view_mode",
    "fm_sheet_set_visibility",
    "fm_sheet_set_zoom",
    "fm_styles_add_batch",
    "fm_styles_add_border",
    "fm_styles_add_cell_style_xf",
    "fm_styles_add_cell_xf",
    "fm_styles_add_dxf",
    "fm_styles_add_fill",
    "fm_styles_add_font",
    "fm_styles_add_num_fmt",
    "fm_styles_get_border",
    "fm_styles_get_border_count",
    "fm_styles_get_cell_style",
    "fm_styles_get_cell_style_count",
    "fm_styles_get_cell_style_xf",
    "fm_styles_get_cell_style_xf_count",
    "fm_styles_get_cell_xf",
    "fm_styles_get_cell_xf_count",
    "fm_styles_get_dxf",
    "fm_styles_get_dxf_count",
    "fm_styles_get_fill",
    "fm_styles_get_fill_count",
    "fm_styles_get_font",
    "fm_styles_get_font_count",
    "fm_styles_get_num_fmt_string",
    "fm_styles_set_cell_style",
    "fm_styles_set_font",
    "fm_workbook_add_sheet",
    "fm_workbook_calc_mode",
    "fm_workbook_clear_pinned_now",
    "fm_workbook_cell_at",
    "fm_workbook_cell_count",
    "fm_workbook_cf_evaluate_range",
    "fm_workbook_create",
    "fm_workbook_create_empty",
    "fm_workbook_defined_name_at",
    "fm_workbook_delete_cols",
    "fm_workbook_delete_rows",
    "fm_workbook_dependents",
    "fm_workbook_evaluate_cf_formula",
    "fm_workbook_evaluate_formula_array",
    "fm_workbook_evaluate_formula_array_cell",
    "fm_workbook_excel_profile_id",
    "fm_workbook_external_link_at",
    "fm_workbook_external_link_count",
    "fm_workbook_get_cell_phonetic",
    "fm_workbook_get_cell_phonetic_properties",
    "fm_workbook_get_cell_phonetic_run",
    "fm_workbook_get_cell_phonetic_run_count",
    "fm_workbook_get_iterative",
    "fm_workbook_get_value",
    "fm_workbook_insert_cols",
    "fm_workbook_insert_rows",
    "fm_workbook_lambda_text_at",
    "fm_workbook_load",
    "fm_workbook_memory_usage",
    "fm_workbook_move_sheet",
    "fm_workbook_paginate",
    "fm_workbook_partial_recalc",
    "fm_workbook_passthrough_at",
    "fm_workbook_pinned_now",
    "fm_workbook_pivot_cache_count",
    "fm_workbook_pivot_cache_create",
    "fm_workbook_pivot_cache_field_add",
    "fm_workbook_pivot_cache_field_add_shared_item_blank",
    "fm_workbook_pivot_cache_field_add_shared_item_bool",
    "fm_workbook_pivot_cache_field_add_shared_item_error",
    "fm_workbook_pivot_cache_field_add_shared_item_number",
    "fm_workbook_pivot_cache_field_add_shared_item_text",
    "fm_workbook_pivot_cache_field_clear",
    "fm_workbook_pivot_cache_field_clear_shared_items",
    "fm_workbook_pivot_cache_field_count",
    "fm_workbook_pivot_cache_field_name",
    "fm_workbook_pivot_cache_field_shared_item_count",
    "fm_workbook_pivot_cache_get_worksheet_source",
    "fm_workbook_pivot_cache_id_at",
    "fm_workbook_pivot_cache_record_add",
    "fm_workbook_pivot_cache_record_clear",
    "fm_workbook_pivot_cache_record_count",
    "fm_workbook_pivot_cache_record_set_blank",
    "fm_workbook_pivot_cache_record_set_bool",
    "fm_workbook_pivot_cache_record_set_error",
    "fm_workbook_pivot_cache_record_set_number",
    "fm_workbook_pivot_cache_record_set_text",
    "fm_workbook_pivot_cache_remove",
    "fm_workbook_pivot_cache_set_worksheet_source",
    "fm_workbook_pivot_count",
    "fm_workbook_pivot_create",
    "fm_workbook_pivot_data_field_add",
    "fm_workbook_pivot_data_field_clear",
    "fm_workbook_pivot_data_field_count",
    "fm_workbook_pivot_data_field_set",
    "fm_workbook_pivot_field_add",
    "fm_workbook_pivot_field_add_item",
    "fm_workbook_pivot_field_add_item_at",
    "fm_workbook_pivot_field_add_subtotal_fn",
    "fm_workbook_pivot_field_clear",
    "fm_workbook_pivot_field_clear_date_group",
    "fm_workbook_pivot_field_clear_items",
    "fm_workbook_pivot_field_clear_subtotal_fns",
    "fm_workbook_pivot_field_count",
    "fm_workbook_pivot_field_set_axis",
    "fm_workbook_pivot_field_set_date_group",
    "fm_workbook_pivot_field_set_item_visible",
    "fm_workbook_pivot_field_set_number_format",
    "fm_workbook_pivot_field_set_sort",
    "fm_workbook_pivot_field_set_subtotal_top",
    "fm_workbook_pivot_filter_add",
    "fm_workbook_pivot_filter_at",
    "fm_workbook_pivot_filter_clear",
    "fm_workbook_pivot_filter_count",
    "fm_workbook_pivot_filter_remove_at",
    "fm_workbook_pivot_get_layout",
    "fm_workbook_pivot_layout",
    "fm_workbook_pivot_remove",
    "fm_workbook_pivot_set_anchor",
    "fm_workbook_pivot_set_col_field_order",
    "fm_workbook_pivot_set_grand_totals",
    "fm_workbook_pivot_set_layout",
    "fm_workbook_pivot_set_name",
    "fm_workbook_pivot_set_row_field_order",
    "fm_workbook_precedents",
    "fm_workbook_read_diagnostics",
    "fm_workbook_recalc",
    "fm_workbook_remove_sheet",
    "fm_workbook_rename_sheet",
    "fm_workbook_save",
    "fm_workbook_save_as",
    "fm_workbook_save_with_diagnostics",
    "fm_workbook_set_blank",
    "fm_workbook_set_bool",
    "fm_workbook_set_calc_mode",
    "fm_workbook_set_cell_phonetic",
    "fm_workbook_set_cell_phonetic_properties",
    "fm_workbook_set_cell_phonetic_runs",
    "fm_workbook_set_default_font",
    "fm_workbook_set_defined_name",
    "fm_workbook_set_defined_name_scoped",
    "fm_workbook_set_error",
    "fm_workbook_set_excel_profile_id",
    "fm_workbook_set_formula",
    "fm_workbook_set_iterative",
    "fm_workbook_set_iterative_enabled",
    "fm_workbook_set_iterative_progress",
    "fm_workbook_set_number",
    "fm_workbook_set_pinned_now",
    "fm_workbook_set_text",
    "fm_workbook_sheet_name",
    "fm_workbook_spill_info",
    "fm_workbook_table_at",
    "fm_workbook_table_create",
    "fm_workbook_table_remove",
    "fm_workbook_table_update",
)
_STATUS_RETURNING_EXPORTS = frozenset(_STATUS_RETURNING_EXPORT_NAMES)


# ---------------------------------------------------------------------------
# ValueKind enum (matches fm_value_kind_t)
# ---------------------------------------------------------------------------


class ValueKind(IntEnum):
    """Mirror of ``fm_value_kind_t`` in ``formulon_c.h``."""

    BLANK = 0
    NUMBER = 1
    BOOL = 2
    TEXT = 3
    ERROR = 4
    ARRAY = 5
    REF = 6
    LAMBDA = 7


# ---------------------------------------------------------------------------
# WASM module location
# ---------------------------------------------------------------------------


def _locate_wasm() -> Path:
    """Return the on-disk path to ``formulon_capi.wasm``.

    Search order:
      1. ``packages/python/formulon/_wasm/formulon_capi.wasm`` -- the
         package-data location populated by ``make python-package``
         and shipped inside the wheel.
      2. ``$FORMULON_WASM_PATH`` -- explicit override for development.

    Raises:
      FileNotFoundError: when the WASM is not on disk in either
        location. The error message lists every path actually probed --
        always the bundled one, plus the ``FORMULON_WASM_PATH`` value
        when that variable was set.
    """
    here = Path(__file__).resolve().parent
    bundled = here / "_wasm" / "formulon_capi.wasm"
    if bundled.is_file():
        return bundled

    import os

    override = os.environ.get("FORMULON_WASM_PATH")
    if override:
        p = Path(override)
        if p.is_file():
            return p

    tried = [str(bundled)]
    if override:
        tried.append(f"{override} (from FORMULON_WASM_PATH)")
    raise FileNotFoundError(
        "formulon: failed to locate formulon_capi.wasm. "
        f"Tried: {', '.join(tried)}. "
        "Run `make python-package` to stage the artifact, install a "
        "wheel that ships it under formulon/_wasm/, or point "
        "FORMULON_WASM_PATH at an existing build."
    )


# ---------------------------------------------------------------------------
# WASM module / store wrapper
# ---------------------------------------------------------------------------


class _WasmInstance:
    """Owns a single ``wasmtime`` engine, store, and module instance.

    The instance is created lazily on first attribute access; subsequent
    Workbook creations reuse it. This trades cold-start latency for
    repeat-call speed and keeps the engine cache around for the life of
    the Python process.

    The store is `not` thread-safe (per the wasmtime-py docs); a single
    process-wide ``_call_lock`` serialises every WASM invocation. The
    underlying calculation engine is already safe for one outstanding
    recalc per ``Workbook`` handle, so the additional lock only prevents
    cross-handle reentry on the wasmtime store itself.

    A ``wasmtime.Trap`` (an allocation failure, an assertion the engine
    compiles in, ...) aborts the WASM call that raised it partway through,
    which can leave shared state inside the one process-wide instance --
    the allocator's free list, a diagnostic buffer -- half-written. Once
    that happens this wrapper is poisoned: every further export call
    raises :class:`FormulonError` without touching WASM memory again,
    rather than risk operating on a corrupted module.
    """

    def __init__(self) -> None:
        self._engine: Optional[wasmtime.Engine] = None
        self._store: Optional[wasmtime.Store] = None
        self._instance: Optional[wasmtime.Instance] = None
        self._memory: Optional[wasmtime.Memory] = None
        self._exports: dict = {}
        self._init_lock = threading.Lock()
        self._call_lock = threading.RLock()
        self._last_diagnostic = threading.local()
        self._poison_message: Optional[str] = None

    def _ensure(self) -> None:
        if self._instance is not None:
            return
        with self._init_lock:
            if self._instance is not None:
                return
            engine = wasmtime.Engine()
            store = wasmtime.Store(engine)

            # WASI: provide a minimal config. The engine never reads
            # files or stdin; stdout/stderr inherit so any diagnostic
            # libc calls (e.g. trap reasons) surface to the host.
            wasi = wasmtime.WasiConfig()
            wasi.inherit_stdout()
            wasi.inherit_stderr()
            store.set_wasi(wasi)

            wasm_path = _locate_wasm()
            module = wasmtime.Module.from_file(engine, str(wasm_path))

            linker = wasmtime.Linker(engine)
            linker.define_wasi()

            # Stub for env.emscripten_notify_memory_growth. emcc emits
            # this import even under STANDALONE_WASM; it is called on
            # memory.grow but the host has nothing useful to do with it.
            ty = wasmtime.FuncType([wasmtime.ValType.i32()], [])
            linker.define(
                store,
                "env",
                "emscripten_notify_memory_growth",
                wasmtime.Func(store, ty, lambda _i: None),
            )

            instance = linker.instantiate(store, module)
            exports = dict(instance.exports(store).items())

            # Reactor init must run before any export is callable.
            init_fn = exports.get("_initialize")
            if init_fn is not None:
                init_fn(store)

            self._engine = engine
            self._store = store
            self._instance = instance
            self._memory = exports["memory"]
            self._exports = exports

    # ----- raw export accessor --------------------------------------------
    def _call_or_poison(self, fn, name: str, *args):
        """Invoke a raw WASM export, converting a poisoned instance or a
        fresh ``wasmtime.Trap`` into :class:`FormulonError`.

        Shared by ``__getattr__``'s ``_wrapped`` closure and by
        :meth:`alloc` / :meth:`free`, which call ``malloc`` / ``free``
        directly rather than through a named export lookup. The caller
        must already hold ``_call_lock``.
        """
        if self._poison_message is not None:
            raise FormulonError(
                _STATUS_INTERNAL_ERROR,
                op=name,
                _diagnostic_override=(self._poison_message, ""),
                _status_name_override="kInternalError",
            )
        try:
            return fn(self._store, *args)
        except wasmtime.Trap as trap:
            # The call aborted partway through; the module's shared state
            # (allocator, diagnostic buffers) is no longer trustworthy.
            # Poison the instance so every later call, on this or any
            # other Workbook, fails the same way instead of running
            # against it.
            self._poison_message = str(trap)
            raise FormulonError(
                _STATUS_INTERNAL_ERROR,
                op=name,
                _diagnostic_override=(self._poison_message, ""),
                _status_name_override="kInternalError",
            ) from trap

    def __getattr__(self, name: str):
        self._ensure()
        fn = self._exports.get(name)
        if fn is None:
            raise AttributeError(f"WASM export '{name}' not found")
        lock = self._call_lock
        captures_diagnostics = name in _STATUS_RETURNING_EXPORTS

        # Wrap so the caller can use a ctypes-like call syntax.
        def _wrapped(*args):
            with lock:
                if captures_diagnostics:
                    # The C ABI's diagnostic buffers are thread-local in the
                    # engine, but this WASM instance has no pthreads. Keep a
                    # Python-side snapshot per caller and clear it before
                    # every status-returning call so success cannot leave a
                    # previous failure pending.
                    self._last_diagnostic.value = None
                result = self._call_or_poison(fn, name, *args)
                # Only status-returning exports are allowed to populate the
                # snapshot. Counts, indices, and pointers are all i32 in
                # WASM and must never be mistaken for a failure status.
                if captures_diagnostics and result != 0:
                    message_fn = self._exports.get("fm_last_error_message")
                    context_fn = self._exports.get("fm_last_error_context")
                    if message_fn is not None and context_fn is not None:
                        self._last_diagnostic.value = (
                            result,
                            self._read_cstr_unlocked(message_fn(self._store)),
                            self._read_cstr_unlocked(context_fn(self._store)),
                        )
                return result

        # Callers that derive a FormulonError's ``op`` prefix from
        # ``fn.__name__`` must see the C ABI entry point, not this
        # binding-internal closure.
        _wrapped.__name__ = name
        _wrapped.__qualname__ = name
        return _wrapped

    # ----- memory primitives ----------------------------------------------
    @property
    def store(self) -> wasmtime.Store:
        self._ensure()
        assert self._store is not None
        return self._store

    @property
    def memory(self) -> wasmtime.Memory:
        self._ensure()
        assert self._memory is not None
        return self._memory

    @property
    def call_lock(self) -> threading.RLock:
        return self._call_lock

    def read_bytes(self, ptr: int, length: int) -> bytes:
        """Copy ``length`` bytes from WASM memory starting at ``ptr``.

        Takes ``_call_lock`` for the duration of the read: the
        wasmtime ``Store`` is not thread-safe, so a concurrent WASM
        call on another thread (which can grow linear memory) must not
        race with this read.
        """
        self._ensure()
        assert self._memory is not None
        with self._call_lock:
            return bytes(self._memory.read(self._store, ptr, ptr + length))

    def write_bytes(self, ptr: int, data: bytes) -> None:
        """Write ``data`` into WASM memory starting at ``ptr``.

        Takes ``_call_lock`` for the same reason as :meth:`read_bytes`.
        """
        self._ensure()
        assert self._memory is not None
        with self._call_lock:
            self._memory.write(self._store, data, ptr)

    def read_u32(self, ptr: int) -> int:
        return struct.unpack("<I", self.read_bytes(ptr, 4))[0]

    def read_i32(self, ptr: int) -> int:
        return struct.unpack("<i", self.read_bytes(ptr, 4))[0]

    def read_f64(self, ptr: int) -> float:
        return struct.unpack("<d", self.read_bytes(ptr, 8))[0]

    def read_cstr(self, ptr: int) -> str:
        """Decode a NUL-terminated UTF-8 C string from ``ptr``.

        Returns the empty string when ``ptr`` is 0. The whole scan runs
        under ``_call_lock`` (see :meth:`read_bytes`) so a concurrent
        WASM call on another thread cannot grow memory mid-read.
        """
        if ptr == 0:
            return ""
        self._ensure()
        assert self._memory is not None
        # Stream a chunk at a time to avoid copying the whole memory.
        chunks: list[bytes] = []
        offset = ptr
        chunk_size = 256
        with self._call_lock:
            mem_len = self._memory.data_len(self._store)
            while offset < mem_len:
                end = min(offset + chunk_size, mem_len)
                buf = bytes(self._memory.read(self._store, offset, end))
                nul = buf.find(b"\x00")
                if nul >= 0:
                    chunks.append(buf[:nul])
                    break
                chunks.append(buf)
                offset = end
        return b"".join(chunks).decode("utf-8", errors="replace")

    def _read_cstr_unlocked(self, ptr: int) -> str:
        """Decode a C string while the caller already owns ``_call_lock``."""
        if ptr == 0:
            return ""
        assert self._memory is not None
        chunks: list[bytes] = []
        offset = ptr
        mem_len = self._memory.data_len(self._store)
        while offset < mem_len:
            end = min(offset + 256, mem_len)
            buf = bytes(self._memory.read(self._store, offset, end))
            nul = buf.find(b"\x00")
            if nul >= 0:
                chunks.append(buf[:nul])
                break
            chunks.append(buf)
            offset = end
        return b"".join(chunks).decode("utf-8", errors="replace")

    def last_diagnostic(self, status: int) -> Tuple[str, str]:
        """Take the diagnostic captured for ``status`` once, if any.

        A missing or mismatched snapshot is deliberately not filled by
        reading the shared WASM diagnostic exports: doing so would race a
        later call and would make non-status APIs look like failures.
        """
        snapshot = getattr(self._last_diagnostic, "value", None)
        self._last_diagnostic.value = None
        if snapshot is not None and snapshot[0] == status:
            return snapshot[1], snapshot[2]
        return "", ""

    def alloc(self, size: int) -> int:
        """Allocate ``size`` bytes in WASM memory; return the pointer.

        Raises:
          MemoryError: when the WASM-side allocator returns NULL.
        """
        if size <= 0:
            return 0
        self._ensure()
        with self._call_lock:
            ptr = self._call_or_poison(self._exports["malloc"], "malloc", size)
        if ptr == 0:
            raise MemoryError(f"formulon: WASM malloc({size}) returned NULL")
        return ptr

    def free(self, ptr: int) -> None:
        if ptr == 0:
            return
        self._ensure()
        with self._call_lock:
            self._call_or_poison(self._exports["free"], "free", ptr)

    def alloc_utf8(self, s: str) -> Tuple[int, int]:
        """Encode ``s`` as UTF-8 and copy it into WASM memory.

        Returns ``(ptr, length_with_nul)``. The caller MUST free ``ptr``
        with :meth:`free` once the call that consumed it returns.
        """
        if not isinstance(s, str):
            raise TypeError(f"expected str, got {type(s).__name__}")
        buf = s.encode("utf-8") + b"\x00"
        ptr = self.alloc(len(buf))
        self.write_bytes(ptr, buf)
        return ptr, len(buf)

    def alloc_bytes(self, data: bytes) -> int:
        """Copy ``data`` into WASM memory; return the pointer."""
        if len(data) == 0:
            return 0
        ptr = self.alloc(len(data))
        self.write_bytes(ptr, data)
        return ptr


# ---------------------------------------------------------------------------
# Module-level singleton + thin compatibility shim
# ---------------------------------------------------------------------------


LIB = _WasmInstance()


def decode_cstr(ptr_or_bytes) -> str:
    """Backwards-compat shim for :class:`formulon.workbook.FormulonError`.

    Accepts either an int WASM pointer (decoded via :class:`LIB`) or a
    bytes object (legacy ctypes path). Returns the empty string when
    the input is ``None``, ``0``, or empty.
    """
    if ptr_or_bytes is None:
        return ""
    if isinstance(ptr_or_bytes, bytes):
        return ptr_or_bytes.decode("utf-8", errors="replace")
    if isinstance(ptr_or_bytes, int):
        return LIB.read_cstr(ptr_or_bytes)
    raise TypeError(f"decode_cstr: unexpected type {type(ptr_or_bytes).__name__}")
