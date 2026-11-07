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
``fm_sheet_get_value``) are read directly from WASM memory and
decoded eagerly into Python ``str`` so the result outlives any
subsequent WASM mutation.

Excel formula profiles
----------------------

New workbooks use ``win-365-en_US`` by default. The profile getter returns
one of the 28 ids
``{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}``;
the setter accepts the same ids.

The ``fm_value_t`` POD has the wasm32 layout::

    offset 0: int32  kind
    offset 4: int32  (padding for 8-byte alignment of the union)
    offset 8: union  { double number ; int32 boolean ; int32 error_code ; ptr text }
    total: 16 bytes

The union member is selected by ``kind``; reading any other member is
undefined per the C ABI contract.
"""

from __future__ import annotations

import math
import numbers
import operator
import struct
import threading
from dataclasses import dataclass
from enum import IntEnum
from pathlib import Path
from typing import Optional, Tuple, Union

import wasmtime

from ._status import INDEX_OUT_OF_RANGE_CODES, NOT_FOUND_CODES, StatusCode
from ._structs import PARSE_ERROR

__all__ = [
    "LIB",
    "FormulonError",
    "IndexOutOfRangeError",
    "FormulaSyntaxError",
    "InvalidArgumentError",
    "InvalidFileError",
    "NotFoundError",
    "ParseError",
    "ParseErrorCode",
    "WorkbookClosedError",
    "ValueKind",
    "decode_cstr",
    "fm_value_t_size",
]

# fm_value_t layout: int32 kind + 4 pad + 8 union = 16 bytes.
fm_value_t_size = 16


def _coerce_integral(value: int, name: str) -> int:
    """Convert an integer or integral real without truncating a fraction.

    A non-number raises ``TypeError``; a fractional or non-finite real raises
    ``ValueError``.
    """

    try:
        return operator.index(value)
    except TypeError:
        pass
    if not isinstance(value, numbers.Real):
        raise TypeError(f"formulon: {name} must be an integer, got {type(value).__name__}")
    if not math.isfinite(value) or value != int(value):
        raise ValueError(f"formulon: {name} must be an integer, got {value!r}")
    return int(value)


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
      TypeError: when ``value`` is not a number.
      ValueError: when ``value`` is fractional, negative, or does not fit
        ``bits``.
    """
    ivalue = _coerce_integral(value, name)
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
      TypeError: when ``value`` is not a number.
      ValueError: when ``value`` is fractional or does not fit a signed
        ``bits``-wide int.
    """
    ivalue = _coerce_integral(value, name)
    limit = 1 << (bits - 1)
    if ivalue < -limit or ivalue >= limit:
        raise ValueError(f"formulon: {name} out of range for int{bits}: {ivalue}")
    return ivalue


class ParseErrorCode(IntEnum):
    """Mirror of ``fm_parse_error_code_t`` in ``formulon_c.h``."""

    LEXER_INVALID_CHARACTER = 0
    LEXER_UNTERMINATED_STRING = 1
    LEXER_UNTERMINATED_SHEET_QUOTE = 2
    LEXER_INVALID_NUMBER_LITERAL = 3
    LEXER_INVALID_ERROR_LITERAL = 4
    LEXER_INVALID_ESCAPE = 5
    LEXER_INVALID_REFERENCE = 6
    LEXER_EXCESSIVE_LENGTH = 7
    UNEXPECTED_TOKEN = 8
    UNEXPECTED_EOF = 9
    EXPECTED_EXPRESSION = 10
    EXPECTED_CLOSE_PAREN = 11
    UNBALANCED_BRACES = 12
    ARRAY_ROW_MISMATCH = 13
    EXPECTED_RPAREN_OR_COMMA = 14
    EXPECTED_COMMA_OR_SEMI_IN_ARRAY = 15
    INVALID_REFERENCE = 16
    UNSUPPORTED_CONSTRUCT = 17
    EXPECTED_OPEN_PAREN = 18
    EXPECTED_COMMA = 19
    UNBALANCED_BRACKETS = 20
    INVALID_RANGE = 21
    NESTED_FORMULA_TOO_DEEP = 22
    TOO_MANY_ERRORS = 23
    LET_INVALID_NAME = 24
    LET_WRONG_ARITY = 25
    LAMBDA_INVALID_PARAM = 26
    LAMBDA_EMPTY = 27
    LAMBDA_DUPLICATE_PARAM = 28
    ENTRY_WRONG_ARITY = 29
    ENTRY_REFERENCE_REQUIRED = 30
    ENTRY_QUALIFIED_BUILTIN = 31
    ENTRY_REFUSED_WHITESPACE = 32
    ENTRY_PARAM_SHAPED_NAME = 33
    ENTRY_DOUBLE_EQUALS = 34
    ENTRY_BOM_NOT_SUPPORTED = 35
    STRUCTURED_REF_INVALID = 36


@dataclass(frozen=True)
class ParseError:
    """Position of a formula syntax failure (``fm_parse_error_t``).

    ``start`` / ``end`` are UTF-16 code units into the text the caller passed;
    ``line`` and ``column`` are 1-based.
    """

    code: ParseErrorCode
    start: int
    end: int
    line: int
    column: int


# Statuses 1000..1014 are formula syntax failures; each records a ParseError.
_SYNTAX_STATUS_FIRST = 1000
_SYNTAX_STATUS_LAST = 1014

# The I/O statuses that mean the package cannot be opened as a workbook.
_INVALID_FILE_CODES = frozenset(
    {
        StatusCode.IO_ZIP_CORRUPT,
        StatusCode.IO_ZIP_BOMB,
        StatusCode.IO_ZIP_SLIP,
        StatusCode.IO_XML_PARSE,
        StatusCode.IO_XML_DOCTYPE,
        StatusCode.IO_XML_ENTITY_EXPLOSION,
        StatusCode.IO_RELATIONSHIP_BROKEN,
        StatusCode.IO_CONTENT_TYPE_INVALID,
        StatusCode.IO_SHEET_CORRUPT,
        StatusCode.IO_XLSB_RECORD_CORRUPT,
        StatusCode.IO_XLSB_RECORD_TRUNCATED,
        StatusCode.IO_XLSB_UNKNOWN_RECORD,
        StatusCode.IO_XLSB_UNSUPPORTED_PTG,
        StatusCode.IO_XLSB_CORRUPT,
        StatusCode.IO_XLSB_RETAINED_PART_STALE,
    }
)


class FormulonError(Exception):
    """Raised for host-side failures of the Formulon engine.

    Constructing ``FormulonError`` with a status that has a dedicated
    subclass yields that subclass (:class:`IndexOutOfRangeError`,
    :class:`InvalidArgumentError`, :class:`FormulaSyntaxError`,
    :class:`InvalidFileError`, :class:`NotFoundError`,
    :class:`WorkbookClosedError`), so ``except IndexError`` / ``except
    ValueError`` / ``except KeyError`` also catch the matching failures.

    ``str(e)`` is the engine's message, followed by ``[context]`` when the
    context is non-empty.

    Attributes:
      code: the :class:`~formulon.StatusCode` member of the failure (a plain
        ``int`` for a value the enum does not list).
      message: human-readable description of the failure.
      context: structured ``key=value`` details; empty when there are none.

    The message and context are copied out when the failure is detected, so
    they stay valid whatever later calls do.

    Defined here rather than in :mod:`formulon.workbook` (which
    re-exports it) so the low-level export trampoline below can raise it
    directly on a trap, without importing back into the module that
    imports ``LIB`` from here.
    """

    code: Union[StatusCode, int]
    message: str
    context: str
    parse_error: Optional[ParseError]

    def __new__(cls, status: int, *args: object, **kwargs: object) -> "FormulonError":
        if cls is FormulonError:
            if _SYNTAX_STATUS_FIRST <= status <= _SYNTAX_STATUS_LAST:
                cls = FormulaSyntaxError
            elif status in INDEX_OUT_OF_RANGE_CODES:
                cls = IndexOutOfRangeError
            elif status == StatusCode.INVALID_ARGUMENT:
                cls = InvalidArgumentError
            elif status in NOT_FOUND_CODES:
                cls = NotFoundError
            elif status == StatusCode.BINDING_INVALID_HANDLE:
                cls = WorkbookClosedError
            elif status in _INVALID_FILE_CODES:
                cls = InvalidFileError
        return super().__new__(cls)

    def __init__(
        self,
        status: int,
        *,
        _diagnostic_override: Optional[Tuple[str, str]] = None,
        _host_only: bool = False,
    ) -> None:
        try:
            self.code = StatusCode(int(status))
        except ValueError:
            self.code = int(status)
        self.parse_error = None
        if _host_only:
            # A trap leaves the module instance poisoned (see `_wrapped`
            # below); reading the diagnostic buffers would call back into
            # it, so this path never touches WASM memory.
            self.message, self.context = _diagnostic_override or ("", "")
        else:
            # An override still drains a pending C diagnostic so it cannot
            # leak into a later exception on this Python thread.
            message, context = LIB.last_diagnostic(int(status))
            if _diagnostic_override is not None:
                message, context = _diagnostic_override
            self.message, self.context = message, context
            self.parse_error = LIB.last_parse_error(int(status))
        text = self.message
        if self.context:
            text += f" [{self.context}]"
        super().__init__(text)

    def __str__(self) -> str:
        # `KeyError` (a base of `NotFoundError`) would otherwise render the
        # message through `repr`.
        return Exception.__str__(self)


class InvalidArgumentError(FormulonError, ValueError):
    """The engine refused an argument (``StatusCode.INVALID_ARGUMENT``)."""


class FormulaSyntaxError(InvalidArgumentError):
    """A formula or reference text failed to parse (statuses 1000 to 1014).

    ``parse_error_code``, ``start``, ``end``, ``line`` and ``column`` describe
    the offending span (see :class:`ParseError`) and are ``None`` when the
    engine recorded no position.
    """

    @property
    def parse_error_code(self) -> Optional[ParseErrorCode]:
        return None if self.parse_error is None else self.parse_error.code

    @property
    def start(self) -> Optional[int]:
        return None if self.parse_error is None else self.parse_error.start

    @property
    def end(self) -> Optional[int]:
        return None if self.parse_error is None else self.parse_error.end

    @property
    def line(self) -> Optional[int]:
        return None if self.parse_error is None else self.parse_error.line

    @property
    def column(self) -> Optional[int]:
        return None if self.parse_error is None else self.parse_error.column


class NotFoundError(FormulonError, KeyError):
    """A named item does not exist (``NOT_FOUND`` and its variants)."""


class IndexOutOfRangeError(FormulonError, IndexError):
    """An index is past the end of its collection (``SHEET_INDEX_OUT_OF_RANGE``, ``INDEX_OUT_OF_RANGE``)."""


class WorkbookClosedError(FormulonError, ValueError):
    """The workbook was used after :meth:`~formulon.Workbook.close` (``BINDING_INVALID_HANDLE``)."""


class InvalidFileError(FormulonError):
    """The bytes cannot be opened as a workbook package.

    Raised for a corrupt or hostile archive, unparsable XML, broken
    relationships or content types, and corrupt XLSB records. A damaged part
    the workbook still opens without (a drawing, a chart) is not one, nor is
    an unsupported variant (``StatusCode.IO_UNSUPPORTED_VARIANT``).
    """


# Generated from the current working-tree C header and WASM export manifest:
# {src/c_api/formulon_c.h declarations returning fm_status_t} intersection
# {tools/wasm/capi_exports.txt}. Keep this literal so the drift checker can
# validate the binding without importing this module (or wasmtime).
_STATUS_RETURNING_EXPORT_NAMES = (
    "fm_get_last_parse_error",
    "fm_sheet_get_tab_color",
    "fm_sheet_set_tab_color",
    "fm_sheet_clear_cell_xf_index",
    "fm_sheet_get_cell_xf_index",
    "fm_cell_nodes_get_node_at",
    "fm_cell_range_get_cell_at",
    "fm_cell_range_get_cell_count",
    "fm_cell_range_next_cursor",
    "fm_sheet_set_cell_xf_index",
    "fm_cf_results_get_cell_at",
    "fm_cf_results_get_match_at",
    "fm_chart_get_axis_at",
    "fm_chart_data_get_axis_title_at",
    "fm_chart_data_get_category_at",
    "fm_chart_data_get_category_level_at",
    "fm_chart_data_get_category_level_size",
    "fm_chart_data_get_info",
    "fm_chart_data_get_range_at",
    "fm_chart_data_get_series_at",
    "fm_chart_data_get_series_points",
    "fm_chart_data_get_unresolved_at",
    "fm_chart_get_hidden_category_at",
    "fm_chart_get_info",
    "fm_chart_get_legend_hidden_at",
    "fm_chart_get_plot_at",
    "fm_chart_get_plot_axis_id_at",
    "fm_chart_get_series_at",
    "fm_chart_get_series_literal_at",
    "fm_chart_get_series_point_at",
    "fm_get_collation_classes",
    "fm_canonicalize_formula",
    "fm_localize_formula",
    "fm_canonicalize_function_name",
    "fm_localize_function_name",
    "fm_get_function_metadata",
    "fm_get_function_name_at",
    "fm_get_locale_error_name_at",
    "fm_get_locale_facts",
    "fm_pagination_get_row_break_at",
    "fm_pagination_is_row_break_manual",
    "fm_pagination_get_margins",
    "fm_pagination_get_page_at",
    "fm_pagination_get_page_order",
    "fm_pagination_get_paper",
    "fm_pagination_get_print_area_at",
    "fm_pagination_get_print_titles",
    "fm_pagination_get_printable",
    "fm_pagination_get_scale",
    "fm_pagination_get_column_break_at",
    "fm_pagination_is_column_break_manual",
    "fm_pivot_cells_get_cell_at",
    "fm_pivot_cells_get_bounds",
    "fm_set_log_min_level",
    "fm_set_log_sink",
    "fm_sheet_add_chart_series",
    "fm_sheet_add_column_break",
    "fm_sheet_add_data_table",
    "fm_sheet_add_hyperlink",
    "fm_sheet_add_merge",
    "fm_sheet_add_protected_range",
    "fm_sheet_add_row_break",
    "fm_sheet_add_slicer",
    "fm_sheet_add_sparkline_group",
    "fm_sheet_add_threaded_comment",
    "fm_sheet_add_validation",
    "fm_sheet_apply_auto_filter",
    "fm_sheet_get_cell_rect_pt",
    "fm_sheet_get_cells_in_range",
    "fm_sheet_add_cf_rule",
    "fm_sheet_clear_cf_rules",
    "fm_sheet_get_cf_rule_count",
    "fm_sheet_get_cf_rule",
    "fm_sheet_remove_cf_rule",
    "fm_sheet_remove_comment",
    "fm_sheet_get_chart_data",
    "fm_sheet_get_chart",
    "fm_sheet_get_chart_preview",
    "fm_sheet_clear_auto_filter",
    "fm_sheet_clear_breaks",
    "fm_sheet_clear_column_style",
    "fm_sheet_clear_column_width",
    "fm_sheet_clear_hyperlinks",
    "fm_sheet_clear_merges",
    "fm_sheet_clear_row_height",
    "fm_sheet_clear_row_style",
    "fm_sheet_clear_validations",
    "fm_sheet_get_column_break_at",
    "fm_sheet_convert_column_chars_to_pt",
    "fm_sheet_convert_column_pt_to_chars",
    "fm_sheet_get_column_width_pt",
    "fm_sheet_get_data_table",
    "fm_sheet_get_drawing_object_at",
    "fm_sheet_get_drawing_object",
    "fm_sheet_get_drawing_object_count",
    "fm_sheet_update_threaded_comment",
    "fm_sheet_evaluate_auto_filter",
    "fm_sheet_get_auto_filter",
    "fm_sheet_get_auto_filter_xml",
    "fm_sheet_get_column_override_at",
    "fm_sheet_get_column_override_count",
    "fm_sheet_get_comment",
    "fm_sheet_get_comment_at",
    "fm_sheet_get_comment_count",
    "fm_sheet_get_comment_visible",
    "fm_sheet_get_effective_style",
    "fm_sheet_get_format_defaults",
    "fm_sheet_get_header_footer_xml",
    "fm_sheet_get_hyperlink_at",
    "fm_sheet_get_hyperlink",
    "fm_sheet_get_hyperlink_count",
    "fm_sheet_get_image",
    "fm_sheet_get_merge_at",
    "fm_sheet_get_merge_count",
    "fm_sheet_get_page_margins",
    "fm_sheet_get_page_margins_xml",
    "fm_sheet_get_page_setup",
    "fm_sheet_get_print_options",
    "fm_sheet_get_header_footer",
    "fm_sheet_get_page_setup_xml",
    "fm_sheet_get_print_area",
    "fm_sheet_get_print_options_xml",
    "fm_sheet_get_print_titles",
    "fm_sheet_get_protected_range",
    "fm_sheet_get_sheet_protection",
    "fm_sheet_get_row_override_at",
    "fm_sheet_get_row_override_count",
    "fm_sheet_get_selection",
    "fm_sheet_get_sheet_pr_xml",
    "fm_sheet_get_validation",
    "fm_sheet_get_validation_count",
    "fm_sheet_get_view",
    "fm_sheet_add_chart",
    "fm_sheet_add_image",
    "fm_sheet_add_shape",
    "fm_sheet_get_invalid_cells",
    "fm_sheet_get_merges_in_range",
    "fm_sheet_get_protected_range_count",
    "fm_sheet_get_range_style_bases",
    "fm_sheet_get_range_style_bases_excluding",
    "fm_sheet_remap_range_style",
    "fm_sheet_remove_auto_filter",
    "fm_sheet_remove_chart_series",
    "fm_sheet_remove_column_break",
    "fm_sheet_remove_data_table",
    "fm_sheet_remove_drawing",
    "fm_sheet_remove_hyperlink",
    "fm_sheet_remove_hyperlink_at",
    "fm_sheet_remove_merge",
    "fm_sheet_remove_merge_at",
    "fm_sheet_remove_protected_range",
    "fm_sheet_remove_row_break",
    "fm_sheet_remove_shape",
    "fm_sheet_remove_slicer",
    "fm_sheet_remove_sparkline_group",
    "fm_sheet_remove_threaded_comment",
    "fm_sheet_remove_validation",
    "fm_sheet_restore_drawing",
    "fm_sheet_restore_range_style",
    "fm_sheet_get_row_break_at",
    "fm_sheet_get_row_height_pt",
    "fm_sheet_set_auto_filter",
    "fm_sheet_set_auto_filter_xml",
    "fm_sheet_set_chart_axis",
    "fm_sheet_set_chart_display",
    "fm_sheet_set_chart_filter",
    "fm_sheet_set_chart_legend",
    "fm_sheet_set_chart_plot",
    "fm_sheet_set_chart_point",
    "fm_sheet_set_chart_series",
    "fm_sheet_set_chart_source",
    "fm_sheet_set_chart_title",
    "fm_sheet_set_chart_type",
    "fm_sheet_set_column_collapsed",
    "fm_sheet_set_column_hidden",
    "fm_sheet_set_column_outline",
    "fm_sheet_set_column_style",
    "fm_sheet_set_column_width",
    "fm_sheet_set_comment",
    "fm_sheet_set_comment_visible",
    "fm_sheet_set_drawing_anchor",
    "fm_sheet_set_drawing_name",
    "fm_sheet_set_drawing_z_order",
    "fm_sheet_set_format_defaults",
    "fm_sheet_set_freeze",
    "fm_sheet_set_header_footer",
    "fm_sheet_set_header_footer_xml",
    "fm_sheet_set_page_margins",
    "fm_sheet_set_page_margins_xml",
    "fm_sheet_set_page_setup",
    "fm_sheet_set_page_setup_xml",
    "fm_sheet_set_pane_top_left",
    "fm_sheet_set_print_area",
    "fm_sheet_set_print_options",
    "fm_sheet_set_print_options_xml",
    "fm_sheet_set_print_titles",
    "fm_sheet_set_sheet_protection",
    "fm_sheet_set_range_xf_index",
    "fm_sheet_set_right_to_left",
    "fm_sheet_set_row_auto_height",
    "fm_sheet_set_row_collapsed",
    "fm_sheet_set_row_height",
    "fm_sheet_set_row_hidden",
    "fm_sheet_set_row_outline",
    "fm_sheet_set_row_style",
    "fm_sheet_set_selection",
    "fm_sheet_set_sheet_pr_xml",
    "fm_sheet_set_show_formulas",
    "fm_sheet_set_show_grid_lines",
    "fm_sheet_set_show_outline_symbols",
    "fm_sheet_set_show_row_col_headers",
    "fm_sheet_set_show_zeros",
    "fm_sheet_set_split",
    "fm_sheet_set_tab_selected",
    "fm_sheet_set_top_left",
    "fm_sheet_set_view_mode",
    "fm_sheet_set_visibility",
    "fm_sheet_set_zoom",
    "fm_sheet_get_shape_at",
    "fm_sheet_get_shape",
    "fm_sheet_get_shape_count",
    "fm_sheet_get_slicer_at",
    "fm_sheet_get_slicer",
    "fm_sheet_get_slicer_count",
    "fm_sheet_snapshot_drawing",
    "fm_sheet_snapshot_range_style",
    "fm_sheet_get_sparkline",
    "fm_sheet_get_sparkline_group",
    "fm_sheet_get_sparkline_group_count",
    "fm_sheet_get_threaded_comment_at",
    "fm_sheet_get_threaded_comment",
    "fm_sheet_get_threaded_comment_count",
    "fm_sheet_update_protected_range",
    "fm_sheet_update_shape",
    "fm_sheet_update_slicer",
    "fm_sheet_update_sparkline_group",
    "fm_sheet_validate_value",
    "fm_sheet_get_width_model",
    "fm_workbook_add_batch",
    "fm_workbook_add_border",
    "fm_workbook_add_cell_style_xf",
    "fm_workbook_add_cell_xf",
    "fm_workbook_add_dxf",
    "fm_workbook_add_fill",
    "fm_workbook_add_font",
    "fm_workbook_add_num_fmt",
    "fm_workbook_get_border",
    "fm_workbook_get_border_count",
    "fm_workbook_get_cell_style_at",
    "fm_workbook_get_cell_style",
    "fm_workbook_get_cell_style_count",
    "fm_workbook_get_cell_style_xf",
    "fm_workbook_get_cell_style_xf_count",
    "fm_workbook_get_cell_xf",
    "fm_workbook_get_cell_xf_count",
    "fm_workbook_get_dxf",
    "fm_workbook_get_dxf_count",
    "fm_workbook_get_fill",
    "fm_workbook_get_fill_count",
    "fm_workbook_get_font",
    "fm_workbook_get_font_count",
    "fm_workbook_get_num_fmt_string",
    "fm_workbook_get_table_style_at",
    "fm_workbook_get_table_style",
    "fm_workbook_get_table_style_count",
    "fm_workbook_get_table_style_defaults",
    "fm_workbook_remove_cell_style",
    "fm_workbook_remove_table_style",
    "fm_workbook_set_cell_style",
    "fm_workbook_set_font",
    "fm_workbook_set_table_style",
    "fm_workbook_set_table_style_defaults",
    "fm_workbook_apply_table_auto_filter",
    "fm_workbook_clear_table_auto_filter",
    "fm_workbook_evaluate_table_auto_filter",
    "fm_workbook_get_table_auto_filter",
    "fm_workbook_remove_table_auto_filter",
    "fm_workbook_set_table_auto_filter",
    "fm_compare_text",
    "fm_workbook_add_person",
    "fm_workbook_add_sheet",
    "fm_workbook_get_calc_mode",
    "fm_workbook_get_calc_on_save",
    "fm_sheet_get_cell_at",
    "fm_sheet_get_cell_count",
    "fm_sheet_evaluate_cf_range",
    "fm_workbook_clear_pinned_now",
    "fm_workbook_copy_sheet",
    "fm_workbook_create_default",
    "fm_workbook_create_empty",
    "fm_workbook_create_for_profile",
    "fm_workbook_get_date1904",
    "fm_workbook_get_defined_name_at",
    "fm_workbook_get_defined_name",
    "fm_workbook_remove_defined_name",
    "fm_sheet_delete_columns",
    "fm_sheet_delete_rows",
    "fm_sheet_get_dependents",
    "fm_sheet_evaluate_cf_formula",
    "fm_sheet_evaluate_formula_array",
    "fm_value_array_get_dimensions",
    "fm_value_array_get_value_count",
    "fm_value_array_get_value_at",
    "fm_workbook_get_profile_id",
    "fm_workbook_get_external_link",
    "fm_workbook_get_external_link_count",
    "fm_workbook_get_file_name",
    "fm_workbook_format_reference",
    "fm_workbook_format_value",
    "fm_workbook_get_full_precision",
    "fm_workbook_get_active_sheet",
    "fm_sheet_get_array_formula",
    "fm_sheet_get_cell_phonetic",
    "fm_sheet_get_cell_phonetic_properties",
    "fm_sheet_get_cell_phonetic_run",
    "fm_sheet_get_cell_phonetic_run_count",
    "fm_sheet_get_cell_show_phonetic",
    "fm_sheet_get_display_text",
    "fm_sheet_get_formula",
    "fm_sheet_get_formula_r1c1",
    "fm_workbook_get_iterative",
    "fm_workbook_get_workbook_protection",
    "fm_workbook_get_theme_colors",
    "fm_workbook_get_theme_fonts",
    "fm_sheet_get_value",
    "fm_sheet_has_cell_phonetic_properties",
    "fm_sheet_insert_columns",
    "fm_sheet_insert_rows",
    "fm_sheet_get_lambda_text",
    "fm_workbook_load",
    "fm_workbook_get_memory_usage",
    "fm_workbook_move_sheet",
    "fm_sheet_get_pagination",
    "fm_workbook_parse_reference",
    "fm_sheet_recalc_viewport",
    "fm_workbook_get_passthrough_at",
    "fm_workbook_get_passthrough",
    "fm_workbook_get_person_at",
    "fm_workbook_get_person",
    "fm_workbook_get_person_count",
    "fm_workbook_get_pinned_now",
    "fm_workbook_add_pivot_cache_calculated_field",
    "fm_workbook_remove_pivot_cache_calculated_field",
    "fm_workbook_set_pivot_cache_calculated_field_formula",
    "fm_workbook_add_pivot_cache_calculated_item",
    "fm_workbook_get_pivot_cache_calculated_item",
    "fm_workbook_get_pivot_cache_calculated_item_count",
    "fm_workbook_move_pivot_cache_calculated_item",
    "fm_workbook_remove_pivot_cache_calculated_item",
    "fm_workbook_set_pivot_cache_calculated_item_formula",
    "fm_workbook_get_pivot_cache_count",
    "fm_workbook_add_pivot_cache",
    "fm_workbook_add_pivot_cache_field",
    "fm_workbook_add_pivot_cache_field_shared_item_blank",
    "fm_workbook_add_pivot_cache_field_shared_item_bool",
    "fm_workbook_add_pivot_cache_field_shared_item_date",
    "fm_workbook_add_pivot_cache_field_shared_item_error",
    "fm_workbook_add_pivot_cache_field_shared_item_number",
    "fm_workbook_add_pivot_cache_field_shared_item_text",
    "fm_workbook_clear_pivot_cache_fields",
    "fm_workbook_clear_pivot_cache_field_shared_items",
    "fm_workbook_get_pivot_cache_field_count",
    "fm_workbook_get_pivot_cache_field_formula",
    "fm_workbook_get_pivot_cache_field_name",
    "fm_workbook_get_pivot_cache_field_shared_item_count",
    "fm_workbook_is_pivot_cache_field_shared_item_date",
    "fm_workbook_get_pivot_cache_field_shared_item_value",
    "fm_workbook_get_pivot_cache_options",
    "fm_workbook_get_pivot_cache_worksheet_source",
    "fm_workbook_get_pivot_cache_id_at",
    "fm_workbook_add_pivot_cache_record",
    "fm_workbook_clear_pivot_cache_records",
    "fm_workbook_get_pivot_cache_record_count",
    "fm_workbook_set_pivot_cache_record_blank",
    "fm_workbook_set_pivot_cache_record_bool",
    "fm_workbook_set_pivot_cache_record_date",
    "fm_workbook_set_pivot_cache_record_error",
    "fm_workbook_set_pivot_cache_record_number",
    "fm_workbook_set_pivot_cache_record_text",
    "fm_workbook_remove_pivot_cache",
    "fm_workbook_set_pivot_cache_options",
    "fm_workbook_set_pivot_cache_worksheet_source",
    "fm_sheet_get_pivot_count",
    "fm_sheet_add_pivot",
    "fm_sheet_add_pivot_data_field",
    "fm_sheet_get_pivot_data_field",
    "fm_sheet_clear_pivot_data_fields",
    "fm_sheet_get_pivot_data_field_count",
    "fm_sheet_get_pivot_data_field_default_name",
    "fm_sheet_set_pivot_data_field",
    "fm_sheet_add_pivot_field",
    "fm_sheet_add_pivot_field_item",
    "fm_sheet_add_pivot_field_shared_item",
    "fm_sheet_add_pivot_field_subtotal_fn",
    "fm_sheet_get_pivot_field",
    "fm_sheet_clear_pivot_fields",
    "fm_sheet_clear_pivot_field_date_group",
    "fm_sheet_clear_pivot_field_items",
    "fm_sheet_clear_pivot_field_subtotal_fns",
    "fm_sheet_get_pivot_field_count",
    "fm_sheet_get_pivot_field_item",
    "fm_sheet_get_pivot_field_item_count",
    "fm_sheet_set_pivot_field_axis",
    "fm_sheet_set_pivot_field_custom_name",
    "fm_sheet_set_pivot_field_date_group",
    "fm_sheet_set_pivot_field_default_subtotal",
    "fm_sheet_set_pivot_field_fill_down_labels",
    "fm_sheet_set_pivot_field_insert_blank_row",
    "fm_sheet_set_pivot_field_item_caption",
    "fm_sheet_set_pivot_field_item_show_detail",
    "fm_sheet_set_pivot_field_item_visible",
    "fm_sheet_set_pivot_field_number_format",
    "fm_sheet_set_pivot_field_show_all",
    "fm_sheet_set_pivot_field_show_detail",
    "fm_sheet_set_pivot_field_sort",
    "fm_sheet_set_pivot_field_subtotal_top",
    "fm_sheet_add_pivot_filter",
    "fm_sheet_get_pivot_filter",
    "fm_sheet_clear_pivot_filters",
    "fm_sheet_get_pivot_filter_count",
    "fm_sheet_remove_pivot_filter",
    "fm_sheet_get_pivot_cache_id",
    "fm_sheet_get_pivot_column_field_order",
    "fm_sheet_get_pivot_grand_totals",
    "fm_sheet_get_pivot_report_layout",
    "fm_sheet_get_pivot_name",
    "fm_sheet_get_pivot_runtime_id",
    "fm_sheet_resolve_pivot_runtime_id",
    "fm_sheet_get_pivot_options",
    "fm_sheet_get_pivot_row_field_order",
    "fm_sheet_get_pivot_style_info",
    "fm_sheet_get_pivot_layout",
    "fm_sheet_project_pivot",
    "fm_sheet_is_pivot_projection_pending",
    "fm_sheet_remove_pivot",
    "fm_sheet_set_pivot_anchor",
    "fm_sheet_set_pivot_column_field_order",
    "fm_sheet_set_pivot_grand_totals",
    "fm_sheet_set_pivot_report_layout",
    "fm_sheet_set_pivot_name",
    "fm_sheet_set_pivot_options",
    "fm_sheet_set_pivot_row_field_order",
    "fm_sheet_set_pivot_style_info",
    "fm_sheet_show_pivot_details",
    "fm_sheet_get_precedents",
    "fm_get_image_info",
    "fm_workbook_get_diagnostics",
    "fm_workbook_has_dirty_cells",
    "fm_workbook_recalc",
    "fm_workbook_recalc_data_tables",
    "fm_workbook_get_ref_mode",
    "fm_workbook_remove_person",
    "fm_workbook_remove_sheet",
    "fm_workbook_rename_sheet",
    "fm_workbook_clear_theme",
    "fm_workbook_resolve_color",
    "fm_workbook_save",
    "fm_workbook_save_as",
    "fm_workbook_save_with_diagnostics",
    "fm_workbook_set_active_sheet",
    "fm_sheet_set_array_formula",
    "fm_sheet_set_blank",
    "fm_sheet_set_bool",
    "fm_workbook_set_calc_mode",
    "fm_workbook_set_calc_on_save",
    "fm_sheet_set_cell_phonetic",
    "fm_sheet_set_cell_phonetic_properties",
    "fm_sheet_set_cell_phonetic_runs",
    "fm_sheet_set_cell_show_phonetic",
    "fm_workbook_set_date1904",
    "fm_workbook_set_default_font",
    "fm_workbook_set_defined_name",
    "fm_sheet_set_error",
    "fm_workbook_set_profile_id",
    "fm_workbook_set_file_name",
    "fm_sheet_set_formula",
    "fm_workbook_set_full_precision",
    "fm_workbook_set_iterative",
    "fm_workbook_set_iterative_progress",
    "fm_sheet_set_number",
    "fm_workbook_set_pinned_now",
    "fm_convert_civil_to_serial",
    "fm_convert_serial_to_civil",
    "fm_workbook_set_workbook_protection",
    "fm_workbook_set_ref_mode",
    "fm_sheet_set_text",
    "fm_workbook_set_theme_colors",
    "fm_workbook_set_theme_fonts",
    "fm_workbook_set_theme_script_font",
    "fm_workbook_get_sheet_name",
    "fm_workbook_get_sheet_index",
    "fm_sheet_get_spill_info",
    "fm_workbook_get_table",
    "fm_workbook_get_table_runtime_id",
    "fm_workbook_resolve_table_runtime_id",
    "fm_workbook_add_table",
    "fm_workbook_get_table_column",
    "fm_workbook_get_table_info",
    "fm_workbook_remove_table",
    "fm_workbook_rename_table",
    "fm_workbook_rename_table_column",
    "fm_workbook_set_table_column_totals",
    "fm_workbook_set_table_style_info",
    "fm_workbook_update_table",
    "fm_workbook_get_theme_script_font_at",
    "fm_workbook_get_theme_script_font",
    "fm_workbook_get_theme_script_font_count",
    "fm_cell_nodes_get_node_count",
    "fm_cf_results_get_cell_count",
    "fm_get_function_count",
    "fm_get_locale_error_name_count",
    "fm_pagination_get_column_break_count",
    "fm_pagination_get_page_count",
    "fm_pagination_get_print_area_count",
    "fm_pagination_get_row_break_count",
    "fm_pivot_cells_get_cell_count",
    "fm_sheet_get_column_break_count",
    "fm_sheet_get_row_break_count",
    "fm_workbook_get_defined_name_count",
    "fm_workbook_get_passthrough_count",
    "fm_workbook_get_sheet_count",
    "fm_workbook_get_table_count",
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
        self._export_wrappers: dict = {}
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
                StatusCode.INTERNAL_ERROR,
                _diagnostic_override=(self._poison_message, ""),
                _host_only=True,
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
                StatusCode.INTERNAL_ERROR,
                _diagnostic_override=(self._poison_message, ""),
                _host_only=True,
            ) from trap

    def __getattr__(self, name: str):
        self._ensure()
        with self._call_lock:
            # A few diagnostics tests construct this small wrapper with
            # object.__new__ so they can inject fake exports. Keep the lazy
            # default here as well as in __init__ for that supported seam.
            wrappers = self.__dict__.setdefault("_export_wrappers", {})
            cached = wrappers.get(name)
            if cached is not None:
                return cached
            fn = self._exports.get(name)
            if fn is None:
                raise AttributeError(f"WASM export '{name}' not found")
            lock = self._call_lock
            captures_diagnostics = name in _STATUS_RETURNING_EXPORTS
            captures_parse_error = name != "fm_get_last_parse_error"
            parse_slot = self._parse_error_slot()

            # Wrap so the caller can use a ctypes-like call syntax. Cache the
            # wrapper: besides avoiding a closure allocation on every lookup,
            # this makes every call share the same poison and diagnostics
            # trampoline for the life of this module instance.
            def _wrapped(*args):
                with lock:
                    if captures_diagnostics:
                        # The C ABI's diagnostic buffers are thread-local in
                        # the engine, but this WASM instance has no pthreads.
                        # Keep a Python-side snapshot per caller and clear it
                        # before every status-returning call so success cannot
                        # leave a previous failure pending.
                        self._last_diagnostic.value = None
                        parse_slot.value = None
                    result = self._call_or_poison(fn, name, *args)
                    # Only status-returning exports are allowed to populate
                    # the snapshot. Counts, indices, and pointers are all
                    # i32 in WASM and must never be mistaken for a failure
                    # status.
                    if captures_diagnostics and result != 0:
                        message_fn = self._exports.get("fm_get_last_error_message")
                        context_fn = self._exports.get("fm_get_last_error_context")
                        if message_fn is not None and context_fn is not None:
                            self._last_diagnostic.value = (
                                result,
                                self._read_cstr_unlocked(message_fn(self._store)),
                                self._read_cstr_unlocked(context_fn(self._store)),
                            )
                        if captures_parse_error and _SYNTAX_STATUS_FIRST <= result <= _SYNTAX_STATUS_LAST:
                            parse_slot.value = (result, self._read_parse_error_unlocked())
                    return result

            # Tracebacks and reprs name the C ABI entry point, not this
            # binding-internal closure.
            _wrapped.__name__ = name
            _wrapped.__qualname__ = name
            wrappers[name] = _wrapped
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

    def write_bytes(self, ptr: int, data: Union[bytes, bytearray]) -> None:
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

        Returns the empty string when ``ptr`` is 0. The memory/instance
        check happens before taking ``_call_lock``; the scan itself is
        delegated to the unlocked helper while the lock is held so a
        concurrent WASM call cannot grow memory mid-read.
        """
        if ptr == 0:
            return ""
        self._ensure()
        with self._call_lock:
            return self._read_cstr_unlocked(ptr)

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

    def _read_parse_error_unlocked(self) -> Optional[ParseError]:
        """Read ``fm_get_last_parse_error`` while the caller owns ``_call_lock``.

        The engine clears the record on the next ``fm_*`` call, so it is read
        in the failing call's critical section.
        """
        getter = self._exports.get("fm_get_last_parse_error")
        if getter is None:
            return None
        ptr = self._call_or_poison(self._exports["malloc"], "malloc", PARSE_ERROR.size)
        if ptr == 0:
            return None
        try:
            assert self._memory is not None
            self._memory.write(self._store, PARSE_ERROR.blank(), ptr)
            if self._call_or_poison(getter, "fm_get_last_parse_error", ptr) != 0:
                return None
            raw = bytes(self._memory.read(self._store, ptr, ptr + PARSE_ERROR.size))
        finally:
            self._call_or_poison(self._exports["free"], "free", ptr)
        d = PARSE_ERROR.unpack_from(raw)
        return ParseError(ParseErrorCode(d["code"]), d["start"], d["end"], d["line"], d["column"])

    def _parse_error_slot(self) -> threading.local:
        # Lazy for the same seam as `_export_wrappers`: tests build this
        # wrapper with object.__new__.
        return self.__dict__.setdefault("_last_parse_error", threading.local())

    def last_parse_error(self, status: int) -> Optional[ParseError]:
        """Take the parse error captured for ``status`` once, if any."""
        slot = self._parse_error_slot()
        snapshot = getattr(slot, "value", None)
        slot.value = None
        if snapshot is not None and snapshot[0] == status:
            return snapshot[1]
        return None

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
