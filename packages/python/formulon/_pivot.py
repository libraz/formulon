"""The :class:`Pivot` handle: every operation that targets one PivotTable of a sheet.

A :class:`~formulon.Sheet` hands out one :class:`Pivot` per pivot table and keeps
it pointing at the same pivot across edits that move it within the sheet's list;
see :class:`Pivot`.
"""

from __future__ import annotations

import inspect
import struct
from typing import TYPE_CHECKING, Any, Dict, List, Optional, Sequence, Tuple, Union, overload

from . import _structs as S
from ._c import LIB, FormulonError, _NoCopy, _sint, _uint
from ._handle_tables import CELL_ARGS
from ._status import StatusCode
from ._workbook import (
    _PATCH_CLEAR,
    _PATCH_SET,
    _PIVOT_NO_DATA_FIELD,
    _PIVOT_VALUES_FIELD_C,
    _STATUS_NOT_FOUND,
    PIVOT_VALUES_FIELD,
    UNCHANGED,
    CellAddress,
    PivotAggregation,
    PivotAxis,
    PivotCalendar,
    PivotCell,
    PivotCellKind,
    PivotDataField,
    PivotDataFieldInput,
    PivotDateGroupInput,
    PivotDatePeriod,
    PivotField,
    PivotFieldInput,
    PivotFilter,
    PivotFilterInput,
    PivotFilterKind,
    PivotGrandTotals,
    PivotItem,
    PivotLayout,
    PivotOptions,
    PivotReportLayout,
    PivotShowValuesAs,
    PivotSortOrder,
    PivotStyleOptions,
    Unchanged,
    Value,
    _ArgScope,
    _check,
    _decode_cell_rect_at,
    _decode_pivot_filter_operand,
    _encode_pivot_filter_operands,
    _install_workbook_locks,
    _nullable_enum,
    _pack_pivot_date_group,
    _pack_pivot_field_patch,
    _pack_pivot_item_patch,
    _pack_pivot_style_options_patch,
    _read_count,
    _read_cstr_out,
    _tristate_default,
    to_native,
)

if TYPE_CHECKING:
    from ._pivot_cache import PivotCache
    from ._sheet import Sheet
    from ._workbook import PivotProperties

__all__ = ["Pivot"]

# The PivotTable Options :meth:`Pivot.update_options` writes, by C field kind.
_PIVOT_OPTION_FLAGS = (
    "show_error",
    "show_missing",
    "show_headers",
    "show_drill",
    "show_data_tips",
    "grid_drop_zones",
    "show_empty_row",
    "show_empty_column",
    "show_data_drop_down",
    "print_drill",
    "field_print_titles",
    "item_print_titles",
    "merge_item",
    "page_over_then_down",
    "multiple_field_filters",
    "preserve_formatting",
    "custom_list_sort",
    "auto_formatting",
    "enable_drill",
)
_PIVOT_OPTION_TEXTS = ("error_caption", "missing_caption", "alt_text", "alt_text_summary", "grand_total_caption")
_PIVOT_OPTION_NAMES = frozenset((*_PIVOT_OPTION_FLAGS, *_PIVOT_OPTION_TEXTS, "indent", "page_wrap"))


def _flag(value: Any, name: str) -> int:
    """``value`` as a C 0/1 flag; ``TypeError`` unless it is a ``bool``."""
    from ._sheet import _flag as flag

    return flag(value, name)


class Pivot(_NoCopy):
    """One PivotTable of a :class:`Sheet`; every operation on the pivot lives here.

    Obtain a handle from :meth:`Sheet.pivot`, :meth:`Sheet.list_pivots` or
    :meth:`Sheet.add_pivot`. A sheet hands out one handle per pivot, so every
    lookup of the same pivot returns the same object. The handle follows its
    pivot when pivots are added or removed, rows or columns are edited, or the
    pivot is renamed, and survives an edit that fails. Once the pivot is
    removed (directly, by a row or column delete over it, with its sheet, or
    with its cache's source sheet), every operation but reading :attr:`sheet`
    raises :class:`NotFoundError`. Remove a pivot with :meth:`Sheet.remove_pivot`.
    The definition read-backs work for pivots created in-session and loaded
    from a package alike.
    """

    __slots__ = ("_sheet", "_runtime_id")

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain a Pivot from Sheet.pivot() or Sheet.add_pivot()")

    @classmethod
    def _new(cls, sheet: Sheet, runtime_id: int) -> "Pivot":
        handle = object.__new__(cls)
        handle._sheet = sheet
        handle._runtime_id = runtime_id
        return handle

    def _locate(self) -> Tuple[int, int, int]:
        """The workbook handle, the sheet index and the pivot's current index.

        Raises:
          NotFoundError: when the sheet or the pivot was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            status = LIB.fm_sheet_resolve_pivot_runtime_id(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(self._runtime_id, "id"), out
            )
            if status == _STATUS_NOT_FOUND:
                raise FormulonError(StatusCode.NOT_FOUND, "the pivot was removed", "")
            _check(status)
            return h, sheet, LIB.read_u32(out)

    # -- Identification ----------------------------------------------------
    @property
    def sheet(self) -> Sheet:
        """The sheet this pivot is anchored on."""
        return self._sheet

    @property
    def index(self) -> int:
        """The pivot's current 0-based index on its sheet (its position in :meth:`Sheet.list_pivots`).

        Raises:
          NotFoundError: when the pivot or its sheet was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._locate()[2]

    @property
    def name(self) -> str:
        """The pivot's current name.

        Raises:
          NotFoundError: when the pivot or its sheet was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        h, sheet, pivot_index = self._locate()
        return _read_cstr_out(
            LIB.fm_sheet_get_pivot_name,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )

    def _cache_id(self) -> int:
        h, sheet, pivot_index = self._locate()
        return _read_count(
            LIB.fm_sheet_get_pivot_cache_id,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )

    def get_properties(self) -> PivotProperties:
        """Read the pivot's name and the id of the cache it draws from."""
        from ._workbook import PivotProperties

        return PivotProperties(name=self.name, cache_id=self._cache_id())

    def rename(self, name: str) -> None:
        """Rename the pivot table."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(
                LIB.fm_sheet_rename_pivot(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    name_ptr,
                )
            )

    def get_cache(self) -> PivotCache:
        """The handle of the cache the pivot draws from."""
        return self._sheet._workbook.pivot_cache(self._cache_id())

    # -- Layout ------------------------------------------------------------
    def get_layout(self) -> PivotLayout:
        """Return the pivot's layout, computed from its cache records; the sheet cells are not written (:meth:`project` writes them)."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            handle = 0
            try:
                _check(
                    LIB.fm_sheet_get_pivot_layout(
                        h,
                        _uint(sheet, "sheet_index", role="sheet"),
                        _uint(pivot_index, "pivot_index", role="position"),
                        out,
                    )
                )
                handle = LIB.read_u32(out)
                present_p = scope.out_ptr()
                rect_p = scope.struct(S.CELL_RECT)
                _check(LIB.fm_pivot_report_get_bounds(handle, present_p, rect_p))
                rect = _decode_cell_rect_at(rect_p) if LIB.read_u32(present_p) else None
                cells: List[PivotCell] = []
                n = _read_count(LIB.fm_pivot_report_get_cell_count, handle)
                # One scratch block for the whole walk.
                cptr = scope.struct(S.PIVOT_CELL)
                for i in range(n):
                    S.zero_struct(LIB, S.PIVOT_CELL, cptr)
                    _check(LIB.fm_pivot_report_get_cell_at(handle, _uint(i, "idx", role="position"), cptr))
                    d = S.PIVOT_CELL.unpack(LIB, cptr)
                    value = Value._from_wasm(cptr + S.PIVOT_CELL_VALUE_OFFSET)
                    cells.append(
                        PivotCell(
                            row=d["row"],
                            col=d["col"],
                            value=to_native(value),
                            kind=PivotCellKind(d["kind"]),
                            depth=d["depth"],
                            field_name=LIB.read_optional_cstr(d["field_name"]),
                            number_format=LIB.read_optional_cstr(d["number_format"]),
                        )
                    )
                return PivotLayout(range=rect, cells=cells)
            finally:
                if handle:
                    LIB.fm_pivot_report_destroy(handle)

    def set_anchor(self, cell: Union[CellAddress, str]) -> None:
        """Move the pivot so its top-left cell is ``cell``, a :class:`CellAddress` or one A1 cell (``"C3"``).

        It has the meaning of ``PivotInput.anchor``.

        Raises:
          TypeError: when ``cell`` is neither a :class:`CellAddress` nor a ``str``.
          InvalidArgumentError: when the text is not one unqualified A1 cell.
        """
        if isinstance(cell, CellAddress):
            row, col = cell.row, cell.col
        elif isinstance(cell, str):
            row, col = self._sheet._cell(cell)
        else:
            raise TypeError(f"formulon: cell must be a CellAddress or an A1 string, got {type(cell).__name__}")
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_anchor(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(row, "anchor_row", role="coordinate"),
                _uint(col, "anchor_col", role="coordinate"),
            )
        )

    def get_grand_totals(self) -> PivotGrandTotals:
        """Return the row and column flags :meth:`update_grand_totals` writes."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            rows = scope.out_ptr()
            cols = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_pivot_grand_totals(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    rows,
                    cols,
                )
            )
            return PivotGrandTotals(rows=LIB.read_i32(rows) != 0, columns=LIB.read_i32(cols) != 0)

    def update_grand_totals(
        self, *, rows: Union[bool, Unchanged] = UNCHANGED, columns: Union[bool, Unchanged] = UNCHANGED
    ) -> None:
        """Toggle the row / column grand-total bands; an omitted argument keeps its band.

        Raises:
          TypeError: when a given flag is not a ``bool``.
        """
        for name, value in (("rows", rows), ("columns", columns)):
            if value is not UNCHANGED and not isinstance(value, bool):
                raise TypeError(f"formulon: {name} must be a bool, got {type(value).__name__}")
        values: Dict[str, object] = {}
        for name, value in (("rows", rows), ("columns", columns)):
            if value is not UNCHANGED:
                values[name + "_op"] = _PATCH_SET
                values[name] = 1 if value else 0
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_GRAND_TOTALS_PATCH)
            S.PIVOT_GRAND_TOTALS_PATCH.pack(LIB, ptr, values)
            _check(
                LIB.fm_sheet_update_pivot_grand_totals(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                )
            )

    def get_style_options(self) -> Optional[PivotStyleOptions]:
        """Return the table style :meth:`update_style_options` writes, or ``None`` without a style element."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.PIVOT_STYLE_OPTIONS)
            _check(
                LIB.fm_sheet_get_pivot_style_options(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    present,
                    ptr,
                )
            )
            if LIB.read_i32(present) == 0:
                return None
            d = S.PIVOT_STYLE_OPTIONS.unpack(LIB, ptr)
            return PivotStyleOptions(
                style_name=LIB.read_cstr(d["style_name"]) or None,
                show_row_headers=d["show_row_headers"] != 0,
                show_column_headers=d["show_column_headers"] != 0,
                show_row_stripes=d["show_row_stripes"] != 0,
                show_column_stripes=d["show_column_stripes"] != 0,
                show_last_column=d["show_last_column"] != 0,
            )

    def update_style_options(
        self,
        *,
        style_name: Union[str, None, Unchanged] = UNCHANGED,
        show_row_headers: Union[bool, Unchanged] = UNCHANGED,
        show_column_headers: Union[bool, Unchanged] = UNCHANGED,
        show_row_stripes: Union[bool, Unchanged] = UNCHANGED,
        show_column_stripes: Union[bool, Unchanged] = UNCHANGED,
        show_last_column: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the pivot's table style, making the style element present; an omitted argument keeps its value.

        ``style_name=None`` writes the element without a style name. On a
        pivot without a style element the arguments left out take the creation
        defaults: the workbook's PivotTable style default (else
        ``PivotStyleLight16``) and row / column headers and last column on.
        The whole update is checked before anything changes.

        Raises:
          TypeError: when a toggle is not a ``bool`` (``None`` included) or
            ``style_name`` is neither a ``str`` nor ``None``.
        """
        patch = {
            key: value
            for key, value in (
                ("style_name", style_name),
                ("show_row_headers", show_row_headers),
                ("show_column_headers", show_column_headers),
                ("show_row_stripes", show_row_stripes),
                ("show_column_stripes", show_column_stripes),
                ("show_last_column", show_last_column),
            )
            if value is not UNCHANGED
        }
        for key, value in patch.items():
            if key == "style_name":
                if value is not None and not isinstance(value, str):
                    raise TypeError(f"formulon: style_name must be a str or None, got {type(value).__name__}")
            else:
                _flag(value, key)
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = _pack_pivot_style_options_patch(patch, scope)
            _check(
                LIB.fm_sheet_update_pivot_style_options(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                )
            )

    def get_report_layout(self) -> PivotReportLayout:
        """Return the pivot's compact, tabular, or outline report layout."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_pivot_report_layout(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    out,
                )
            )
            return PivotReportLayout(LIB.read_i32(out))

    def set_report_layout(self, layout: PivotReportLayout) -> None:
        """Set the pivot's compact, tabular, or outline report layout."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_report_layout(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _sint(layout, "layout"),
            )
        )

    # -- Options and projection --------------------------------------------
    def get_options(self) -> PivotOptions:
        """Read the PivotTable Options of the pivot; see :class:`PivotOptions`."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_OPTIONS)
            _check(
                LIB.fm_sheet_get_pivot_options(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                )
            )
            d = S.PIVOT_OPTIONS.unpack(LIB, ptr)

            return PivotOptions(
                show_error=d["show_error"] != 0,
                error_caption=LIB.read_cstr(d["error_caption"]) or None,
                show_missing=d["show_missing"] != 0,
                missing_caption=LIB.read_cstr(d["missing_caption"]) or None,
                show_headers=d["show_headers"] != 0,
                show_drill=d["show_drill"] != 0,
                show_data_tips=d["show_data_tips"] != 0,
                grid_drop_zones=d["grid_drop_zones"] != 0,
                show_empty_row=d["show_empty_row"] != 0,
                show_empty_column=d["show_empty_column"] != 0,
                show_data_drop_down=d["show_data_drop_down"] != 0,
                print_drill=d["print_drill"] != 0,
                field_print_titles=d["field_print_titles"] != 0,
                item_print_titles=d["item_print_titles"] != 0,
                merge_item=d["merge_item"] != 0,
                indent=d["indent"],
                page_over_then_down=d["page_over_then_down"] != 0,
                page_wrap=d["page_wrap"],
                multiple_field_filters=d["multiple_field_filters"] != 0,
                preserve_formatting=d["preserve_formatting"] != 0,
                custom_list_sort=d["custom_list_sort"] != 0,
                alt_text=LIB.read_cstr(d["alt_text"]) or None,
                alt_text_summary=LIB.read_cstr(d["alt_text_summary"]) or None,
                grand_total_caption=LIB.read_cstr(d["grand_total_caption"]) or None,
                auto_formatting=d["auto_formatting"] != 0,
                enable_drill=d["enable_drill"] != 0,
            )

    def update_options(
        self,
        *,
        show_error: Union[bool, Unchanged] = UNCHANGED,
        error_caption: Union[str, None, Unchanged] = UNCHANGED,
        show_missing: Union[bool, Unchanged] = UNCHANGED,
        missing_caption: Union[str, None, Unchanged] = UNCHANGED,
        show_headers: Union[bool, Unchanged] = UNCHANGED,
        show_drill: Union[bool, Unchanged] = UNCHANGED,
        show_data_tips: Union[bool, Unchanged] = UNCHANGED,
        grid_drop_zones: Union[bool, Unchanged] = UNCHANGED,
        show_empty_row: Union[bool, Unchanged] = UNCHANGED,
        show_empty_column: Union[bool, Unchanged] = UNCHANGED,
        show_data_drop_down: Union[bool, Unchanged] = UNCHANGED,
        print_drill: Union[bool, Unchanged] = UNCHANGED,
        field_print_titles: Union[bool, Unchanged] = UNCHANGED,
        item_print_titles: Union[bool, Unchanged] = UNCHANGED,
        merge_item: Union[bool, Unchanged] = UNCHANGED,
        indent: Union[int, Unchanged] = UNCHANGED,
        page_over_then_down: Union[bool, Unchanged] = UNCHANGED,
        page_wrap: Union[int, Unchanged] = UNCHANGED,
        multiple_field_filters: Union[bool, Unchanged] = UNCHANGED,
        preserve_formatting: Union[bool, Unchanged] = UNCHANGED,
        custom_list_sort: Union[bool, Unchanged] = UNCHANGED,
        alt_text: Union[str, None, Unchanged] = UNCHANGED,
        alt_text_summary: Union[str, None, Unchanged] = UNCHANGED,
        grand_total_caption: Union[str, None, Unchanged] = UNCHANGED,
        auto_formatting: Union[bool, Unchanged] = UNCHANGED,
        enable_drill: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the PivotTable Options; an omitted argument keeps its value.

        Arguments are the :class:`PivotOptions` fields. A caption or alt text
        given as ``None`` (or ``""``) is cleared; a flag or a number cannot be.
        A wrong-typed argument fails without changing the pivot.

        Raises:
          TypeError: when a flag is not a ``bool``, ``indent`` or
            ``page_wrap`` is not an integer, or a string argument is neither a
            ``str`` nor ``None``.
          FormulonError: with ``code == StatusCode.PRECONDITION_FAILED`` when the
            report would overlap another structure.
        """
        arguments = dict(locals())
        given = {name: arguments[name] for name in _PIVOT_OPTION_NAMES if arguments[name] is not UNCHANGED}
        for name, value in given.items():
            if name in _PIVOT_OPTION_TEXTS:
                if value is not None and not isinstance(value, str):
                    raise TypeError(f"formulon: {name} must be a str or None, got {type(value).__name__}")
            elif name in _PIVOT_OPTION_FLAGS:
                _flag(value, name)
            elif value is None:
                raise TypeError(f"formulon: {name} cannot be None")
        h, sheet, pivot_index = self._locate()
        values: Dict[str, object] = {}
        with _ArgScope() as scope:
            for name, value in given.items():
                if name in _PIVOT_OPTION_TEXTS:
                    if value is None:
                        values[name + "_op"] = _PATCH_CLEAR
                        continue
                    values[name] = scope.cstr(value)
                elif name in _PIVOT_OPTION_FLAGS:
                    values[name] = 1 if value else 0
                else:
                    values[name] = _uint(value, name)
                values[name + "_op"] = _PATCH_SET
            ptr = scope.struct(S.PIVOT_OPTIONS_PATCH)
            S.PIVOT_OPTIONS_PATCH.pack(LIB, ptr, values)
            _check(
                LIB.fm_sheet_update_pivot_options(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                )
            )

    def is_projection_pending(self) -> bool:
        """Return whether a cache change left the pivot's sheet cells stale; see :meth:`project`."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_is_pivot_projection_pending(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    out,
                )
            )
            return LIB.read_i32(out) != 0

    def project(self) -> None:
        """Write the pivot's report, computed from its cache records, into its sheet cells now and clear its pending flag.

        It does not re-read the cache's source range. Raises when the report would overlap another structure; the pivot stays pending.
        """
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_project_pivot(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(pivot_index, "pivot_index", role="position")
            )
        )

    @overload
    def show_details(self, row: int, col: int) -> Sheet: ...

    @overload
    def show_details(self, a1: str) -> Sheet: ...

    def show_details(self, *args: Any, **kwargs: Any) -> Sheet:
        """Write the source records behind data cell (``row``, ``col``) onto a new sheet and return its handle.

        The cell is a ``(row, col)`` pair or one unqualified A1 string. The
        sheet is inserted immediately before the pivot's sheet and becomes the
        active sheet; the pivot's sheet and the ones after it move up one
        place, and every existing :class:`Sheet` handle keeps pointing at its
        sheet.
        """
        from ._sheet import _split_cell_args

        cells, _ = _split_cell_args(args, kwargs, CELL_ARGS["Pivot.show_details"], ())
        cell = cells[0]
        row, col = self._sheet._cell(cell) if isinstance(cell, str) else cell
        h, sheet, pivot_index = self._locate()
        workbook = self._sheet._workbook
        handles = workbook._sheet_handles()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_show_pivot_details(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    out,
                )
            )
            index = LIB.read_u32(out)
        return workbook._insert_sheet_handle(handles, index)

    # -- Fields ------------------------------------------------------------
    def list_fields(self) -> List[PivotField]:
        """Every field of the pivot, in cache-field order; a field's position is its ``field_index``."""
        located = self._locate()
        h, sheet, pivot_index = located
        count = _read_count(
            LIB.fm_sheet_get_pivot_field_count,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )
        return [self._field_at(located, i) for i in range(count)]

    @staticmethod
    def _field_at(located: Tuple[int, int, int], field_index: int) -> PivotField:
        """Read the definition of pivot field ``field_index``."""
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FIELD_INFO)
            _check(
                LIB.fm_sheet_get_pivot_field_at(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(field_index, "field_index", role="position"),
                    ptr,
                )
            )
            d = S.PIVOT_FIELD_INFO.unpack(LIB, ptr)
            return PivotField(
                source_name=LIB.read_cstr(d["source_name"]),
                custom_name=LIB.read_cstr(d["custom_name"]) or None,
                axis=PivotAxis(d["axis"]) if d["axis"] >= 0 else None,
                subtotal_top=d["subtotal_top"] != 0,
                number_format=LIB.read_cstr(d["number_format"]) or None,
                cache_field_index=d["cache_field_index"] if d["has_cache_field_index"] else None,
                sort_order=PivotSortOrder(d["sort_order"]),
                sort_by_field=LIB.read_cstr(d["sort_by_field"]) or None,
                date_periods=PivotDatePeriod(d["date_periods"]),
                date_calendar=PivotCalendar(d["date_calendar"]),
                group_base_index=d["group_base_index"] if d["has_group_base_index"] else None,
                default_subtotal=d["default_subtotal"] != 0,
                subtotal_functions=[agg for agg in PivotAggregation if d["subtotal_functions"] & (1 << int(agg))],
                insert_blank_row=d["insert_blank_row"] != 0,
                fill_down_labels=d["fill_down_labels"] != 0,
                show_all=d["show_all"] != 0,
                show_detail=d["show_detail"] != 0,
            )

    def add_field(self, spec: PivotFieldInput) -> int:
        """Place a field on the pivot; return its index.

        A pivot's fields are positional over its cache's fields (field ``i``
        is cache field ``i``). When ``source_name`` names a cache field the
        pivot first gets one bare, axis-less field per cache field and the
        named one is configured in place, so the index is that cache field's
        position and adding it again updates it. A name the cache lacks
        appends a field. A field added on the row or column axis is appended
        to that axis's field order.
        """
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FIELD_SPEC)
            out = scope.out_ptr()
            S.PIVOT_FIELD_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": _sint(spec.axis, "axis"),
                    "subtotal_top": _tristate_default(spec.subtotal_top),
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FIELD_SPEC, "source_name", spec.source_name, scope)
            S.write_str_field(LIB, ptr, S.PIVOT_FIELD_SPEC, "custom_name", spec.custom_name, scope)
            S.write_str_field(
                LIB,
                ptr,
                S.PIVOT_FIELD_SPEC,
                "number_format",
                spec.number_format,
                scope,
            )
            _check(
                LIB.fm_sheet_add_pivot_field(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                    out,
                )
            )
            return LIB.read_u32(out)

    def update_field(
        self,
        field_index: int,
        *,
        custom_name: Union[str, None, Unchanged] = UNCHANGED,
        number_format: Union[str, None, Unchanged] = UNCHANGED,
        default_subtotal: Union[bool, Unchanged] = UNCHANGED,
        insert_blank_row: Union[bool, Unchanged] = UNCHANGED,
        fill_down_labels: Union[bool, Unchanged] = UNCHANGED,
        show_all: Union[bool, Unchanged] = UNCHANGED,
        subtotal_top: Union[bool, Unchanged] = UNCHANGED,
        sort_order: Union[PivotSortOrder, Unchanged] = UNCHANGED,
        sort_by_field: Union[str, None, Unchanged] = UNCHANGED,
        show_detail: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update pivot field ``field_index``; an omitted argument keeps its value.

        ``custom_name`` renames the field in the report; ``None`` restores its
        source name. It may be the field's own source name, not another
        field's source or custom name or a data field's name.
        ``number_format`` is the number format code of the field's items;
        ``None`` clears it. ``default_subtotal`` turns the automatic subtotal
        on or off (custom subtotal functions replace it either way);
        ``insert_blank_row`` adds a blank row after each item block;
        ``fill_down_labels`` repeats item labels on every row of a block;
        ``show_all`` lists items without data. ``sort_order`` is the sibling
        order of the items; ``PivotSortOrder.MANUAL`` keeps them in
        :meth:`add_field_item` order and rejects a ``sort_by_field``, the data
        field the items sort by (``None`` sorts by the item labels).
        ``show_detail`` expands or collapses every item of the field together
        with its subtotal row.

        The whole update is checked first and applied at once: a rejected
        argument leaves every field as it was.

        Raises:
          TypeError: when a flag is not a ``bool`` (``None`` included) or a
            string argument is neither a ``str`` nor ``None``.
        """
        patch = {
            key: value
            for key, value in (
                ("custom_name", custom_name),
                ("number_format", number_format),
                ("default_subtotal", default_subtotal),
                ("insert_blank_row", insert_blank_row),
                ("fill_down_labels", fill_down_labels),
                ("show_all", show_all),
                ("subtotal_top", subtotal_top),
                ("sort_order", sort_order),
                ("sort_by_field", sort_by_field),
                ("show_detail", show_detail),
            )
            if value is not UNCHANGED
        }
        for key, value in patch.items():
            if key in ("custom_name", "number_format", "sort_by_field"):
                if value is not None and not isinstance(value, str):
                    raise TypeError(f"formulon: {key} must be a str or None, got {type(value).__name__}")
            elif key != "sort_order":
                _flag(value, key)
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = _pack_pivot_field_patch(patch, scope)
            _check(
                LIB.fm_sheet_update_pivot_field(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(field_index, "field_index", role="position"),
                    ptr,
                )
            )

    def clear_fields(self) -> None:
        """Drop every field from the pivot, emptying the row and column orders."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_fields(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(pivot_index, "pivot_index", role="position")
            )
        )

    def set_field_axis(self, field_index: int, axis: Optional[PivotAxis]) -> None:
        """Set the axis of pivot field ``field_index``.

        Moving onto the row or column axis appends the field to that axis's
        order; moving off removes it from the order it was in. ``None`` takes
        the field off every axis, as unchecking it in a field list does: its
        data fields are removed (later data-field indices shift down), as are
        filters naming it and value filters scoring a removed data field; its
        item visibility is kept.
        """
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_field_axis(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(field_index, "field_index", role="position"),
                _sint(_nullable_enum(axis, "axis"), "axis"),
            )
        )

    def set_field_date_group(self, field_index: int, group: Optional[PivotDateGroupInput]) -> None:
        """Group the date field ``field_index`` by the periods of ``group.periods``, or ungroup it with ``None``.

        ``field_index`` is a source field or one of its group fields; any grouping
        it has is replaced. Each period becomes a group field appended to the
        pivot cache, finest first, named and labelled in the workbook's locale;
        this pivot shows them in the field's place, coarsest outermost, and
        another pivot showing the field shows them before it.
        ``PivotCalendar.JAPANESE`` groups the years by era year.
        ``group.interval_days`` above 1 groups by runs of days and needs
        ``periods`` to be ``PivotDatePeriod.DAYS`` alone. ``start`` / ``end`` are
        a ``date`` or ``datetime`` (a naive local wall clock; an aware one
        is a ValueError); ``None`` takes the field's first value and its last
        (the midnight after it for a date without a time). An empty or unknown mask, a zero
        interval, an invalid bound, a start after the end, or a field holding
        anything but dates raises ``FormulonError``.

        ``None`` removes the grouping: the group fields leave every pivot on the
        cache and the field returns to the axis position they held; a field that
        is not grouped is left as it is.

        Raises:
          TypeError: when ``group`` is neither ``None`` nor a :class:`PivotDateGroupInput`.
          FormulonError: with ``code == StatusCode.PRECONDITION_FAILED`` when ``group`` is
            ``None`` and the group fields are not the last fields of the cache.
        """
        if group is not None and not isinstance(group, PivotDateGroupInput):
            raise TypeError(f"formulon: group must be a PivotDateGroupInput or None, got {type(group).__name__}")
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = (
                0
                if group is None
                else _pack_pivot_date_group(scope, group, date1904=self._sheet._workbook.get_date1904())
            )
            _check(
                LIB.fm_sheet_set_pivot_field_date_group(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(field_index, "field_index", role="position"),
                    ptr,
                )
            )

    # -- Field items and subtotals -----------------------------------------
    def list_field_items(self, field_index: int) -> List[PivotItem]:
        """Every item of pivot field ``field_index``, in item order; an item's position is its ``item_index``."""
        h, sheet, pivot_index = self._locate()
        count = _read_count(
            LIB.fm_sheet_get_pivot_field_item_count,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
            _uint(field_index, "field_index", role="position"),
        )
        return [self._field_item_at((h, sheet, pivot_index), field_index, i) for i in range(count)]

    def get_field_item(self, field_index: int, item_index: int) -> PivotItem:
        """Read item ``item_index`` of pivot field ``field_index``."""
        return self._field_item_at(self._locate(), field_index, item_index)

    @staticmethod
    def _field_item_at(located: Tuple[int, int, int], field_index: int, item_index: int) -> PivotItem:
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_ITEM_INFO)
            _check(
                LIB.fm_sheet_get_pivot_field_item_at(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(field_index, "field_index", role="position"),
                    _uint(item_index, "item_index", role="position"),
                    ptr,
                )
            )
            d = S.PIVOT_ITEM_INFO.unpack(LIB, ptr)
            return PivotItem(
                name=LIB.read_cstr(d["name"]) or None,
                cache_index=d["cache_index"] if d["has_cache_index"] else None,
                visible=d["visible"] != 0,
                caption=LIB.read_cstr(d["caption"]) if d["caption"] else None,
                show_detail=d["show_detail"] != 0,
            )

    def update_field_item(
        self,
        field_index: int,
        item_index: int,
        *,
        visible: Union[bool, Unchanged] = UNCHANGED,
        show_detail: Union[bool, Unchanged] = UNCHANGED,
        caption: Union[str, None, Unchanged] = UNCHANGED,
    ) -> None:
        """Update item ``item_index`` of pivot field ``field_index``; an omitted argument keeps its value.

        ``show_detail`` expands (``True``) or collapses (``False``) the item.
        ``caption`` is shown instead of the item's value; ``None`` clears it.
        A caption longer than 255 UTF-16 code units, or one matching the label
        another item of the field shows, is rejected. The whole update is
        checked first and applied at once: a rejected argument leaves the
        item as it was.

        Raises:
          TypeError: when a flag is not a ``bool`` (``None`` included) or
            ``caption`` is neither a ``str`` nor ``None``.
        """
        patch = {
            key: value
            for key, value in (("visible", visible), ("show_detail", show_detail), ("caption", caption))
            if value is not UNCHANGED
        }
        for key, value in patch.items():
            if key == "caption":
                if value is not None and not isinstance(value, str):
                    raise TypeError(f"formulon: caption must be a str or None, got {type(value).__name__}")
            else:
                _flag(value, key)
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = _pack_pivot_item_patch(patch, scope)
            _check(
                LIB.fm_sheet_update_pivot_field_item(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(field_index, "field_index", role="position"),
                    _uint(item_index, "item_index", role="position"),
                    ptr,
                )
            )

    def add_field_item(
        self,
        field_index: int,
        name: Optional[str] = None,
        *,
        cache_index: Optional[int] = None,
        visible: bool = True,
    ) -> int:
        """Append a manual-filter item to pivot field ``field_index``; return its item index.

        Give exactly one of ``name`` and ``cache_index``:

        - ``name`` addresses the item by label. It binds to the first shared
          item of the cache field that renders as ``name`` (unbound when none
          does), and records are matched by comparing their rendered label
          against it. An empty label cannot express the blank item, which has
          no label of its own.
        - ``cache_index`` addresses it by position in the bound cache field's
          shared items, the index space of OOXML ``<item x="N">``. The label is
          resolved from that shared item, so this is the form that can express
          the blank item. ``cache_index`` is not validated: an index that
          resolves to nothing yields an item with an empty label that filters
          nothing.

        Raises:
          TypeError: when both or neither of ``name`` and ``cache_index`` are given.
        """
        if (name is None) == (cache_index is None):
            raise TypeError("formulon: give exactly one of name and cache_index")
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            lead = (
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(field_index, "field_index", role="position"),
            )
            if name is not None:
                _check(LIB.fm_sheet_add_pivot_field_item(*lead, scope.cstr(name), 1 if visible else 0, out))
            else:
                _check(
                    LIB.fm_sheet_add_pivot_field_item_by_cache_index(
                        *lead, _uint(cache_index, "cache_index", role="position"), 1 if visible else 0, out
                    )
                )
            return LIB.read_u32(out)

    def clear_field_items(self, field_index: int) -> None:
        """Drop every manual-filter item from pivot field ``field_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_field_items(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(field_index, "field_index", role="position"),
            )
        )

    def add_field_subtotal_function(self, field_index: int, aggregation: PivotAggregation) -> None:
        """Append a subtotal function to pivot field ``field_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_add_pivot_field_subtotal_function(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(field_index, "field_index", role="position"),
                _sint(aggregation, "aggregation"),
            )
        )

    def clear_field_subtotal_functions(self, field_index: int) -> None:
        """Drop every subtotal function from pivot field ``field_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_field_subtotal_functions(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(field_index, "field_index", role="position"),
            )
        )

    # -- Axis orders -------------------------------------------------------
    def _get_field_order(self, fn) -> List[int]:
        h, sheet, pivot_index = self._locate()
        head = (h, _uint(sheet, "sheet_index", role="sheet"), _uint(pivot_index, "pivot_index", role="position"))
        with _ArgScope() as scope:
            out_count = scope.out_ptr()
            # Size the buffer first, then fetch.
            _check(fn(*head, 0, 0, out_count))
            count = LIB.read_u32(out_count)
            if count == 0:
                return []
            buf = scope.alloc(4 * count)
            _check(fn(*head, buf, count, out_count))
            raw = LIB.read_bytes(buf, 4 * count)
            return [PIVOT_VALUES_FIELD if v == _PIVOT_VALUES_FIELD_C else v for v in struct.unpack("<%dI" % count, raw)]

    def _set_field_order(self, fn, indices: Sequence[int]) -> None:
        h, sheet, pivot_index = self._locate()
        arr_ptr = 0
        with _ArgScope() as scope:
            if indices:
                arr_ptr = scope.alloc(4 * len(indices))
                buf = b"".join(
                    struct.pack(
                        "<I", _PIVOT_VALUES_FIELD_C if i == PIVOT_VALUES_FIELD else _uint(i, "index", role="position")
                    )
                    for i in indices
                )
                LIB.write_bytes(arr_ptr, buf)
            _check(
                fn(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    arr_ptr,
                    _uint(len(indices), "count"),
                )
            )

    def get_row_field_order(self) -> List[int]:
        """Return the row-axis field order (the list :meth:`set_row_field_order` takes)."""
        return self._get_field_order(LIB.fm_sheet_get_pivot_row_field_order)

    def set_row_field_order(self, indices: Sequence[int]) -> None:
        """Replace the row-axis field order with ``indices``.

        The list is the axis's whole membership: listed fields move onto the
        row axis, and fields left out come off it. Each index must be below
        the pivot's field count and may appear once; pass an empty list to
        clear. ``PIVOT_VALUES_FIELD`` (``-2``) places the Σ Values level.
        """
        self._set_field_order(LIB.fm_sheet_set_pivot_row_field_order, indices)

    def get_column_field_order(self) -> List[int]:
        """Return the column-axis field order. Same contract as rows."""
        return self._get_field_order(LIB.fm_sheet_get_pivot_column_field_order)

    def set_column_field_order(self, indices: Sequence[int]) -> None:
        """Replace the column-axis field order with ``indices``.

        Same contract as :meth:`set_row_field_order`.
        """
        self._set_field_order(LIB.fm_sheet_set_pivot_column_field_order, indices)

    # -- Data fields -------------------------------------------------------
    def list_data_fields(self) -> List[PivotDataField]:
        """Every data field of the pivot, in order; a data field's position is its ``data_field_index``."""
        h, sheet, pivot_index = self._locate()
        count = _read_count(
            LIB.fm_sheet_get_pivot_data_field_count,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )
        return [self._data_field_at((h, sheet, pivot_index), i) for i in range(count)]

    @staticmethod
    def _data_field_at(located: Tuple[int, int, int], data_field_index: int) -> PivotDataField:
        """Read data field ``data_field_index`` in the shape :meth:`add_data_field` accepts."""
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_DATA_FIELD_SPEC)
            _check(
                LIB.fm_sheet_get_pivot_data_field_at(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(data_field_index, "data_field_index", role="position"),
                    ptr,
                )
            )
            d = S.PIVOT_DATA_FIELD_SPEC.unpack(LIB, ptr)
            return PivotDataField(
                name=LIB.read_cstr(d["name"]),
                field_index=d["field_index"],
                aggregation=PivotAggregation(d["aggregation"]),
                number_format=LIB.read_cstr(d["number_format"]) or None,
                show_as=PivotShowValuesAs(d["show_as"]),
                show_as_base_field=d["show_as_base_field"] if d["has_show_as_base_field"] else None,
                show_as_base_item=d["show_as_base_item"] if d["has_show_as_base_item"] else None,
            )

    @staticmethod
    def _pack_data_field_spec(spec: PivotDataFieldInput, scope: _ArgScope) -> int:
        ptr = scope.struct(S.PIVOT_DATA_FIELD_SPEC)
        S.PIVOT_DATA_FIELD_SPEC.pack(
            LIB,
            ptr,
            {
                "field_index": _uint(spec.field_index, "field_index", role="position"),
                "aggregation": _sint(spec.aggregation, "aggregation"),
                "show_as": _sint(spec.show_as, "show_as"),
                "has_show_as_base_field": 0 if spec.show_as_base_field is None else 1,
                "show_as_base_field": 0
                if spec.show_as_base_field is None
                else _uint(spec.show_as_base_field, "show_as_base_field", role="position"),
                "has_show_as_base_item": 0 if spec.show_as_base_item is None else 1,
                "show_as_base_item": 0
                if spec.show_as_base_item is None
                else _uint(spec.show_as_base_item, "show_as_base_item", role="position"),
            },
        )
        S.write_str_field(LIB, ptr, S.PIVOT_DATA_FIELD_SPEC, "name", spec.name, scope)
        S.write_str_field(
            LIB,
            ptr,
            S.PIVOT_DATA_FIELD_SPEC,
            "number_format",
            spec.number_format,
            scope,
        )
        return ptr

    def add_data_field(self, spec: PivotDataFieldInput) -> int:
        """Append a data-field entry; return its index.

        An omitted or empty ``spec.name`` is derived from the workbook profile,
        as :meth:`get_data_field_default_name` reports it.
        """
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = self._pack_data_field_spec(spec, scope)
            _check(
                LIB.fm_sheet_add_pivot_data_field(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                    out,
                )
            )
            return LIB.read_u32(out)

    def set_data_field(self, data_field_index: int, spec: PivotDataFieldInput) -> None:
        """Replace the data-field entry at ``data_field_index`` in place.

        An omitted or empty ``spec.name`` is derived as in :meth:`add_data_field`.
        """
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = self._pack_data_field_spec(spec, scope)
            _check(
                LIB.fm_sheet_set_pivot_data_field(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(data_field_index, "data_field_index", role="position"),
                    ptr,
                )
            )

    def get_data_field_default_name(
        self,
        field_index: int,
        aggregation: PivotAggregation,
        exclude_data_field: Optional[int] = None,
    ) -> str:
        """Return the name a new data field over ``field_index`` and ``aggregation`` would get, without adding one.

        ``exclude_data_field`` is the entry :meth:`set_data_field` would
        replace, whose name is not counted as taken; ``None`` for a new data
        field.
        """
        h, sheet, pivot_index = self._locate()
        exclude = (
            _PIVOT_NO_DATA_FIELD if exclude_data_field is None else _uint(exclude_data_field, "exclude_data_field")
        )
        return _read_cstr_out(
            LIB.fm_sheet_get_pivot_data_field_default_name,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
            _uint(field_index, "field_index", role="position"),
            _sint(aggregation, "aggregation"),
            exclude,
        )

    def clear_data_fields(self) -> None:
        """Drop every data-field entry from the pivot."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_data_fields(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(pivot_index, "pivot_index", role="position")
            )
        )

    # -- Filters -----------------------------------------------------------
    def list_filters(self) -> List[PivotFilter]:
        """Every listed filter of the pivot, in order; a filter's position is its ``filter_index``.

        The filter list is saved as the definition's ``<filters>`` block. Every
        entry of a loaded block whose type the engine enumerates appears here,
        naming its field by source name, and :meth:`clear_filters` /
        :meth:`remove_filter` remove it. Only entries the model cannot express
        (an unknown type, a member-property or OLAP measure filter, criteria
        that do not decode) are not listed; they filter nothing and are saved
        verbatim in their original position.
        """
        h, sheet, pivot_index = self._locate()
        count = _read_count(
            LIB.fm_sheet_get_pivot_filter_count,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )
        return [self._filter_at((h, sheet, pivot_index), i) for i in range(count)]

    @staticmethod
    def _filter_at(located: Tuple[int, int, int], filter_index: int) -> PivotFilter:
        """Read the filter at ``filter_index`` without mutating it."""
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FILTER_SPEC)
            _check(
                LIB.fm_sheet_get_pivot_filter_at(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    _uint(filter_index, "filter_index", role="position"),
                    ptr,
                )
            )
            d = S.PIVOT_FILTER_SPEC.unpack(LIB, ptr)
            return PivotFilter(
                axis=PivotAxis(d["axis"]),
                field_name=LIB.read_cstr(d["field_name"]),
                type=PivotFilterKind(d["type"]),
                value=_decode_pivot_filter_operand(d, "value"),
                value_high=_decode_pivot_filter_operand(d, "value_high"),
                data_field_index=d["data_field_index"],
            )

    def add_filter(self, spec: PivotFilterInput) -> int:
        """Append a filter to the pivot and return its index; it is saved with it. See :meth:`list_filters`."""
        h, sheet, pivot_index = self._locate()
        operands = _encode_pivot_filter_operands(spec)
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FILTER_SPEC)
            S.PIVOT_FILTER_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": _sint(spec.axis, "axis"),
                    "type": _sint(spec.type, "type"),
                    "data_field_index": _uint(spec.data_field_index, "data_field_index", role="position"),
                    "value_kind": _sint(operands["value_kind"], "value_kind"),
                    "value_int": _sint(operands.get("value_int", 0), "value_int"),
                    "value_double": float(operands.get("value_double", 0.0)),
                    "value_high_kind": _sint(operands["value_high_kind"], "value_high_kind"),
                    "value_high_int": _sint(operands.get("value_high_int", 0), "value_high_int"),
                    "value_high_double": float(operands.get("value_high_double", 0.0)),
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "field_name", spec.field_name, scope)
            for name in ("value_text", "value_high_text"):
                S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, name, operands.get(name), scope, allow_empty=True)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_filter(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(pivot_index, "pivot_index", role="position"),
                    ptr,
                    out,
                )
            )
            return LIB.read_u32(out)

    def remove_filter(self, filter_index: int) -> None:
        """Remove the filter at ``filter_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_remove_pivot_filter(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(pivot_index, "pivot_index", role="position"),
                _uint(filter_index, "filter_index", role="position"),
            )
        )

    def clear_filters(self) -> None:
        """Drop every listed filter from the pivot."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_filters(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(pivot_index, "pivot_index", role="position")
            )
        )


def _show_details_coordinate_form(self, row: int, col: int) -> Sheet: ...


Pivot.show_details.__signature__ = inspect.signature(_show_details_coordinate_form)  # type: ignore[attr-defined]
del _show_details_coordinate_form

_install_workbook_locks(Pivot)
