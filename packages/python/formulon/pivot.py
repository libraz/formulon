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
from ._c import LIB, FormulonError, _sint, _uint
from ._handle_tables import CELL_ARGS
from ._status import StatusCode
from .workbook import (
    _PIVOT_NO_DATA_FIELD,
    _PIVOT_VALUES_FIELD_C,
    _STATUS_NOT_FOUND,
    PIVOT_VALUES_FIELD,
    UNCHANGED,
    CellRect,
    PivotAggregation,
    PivotAxis,
    PivotCalendar,
    PivotCell,
    PivotCellKind,
    PivotDataField,
    PivotDataFieldInput,
    PivotDatePeriod,
    PivotField,
    PivotFieldInput,
    PivotFilter,
    PivotFilterInput,
    PivotFilterType,
    PivotFilterValueKind,
    PivotGrandTotals,
    PivotItem,
    PivotLayout,
    PivotOptions,
    PivotReportLayout,
    PivotShowValuesAs,
    PivotSortOrder,
    PivotStyle,
    Value,
    _ArgScope,
    _check,
    _decode_cell_rect_at,
    _enum_or_none,
    _install_workbook_locks,
    _nullable_enum,
    _pack_cell_rect_at,
    _read_count,
    _read_cstr_out,
)

if TYPE_CHECKING:
    from .pivot_cache import PivotCache
    from .sheet import Sheet
    from .workbook import PivotProperties

__all__ = ["Pivot"]

# The PivotTable Options :meth:`Pivot.set_options` writes, by C field kind.
_PIVOT_OPTION_FLAGS = (
    "show_error",
    "show_missing",
    "show_headers",
    "show_drill",
    "show_data_tips",
    "grid_drop_zones",
    "show_empty_row",
    "show_empty_col",
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
    from .sheet import _flag as flag

    return flag(value, name)


class Pivot:
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

    def __init__(self, sheet: Sheet, runtime_id: int) -> None:
        self._sheet = sheet
        self._runtime_id = runtime_id

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
                h, _uint(sheet, "sheet_index"), _uint(self._runtime_id, "id"), out
            )
            if status == _STATUS_NOT_FOUND:
                raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=("the pivot was removed", ""))
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
            LIB.fm_sheet_get_pivot_name, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def _cache_id(self) -> int:
        h, sheet, pivot_index = self._locate()
        return _read_count(
            LIB.fm_sheet_get_pivot_cache_id, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def get_properties(self) -> PivotProperties:
        """Read the pivot's name and the id of the cache it draws from."""
        from .workbook import PivotProperties

        return PivotProperties(name=self.name, cache_id=self._cache_id())

    def rename(self, name: str) -> None:
        """Rename the pivot table."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(
                LIB.fm_sheet_set_pivot_name(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), name_ptr)
            )

    def get_cache(self) -> PivotCache:
        """The handle of the cache the pivot draws from."""
        return self._sheet._workbook.pivot_cache(self._cache_id())

    # -- Layout ------------------------------------------------------------
    def get_layout(self) -> PivotLayout:
        """Evaluate and project the pivot into concrete grid cells."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            handle = 0
            try:
                _check(
                    LIB.fm_sheet_get_pivot_layout(
                        h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                    )
                )
                handle = LIB.read_u32(out)
                present_p = scope.out_ptr()
                rect_p = scope.struct(S.CELL_RECT)
                _check(LIB.fm_pivot_cells_get_bounds(handle, present_p, rect_p))
                rect = _decode_cell_rect_at(rect_p) if LIB.read_u32(present_p) else None
                cells: List[PivotCell] = []
                n = _read_count(LIB.fm_pivot_cells_get_cell_count, handle)
                # One scratch block for the whole walk.
                cptr = scope.struct(S.PIVOT_CELL)
                for i in range(n):
                    S.zero_struct(LIB, S.PIVOT_CELL, cptr)
                    _check(LIB.fm_pivot_cells_get_cell_at(handle, _uint(i, "idx"), cptr))
                    d = S.PIVOT_CELL.unpack(LIB, cptr)
                    value = Value._from_wasm(cptr + S.PIVOT_CELL_VALUE_OFFSET)
                    cells.append(
                        PivotCell(
                            row=d["row"],
                            col=d["col"],
                            value=value,
                            kind=PivotCellKind(d["kind"]),
                            depth=d["depth"],
                            field_name=LIB.read_cstr(d["field_name"]),
                            number_format=LIB.read_cstr(d["number_format"]),
                        )
                    )
                return PivotLayout(rect=rect, cells=cells)
            finally:
                if handle:
                    LIB.fm_pivot_cells_destroy(handle)

    def set_anchor(self, rect: Union[CellRect, str]) -> None:
        """Set the rectangle the pivot reserves (``<location ref>``); its top-left cell is the anchor.

        ``rect`` is a :class:`CellRect` or an unqualified A1 range.
        """
        rect = self._sheet._rect(rect)
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            _check(LIB.fm_sheet_set_pivot_anchor(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), rp))

    def get_grand_totals(self) -> PivotGrandTotals:
        """Return the row and column flags :meth:`set_grand_totals` writes."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            rows = scope.out_ptr()
            cols = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_pivot_grand_totals(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), rows, cols
                )
            )
            return PivotGrandTotals(rows=LIB.read_i32(rows) != 0, columns=LIB.read_i32(cols) != 0)

    def set_grand_totals(self, *, rows: bool = UNCHANGED, columns: bool = UNCHANGED) -> None:
        """Toggle the row / column grand-total bands; an omitted argument keeps its band.

        Raises:
          TypeError: when a given flag is not a ``bool``.
        """
        if rows is UNCHANGED or columns is UNCHANGED:
            current = self.get_grand_totals()
            rows = current.rows if rows is UNCHANGED else rows
            columns = current.columns if columns is UNCHANGED else columns
        for name, value in (("rows", rows), ("columns", columns)):
            if not isinstance(value, bool):
                raise TypeError(f"formulon: {name} must be a bool, got {type(value).__name__}")
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_grand_totals(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                1 if rows else 0,
                1 if columns else 0,
            )
        )

    def get_style_info(self) -> Optional[PivotStyle]:
        """Return the table style :meth:`set_style_info` writes, or ``None`` without a style element."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_STYLE_INFO)
            _check(
                LIB.fm_sheet_get_pivot_style_info(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr
                )
            )
            d = S.PIVOT_STYLE_INFO.unpack(LIB, ptr)
            if not d["present"]:
                return None
            return PivotStyle(
                name=LIB.read_cstr(d["name"]) or None,
                show_row_headers=d["show_row_headers"] != 0,
                show_col_headers=d["show_col_headers"] != 0,
                show_row_stripes=d["show_row_stripes"] != 0,
                show_col_stripes=d["show_col_stripes"] != 0,
                show_last_column=d["show_last_column"] != 0,
            )

    def set_style_info(
        self,
        *,
        name: Optional[str] = UNCHANGED,
        show_row_headers: bool = UNCHANGED,
        show_col_headers: bool = UNCHANGED,
        show_row_stripes: bool = UNCHANGED,
        show_col_stripes: bool = UNCHANGED,
        show_last_column: bool = UNCHANGED,
    ) -> None:
        """Update the pivot's table style, making the style element present; an omitted argument keeps its value.

        ``name=None`` (or ``""``) writes the element without a style name. A
        pivot without a style element starts from no name and every toggle
        off. New pivots start as ``PivotStyleLight16`` with row / column
        headers and last column on.

        Raises:
          TypeError: when a toggle is not a ``bool`` (``None`` included) or
            ``name`` is neither a ``str`` nor ``None``.
        """
        toggles = {
            "show_row_headers": show_row_headers,
            "show_col_headers": show_col_headers,
            "show_row_stripes": show_row_stripes,
            "show_col_stripes": show_col_stripes,
            "show_last_column": show_last_column,
        }
        for key, value in toggles.items():
            if value is not UNCHANGED:
                _flag(value, key)
        if name is not UNCHANGED and name is not None and not isinstance(name, str):
            raise TypeError(f"formulon: name must be a str or None, got {type(name).__name__}")
        current = self.get_style_info()
        if name is UNCHANGED:
            name = current.name if current is not None else None
        flags = {
            key: (getattr(current, key) if current is not None else False) if value is UNCHANGED else value
            for key, value in toggles.items()
        }
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            name_ptr = scope.nonempty_cstr(name)
            _check(
                LIB.fm_sheet_set_pivot_style_info(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    name_ptr,
                    1 if flags["show_row_headers"] else 0,
                    1 if flags["show_col_headers"] else 0,
                    1 if flags["show_row_stripes"] else 0,
                    1 if flags["show_col_stripes"] else 0,
                    1 if flags["show_last_column"] else 0,
                )
            )

    def get_report_layout(self) -> PivotReportLayout:
        """Return the pivot's compact, tabular, or outline report layout."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_pivot_report_layout(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                )
            )
            return PivotReportLayout(LIB.read_i32(out))

    def set_report_layout(self, layout: PivotReportLayout) -> None:
        """Set the pivot's compact, tabular, or outline report layout."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_report_layout(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _sint(layout, "layout")
            )
        )

    # -- Options and projection --------------------------------------------
    def get_options(self) -> PivotOptions:
        """Read the PivotTable Options of the pivot; see :class:`PivotOptions`."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_OPTIONS)
            _check(
                LIB.fm_sheet_get_pivot_options(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr)
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
                show_empty_col=d["show_empty_col"] != 0,
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

    def set_options(
        self,
        *,
        show_error: bool = UNCHANGED,
        error_caption: Optional[str] = UNCHANGED,
        show_missing: bool = UNCHANGED,
        missing_caption: Optional[str] = UNCHANGED,
        show_headers: bool = UNCHANGED,
        show_drill: bool = UNCHANGED,
        show_data_tips: bool = UNCHANGED,
        grid_drop_zones: bool = UNCHANGED,
        show_empty_row: bool = UNCHANGED,
        show_empty_col: bool = UNCHANGED,
        show_data_drop_down: bool = UNCHANGED,
        print_drill: bool = UNCHANGED,
        field_print_titles: bool = UNCHANGED,
        item_print_titles: bool = UNCHANGED,
        merge_item: bool = UNCHANGED,
        indent: int = UNCHANGED,
        page_over_then_down: bool = UNCHANGED,
        page_wrap: int = UNCHANGED,
        multiple_field_filters: bool = UNCHANGED,
        preserve_formatting: bool = UNCHANGED,
        custom_list_sort: bool = UNCHANGED,
        alt_text: Optional[str] = UNCHANGED,
        alt_text_summary: Optional[str] = UNCHANGED,
        grand_total_caption: Optional[str] = UNCHANGED,
        auto_formatting: bool = UNCHANGED,
        enable_drill: bool = UNCHANGED,
    ) -> None:
        """Update the PivotTable Options; an omitted argument keeps its value.

        Arguments are the :class:`PivotOptions` fields. A caption or alt text
        given as ``None`` (or ``""``) is cleared; a flag or a number cannot be.
        A wrong-typed argument fails without changing the pivot.

        Raises:
          TypeError: when a flag is not a ``bool``, ``indent`` or
            ``page_wrap`` is not an integer, or a string argument is neither a
            ``str`` nor ``None``.
          PreconditionFailedError: when the report would overlap another
            structure.
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
        current = self.get_options()

        def pick(name: str) -> Any:
            return given[name] if name in given else getattr(current, name)

        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_OPTIONS)
            values: Dict[str, object] = {name: 1 if pick(name) else 0 for name in _PIVOT_OPTION_FLAGS}
            values["indent"] = _uint(pick("indent"), "indent")
            values["page_wrap"] = _uint(pick("page_wrap"), "page_wrap")
            S.PIVOT_OPTIONS.pack(LIB, ptr, values)
            for name in _PIVOT_OPTION_TEXTS:
                S.write_str_field(LIB, ptr, S.PIVOT_OPTIONS, name, pick(name), scope)
            _check(
                LIB.fm_sheet_set_pivot_options(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr)
            )

    def is_projection_pending(self) -> bool:
        """Return whether a cache change left the pivot's sheet cells stale; see :meth:`project`."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_is_pivot_projection_pending(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                )
            )
            return LIB.read_i32(out) != 0

    def project(self) -> None:
        """Write the pivot's report into its sheet cells now and clear its pending flag.

        Raises when the report would overlap another structure; the pivot stays pending.
        """
        h, sheet, pivot_index = self._locate()
        _check(LIB.fm_sheet_project_pivot(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

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
        from .sheet import _split_cell_args

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
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(row, "row"),
                    _uint(col, "col"),
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
            LIB.fm_sheet_get_pivot_field_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )
        return [self._field_at(located, i) for i in range(count)]

    @staticmethod
    def _field_at(located: Tuple[int, int, int], field_index: int) -> PivotField:
        """Read the definition of pivot field ``field_index``."""
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FIELD_INFO)
            _check(
                LIB.fm_sheet_get_pivot_field(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
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
                cache_field_index=None if d["cache_field_index"] < 0 else d["cache_field_index"],
                sort_order=PivotSortOrder(d["sort_order"]),
                sort_by_field=LIB.read_cstr(d["sort_by_field"]) or None,
                date_periods=PivotDatePeriod(d["date_periods"]),
                date_calendar=PivotCalendar(d["date_calendar"]),
                group_base_index=None if d["group_base_index"] < 0 else d["group_base_index"],
                default_subtotal=d["default_subtotal"] != 0,
                subtotal_fns=[agg for agg in PivotAggregation if d["subtotal_fns"] & (1 << int(agg))],
                insert_blank_row=d["insert_blank_row"] != 0,
                fill_down_labels=d["fill_down_labels"] != 0,
                show_all=d["show_all"] != 0,
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
                    "subtotal_top": 1 if spec.subtotal_top else 0,
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
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                )
            )
            return LIB.read_u32(out)

    def update_field(
        self,
        field_index: int,
        *,
        custom_name: Optional[str] = UNCHANGED,
        number_format: Optional[str] = UNCHANGED,
        default_subtotal: bool = UNCHANGED,
        insert_blank_row: bool = UNCHANGED,
        fill_down_labels: bool = UNCHANGED,
        show_all: bool = UNCHANGED,
        subtotal_top: bool = UNCHANGED,
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
        ``show_all`` lists items without data. Fields apply in signature order,
        so a failure leaves the fields before it applied.

        Raises:
          TypeError: when a flag is not a ``bool`` (``None`` included) or a
            string argument is neither a ``str`` nor ``None``.
        """
        texts = (
            ("custom_name", custom_name, LIB.fm_sheet_set_pivot_field_custom_name),
            ("number_format", number_format, LIB.fm_sheet_set_pivot_field_number_format),
        )
        flags = (
            ("default_subtotal", default_subtotal, LIB.fm_sheet_set_pivot_field_default_subtotal),
            ("insert_blank_row", insert_blank_row, LIB.fm_sheet_set_pivot_field_insert_blank_row),
            ("fill_down_labels", fill_down_labels, LIB.fm_sheet_set_pivot_field_fill_down_labels),
            ("show_all", show_all, LIB.fm_sheet_set_pivot_field_show_all),
            ("subtotal_top", subtotal_top, LIB.fm_sheet_set_pivot_field_subtotal_top),
        )
        for key, value, _ in texts:
            if value is not UNCHANGED and value is not None and not isinstance(value, str):
                raise TypeError(f"formulon: {key} must be a str or None, got {type(value).__name__}")
        for key, value, _ in flags:
            if value is not UNCHANGED:
                _flag(value, key)
        h, sheet, pivot_index = self._locate()
        lead = (_uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index"))
        for _, value, setter in texts:
            if value is not UNCHANGED:
                with _ArgScope() as scope:
                    _check(setter(h, *lead, scope.cstr(value or "")))
        for key, value, setter in flags:
            if value is not UNCHANGED:
                _check(setter(h, *lead, _flag(value, key)))

    def clear_fields(self) -> None:
        """Drop every field from the pivot, emptying the row and column orders."""
        h, sheet, pivot_index = self._locate()
        _check(LIB.fm_sheet_clear_pivot_fields(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

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
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _sint(_nullable_enum(axis, "axis"), "axis"),
            )
        )

    def set_field_sort(self, field_index: int, order: PivotSortOrder, *, by_field: Optional[str] = None) -> None:
        """Set the sibling order of pivot field ``field_index``.

        ``by_field`` names the data field the items sort by; ``None`` sorts by
        the item labels. ``PivotSortOrder.MANUAL`` keeps items in
        :meth:`add_field_item` order and rejects a ``by_field``.
        """
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            by_ptr = scope.nonempty_cstr(by_field)
            _check(
                LIB.fm_sheet_set_pivot_field_sort(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _sint(order, "order"),
                    by_ptr,
                )
            )

    def set_field_show_detail(self, field_index: int, *, on: bool) -> None:
        """Expand (``on=True``) or collapse (``on=False``) every item of field ``field_index`` and its subtotal row."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_field_show_detail(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                1 if on else 0,
            )
        )

    def set_field_date_group(
        self,
        field_index: int,
        periods: PivotDatePeriod,
        calendar: PivotCalendar = PivotCalendar.GREGORIAN,
        interval_days: int = 1,
        start: float = float("nan"),
        end: float = float("nan"),
    ) -> None:
        """Group the date field ``field_index`` by the periods of ``periods``.

        ``field_index`` is a source field or one of its group fields; any grouping
        it has is replaced. Each period becomes a group field appended to the
        pivot cache, finest first, named and labelled in the workbook's locale;
        this pivot shows them in the field's place, coarsest outermost, and
        another pivot showing the field shows them before it.
        ``PivotCalendar.JAPANESE`` groups the years by era year.
        ``interval_days`` above 1 groups by runs of days and needs ``periods``
        to be ``PivotDatePeriod.DAYS`` alone. ``start`` / ``end`` are date
        serials; NaN takes the field's first value and its last (the midnight
        after it for a date without a time). An empty or unknown mask, a zero
        interval, an invalid bound, a start after the end, or a field holding
        anything but dates raises ``FormulonError``.
        """
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_set_pivot_field_date_group(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _uint(periods, "periods_mask"),
                _sint(calendar, "calendar"),
                _uint(interval_days, "interval_days"),
                float(start),
                float(end),
            )
        )

    def clear_field_date_group(self, field_index: int) -> None:
        """Remove the date grouping of pivot field ``field_index``.

        ``field_index`` is a source field or one of its group fields. The group
        fields leave every pivot on the cache and the field returns to the axis
        position they held; a field that is not grouped is left as it is.

        Raises:
          PreconditionFailedError: when the group fields are not the last
            fields of the cache.
        """
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_field_date_group(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index")
            )
        )

    # -- Field items and subtotals -----------------------------------------
    def list_field_items(self, field_index: int) -> List[PivotItem]:
        """Every item of pivot field ``field_index``, in item order; an item's position is its ``item_index``."""
        h, sheet, pivot_index = self._locate()
        count = _read_count(
            LIB.fm_sheet_get_pivot_field_item_count,
            h,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            _uint(field_index, "field_index"),
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
                LIB.fm_sheet_get_pivot_field_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _uint(item_index, "item_index"),
                    ptr,
                )
            )
            d = S.PIVOT_ITEM_INFO.unpack(LIB, ptr)
            return PivotItem(
                name=LIB.read_cstr(d["name"]) or None,
                cache_index=None if d["cache_index"] < 0 else d["cache_index"],
                visible=d["visible"] != 0,
                caption=LIB.read_cstr(d["caption"]) if d["caption"] else None,
                show_detail=d["show_detail"] != 0,
            )

    def update_item(
        self,
        field_index: int,
        item_index: int,
        *,
        visible: bool = UNCHANGED,
        show_detail: bool = UNCHANGED,
        caption: Optional[str] = UNCHANGED,
    ) -> None:
        """Update item ``item_index`` of pivot field ``field_index``; an omitted argument keeps its value.

        ``show_detail`` expands (``True``) or collapses (``False``) the item.
        ``caption`` is shown instead of the item's value; ``None`` clears it.
        A caption longer than 255 UTF-16 code units, or one matching the label
        another item of the field shows, is rejected. Fields apply in
        signature order, so a failure leaves the fields before it applied.

        Raises:
          TypeError: when a flag is not a ``bool`` (``None`` included) or
            ``caption`` is neither a ``str`` nor ``None``.
        """
        flags = (
            ("visible", visible, LIB.fm_sheet_set_pivot_field_item_visible),
            ("show_detail", show_detail, LIB.fm_sheet_set_pivot_field_item_show_detail),
        )
        for key, value, _ in flags:
            if value is not UNCHANGED:
                _flag(value, key)
        if caption is not UNCHANGED and caption is not None and not isinstance(caption, str):
            raise TypeError(f"formulon: caption must be a str or None, got {type(caption).__name__}")
        h, sheet, pivot_index = self._locate()
        lead = (
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            _uint(field_index, "field_index"),
            _uint(item_index, "item_index"),
        )
        for key, value, setter in flags:
            if value is not UNCHANGED:
                _check(setter(h, *lead, _flag(value, key)))
        if caption is not UNCHANGED:
            with _ArgScope() as scope:
                _check(LIB.fm_sheet_set_pivot_field_item_caption(h, *lead, scope.nonempty_cstr(caption)))

    def add_field_item(self, field_index: int, name: str, *, visible: bool = True) -> int:
        """Append a manual-filter item to pivot field ``field_index``, addressed by its label; return its index.

        The item binds to the first shared item of the cache field that
        renders as ``name`` (unbound when none does), and records are matched
        by comparing their rendered label against ``name``. The blank item
        has no label of its own; build it with :meth:`add_field_shared_item`
        instead.
        """
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_field_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    name_ptr,
                    1 if visible else 0,
                    out,
                )
            )
            return LIB.read_u32(out)

    def add_field_shared_item(self, field_index: int, cache_index: int, *, visible: bool) -> int:
        """Append a manual-filter item addressed by its cache shared-item index; return its index.

        ``cache_index`` selects ``shared_items[cache_index]`` of the bound
        cache field, the same index space as OOXML ``<item x="N">``. The
        label is resolved from that shared item, so a shared item that
        renders to nothing produces the blank item. ``cache_index`` is not
        validated: an index that resolves to nothing yields an item with an
        empty label that filters nothing.
        """
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_field_shared_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _uint(cache_index, "cache_index"),
                    1 if visible else 0,
                    out,
                )
            )
            return LIB.read_u32(out)

    def clear_field_items(self, field_index: int) -> None:
        """Drop every manual-filter item from pivot field ``field_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_field_items(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index")
            )
        )

    def add_field_subtotal_fn(self, field_index: int, agg: PivotAggregation) -> None:
        """Append a subtotal-fn entry to pivot field ``field_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_add_pivot_field_subtotal_fn(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _sint(agg, "agg"),
            )
        )

    def clear_field_subtotal_fns(self, field_index: int) -> None:
        """Drop every subtotal-fn entry from pivot field ``field_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_clear_pivot_field_subtotal_fns(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index")
            )
        )

    # -- Axis orders -------------------------------------------------------
    def _get_field_order(self, fn) -> List[int]:
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            out_ptr = scope.out_ptr()
            out_count = scope.out_ptr()
            _check(fn(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out_ptr, out_count))
            count = LIB.read_u32(out_count)
            if count == 0:
                return []
            raw = LIB.read_bytes(LIB.read_u32(out_ptr), 4 * count)
            return [PIVOT_VALUES_FIELD if v == _PIVOT_VALUES_FIELD_C else v for v in struct.unpack("<%dI" % count, raw)]

    def _set_field_order(self, fn, indices: Sequence[int]) -> None:
        h, sheet, pivot_index = self._locate()
        arr_ptr = 0
        with _ArgScope() as scope:
            if indices:
                arr_ptr = scope.alloc(4 * len(indices))
                buf = b"".join(
                    struct.pack("<I", _PIVOT_VALUES_FIELD_C if i == PIVOT_VALUES_FIELD else _uint(i, "index"))
                    for i in indices
                )
                LIB.write_bytes(arr_ptr, buf)
            _check(
                fn(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
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
            LIB.fm_sheet_get_pivot_data_field_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )
        return [self._data_field_at((h, sheet, pivot_index), i) for i in range(count)]

    @staticmethod
    def _data_field_at(located: Tuple[int, int, int], data_field_index: int) -> PivotDataField:
        """Read data field ``data_field_index`` in the shape :meth:`add_data_field` accepts."""
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_DATA_FIELD_SPEC)
            _check(
                LIB.fm_sheet_get_pivot_data_field(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(data_field_index, "data_field_index"),
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
                show_as_base_field=None if d["show_as_base_field"] < 0 else d["show_as_base_field"],
                show_as_base_item=None if d["show_as_base_item"] < 0 else d["show_as_base_item"],
            )

    @staticmethod
    def _pack_data_field_spec(spec: PivotDataFieldInput, scope: _ArgScope) -> int:
        ptr = scope.struct(S.PIVOT_DATA_FIELD_SPEC)
        S.PIVOT_DATA_FIELD_SPEC.pack(
            LIB,
            ptr,
            {
                "field_index": _uint(spec.field_index, "field_index"),
                "aggregation": _sint(spec.aggregation, "aggregation"),
                "show_as": _sint(spec.show_as, "show_as"),
                "show_as_base_field": _sint(
                    -1 if spec.show_as_base_field is None else spec.show_as_base_field, "show_as_base_field"
                ),
                "show_as_base_item": _sint(
                    -1 if spec.show_as_base_item is None else spec.show_as_base_item, "show_as_base_item"
                ),
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
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
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
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(data_field_index, "data_field_index"),
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
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            _uint(field_index, "field_index"),
            _sint(aggregation, "aggregation"),
            exclude,
        )

    def clear_data_fields(self) -> None:
        """Drop every data-field entry from the pivot."""
        h, sheet, pivot_index = self._locate()
        _check(LIB.fm_sheet_clear_pivot_data_fields(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

    # -- Filters -----------------------------------------------------------
    def list_filters(self) -> List[PivotFilter]:
        """Every listed filter of the pivot, in order; a filter's position is its ``filter_index``.

        The filter list is saved as the definition's ``<filters>`` block. A
        loaded file's caption contains / begins-with, top-N count, value
        greater-than / between and date-between entries appear here, naming
        their field by source name. Its other entries are not listed, still
        filter, survive :meth:`clear_filters` / :meth:`remove_filter`, and are
        saved verbatim.
        """
        h, sheet, pivot_index = self._locate()
        count = _read_count(
            LIB.fm_sheet_get_pivot_filter_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )
        return [self._filter_at((h, sheet, pivot_index), i) for i in range(count)]

    @staticmethod
    def _filter_at(located: Tuple[int, int, int], filter_index: int) -> PivotFilter:
        """Read the filter at ``filter_index`` without mutating it."""
        h, sheet, pivot_index = located
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FILTER_SPEC)
            _check(
                LIB.fm_sheet_get_pivot_filter(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(filter_index, "filter_index"),
                    ptr,
                )
            )
            d = S.PIVOT_FILTER_SPEC.unpack(LIB, ptr)
            return PivotFilter(
                axis=PivotAxis(d["axis"]),
                field_name=LIB.read_cstr(d["field_name"]),
                type=PivotFilterType(d["type"]),
                value_kind=_enum_or_none(PivotFilterValueKind, d["value_kind"]),
                value_int=d["value_int"],
                value_double=d["value_double"],
                value_text=LIB.read_cstr(d["value_text"]),
                value_high_kind=_enum_or_none(PivotFilterValueKind, d["value_high_kind"]),
                value_high_int=d["value_high_int"],
                value_high_double=d["value_high_double"],
                data_field_index=d["data_field_index"],
                value_high_text=LIB.read_cstr(d["value_high_text"]),
            )

    def add_filter(self, spec: PivotFilterInput) -> int:
        """Append a filter to the pivot and return its index; it is saved with it. See :meth:`list_filters`."""
        h, sheet, pivot_index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FILTER_SPEC)
            S.PIVOT_FILTER_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": _sint(spec.axis, "axis"),
                    "type": _sint(spec.type, "type"),
                    "data_field_index": _uint(spec.data_field_index, "data_field_index"),
                    "value_kind": _sint(_nullable_enum(spec.value_kind, "value_kind"), "value_kind"),
                    "value_int": _sint(spec.value_int, "value_int"),
                    "value_double": float(spec.value_double),
                    "value_high_kind": _sint(
                        _nullable_enum(spec.value_high_kind, "value_high_kind"), "value_high_kind"
                    ),
                    "value_high_int": _sint(spec.value_high_int, "value_high_int"),
                    "value_high_double": float(spec.value_high_double),
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "field_name", spec.field_name, scope)
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "value_text", spec.value_text, scope)
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "value_high_text", spec.value_high_text, scope)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_filter(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                )
            )
            return LIB.read_u32(out)

    def remove_filter(self, filter_index: int) -> None:
        """Remove the filter at ``filter_index``."""
        h, sheet, pivot_index = self._locate()
        _check(
            LIB.fm_sheet_remove_pivot_filter(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(filter_index, "filter_index")
            )
        )

    def clear_filters(self) -> None:
        """Drop every listed filter from the pivot."""
        h, sheet, pivot_index = self._locate()
        _check(LIB.fm_sheet_clear_pivot_filters(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))


def _show_details_coordinate_form(self, row: int, col: int) -> Sheet: ...


Pivot.show_details.__signature__ = inspect.signature(_show_details_coordinate_form)  # type: ignore[attr-defined]
del _show_details_coordinate_form

_install_workbook_locks(Pivot)
