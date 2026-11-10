"""The :class:`Table` handle: every operation that targets one worksheet table of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`Table` per table and keeps it
pointing at the same table as the table list changes; see :class:`Table`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Tuple, Union

from . import _structs as S
from ._c import LIB, FormulonError, _argument_message, _NoCopy, _require_bool, _sint, _uint
from ._sheet import _normalize_auto_filter, _read_auto_filter
from ._status import StatusCode
from ._workbook import (
    _PATCH_CLEAR,
    _PATCH_SET,
    UNCHANGED,
    AutoFilter,
    AutoFilterEval,
    AutoFilterInput,
    CellRect,
    SheetRect,
    TableColumn,
    TableColumnTotalsInput,
    TableStyleOptions,
    TableTotalsRowFunction,
    Unchanged,
    _ArgScope,
    _check,
    _decode_cell_rect_at,
    _pack_cell_rect_at,
)

if TYPE_CHECKING:
    from ._sheet import Sheet
    from ._workbook import TableProperties, Workbook

__all__ = ["Table", "TableTotalsRowFunction"]

_STYLE_TOGGLES = ("show_first_column", "show_last_column", "show_row_stripes", "show_column_stripes")


class Table(_NoCopy):
    """One worksheet table of a :class:`Sheet`.

    Obtain a handle from :meth:`Sheet.table`, :meth:`Sheet.list_tables`,
    :meth:`Sheet.add_table` or :meth:`Workbook.table`. A workbook hands out
    one handle per table, so every lookup of the same table returns the same
    object. The handle follows its table when tables are added or removed,
    rows or columns are edited, or the table is renamed. Once the table is
    removed (directly, or with its sheet), every operation except reading
    :attr:`workbook` raises :class:`NotFoundError`. Remove a table with
    :meth:`Sheet.remove_table`.
    """

    __slots__ = ("_workbook", "_runtime_id")

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain a Table from Sheet.table(), Sheet.add_table() or Workbook.table()")

    @classmethod
    def _new(cls, workbook: Workbook, runtime_id: int) -> "Table":
        handle = object.__new__(cls)
        handle._workbook = workbook
        handle._runtime_id = runtime_id
        return handle

    @property
    def workbook(self) -> Workbook:
        """The workbook this table belongs to."""
        return self._workbook

    @property
    def sheet(self) -> Sheet:
        """The sheet the table lives on.

        Raises:
          NotFoundError: when the table was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._workbook.sheet(self._locate()[1])

    @property
    def index(self) -> int:
        """The table's current index on its sheet (its position in :meth:`Sheet.list_tables`).

        Raises:
          NotFoundError: when the table was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._locate()[2]

    @property
    def name(self) -> str:
        """The table's current name: the name formulas and :meth:`Workbook.table` use.

        Raises:
          NotFoundError: when the table was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        h, sheet, index = self._locate()
        return self._name(h, sheet, index)

    def __repr__(self) -> str:
        try:
            h, sheet, index = self._locate()
        except FormulonError:
            return "Table(removed)"
        return f"Table(sheet={sheet}, index={index}, name={self._name(h, sheet, index)!r})"

    def _locate(self) -> Tuple[int, int, int]:
        """The workbook handle, the sheet index and the table's current index, resolved from its runtime id."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            sheet_ptr = scope.out_ptr()
            index_ptr = scope.out_ptr()
            status = LIB.fm_workbook_resolve_table_runtime_id(
                h, _uint(self._runtime_id, "runtime_id"), sheet_ptr, index_ptr
            )
            if status == StatusCode.NOT_FOUND:
                raise FormulonError(StatusCode.NOT_FOUND, "the table was removed", "")
            _check(status)
            return h, LIB.read_u32(sheet_ptr), LIB.read_u32(index_ptr)

    @staticmethod
    def _name(h: int, sheet: int, index: int) -> str:
        """The name of the table at ``(sheet, index)``."""
        with _ArgScope() as scope:
            info_ptr = scope.struct(S.TABLE_INFO)
            _check(
                LIB.fm_sheet_get_table_at(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position"), info_ptr
                )
            )
            return LIB.read_cstr(S.TABLE_INFO.unpack(LIB, info_ptr)["name"])

    # -- Properties --------------------------------------------------------
    def get_properties(self) -> TableProperties:
        """Read the table's name, range, header / totals flags and columns."""
        from ._workbook import TableProperties

        h, sheet, index = self._locate()
        sheet_arg, index_arg = _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position")
        with _ArgScope() as scope:
            info_ptr = scope.struct(S.TABLE_INFO)
            column_ptr = scope.struct(S.TABLE_COLUMN_INFO)
            _check(LIB.fm_sheet_get_table_at(h, sheet_arg, index_arg, info_ptr))
            info = S.TABLE_INFO.unpack(LIB, info_ptr)
            name = LIB.read_cstr(info["name"])
            rect = _decode_cell_rect_at(info_ptr + S.TABLE_INFO.offsets["range"][1])
            columns = []
            for c in range(info["column_count"]):
                S.zero_struct(LIB, S.TABLE_COLUMN_INFO, column_ptr)
                _check(
                    LIB.fm_sheet_get_table_column_at(
                        h, sheet_arg, index_arg, _uint(c, "column", role="position"), column_ptr
                    )
                )
                col = S.TABLE_COLUMN_INFO.unpack(LIB, column_ptr)
                columns.append(
                    TableColumn(
                        name=LIB.read_cstr(col["name"]),
                        totals_row_function=col["totals_row_function"],
                        totals_row_label=LIB.read_optional_cstr(col["totals_row_label"]),
                        totals_row_formula=LIB.read_optional_cstr(col["totals_row_formula"]),
                    )
                )
            return TableProperties(
                name=name,
                range=SheetRect(sheet=sheet, rect=rect),
                header_row=info["header_row_count"] != 0,
                totals_row=info["totals_row_count"] != 0,
                columns=columns,
            )

    def update(
        self,
        *,
        range: Union[SheetRect, CellRect, str, Unchanged] = UNCHANGED,
        header_row: Union[bool, Unchanged] = UNCHANGED,
        totals_row: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the table's range and header / totals rows; an omitted argument keeps its setting.

        ``range`` is a :class:`CellRect`, an A1 range (``"A1:C5"``,
        optionally sheet-qualified) or a :class:`SheetRect`; it must lie on
        the table's sheet and keep the width the table's column list already
        describes. The update is all-or-nothing. Rename with :meth:`rename` and
        change the style with :meth:`update_style_options`.

        Turning ``totals_row`` on without ``range`` inserts one row below the table (the table's columns
        shift down, as :meth:`Sheet.insert_cells` with ``DOWN``) and writes each column's stored totals
        label and function there; the data rows are kept. If something straddles the table's columns
        below it (a merge, another table, a PivotTable, an array formula, a data table), nothing shifts
        and the row directly below is used instead when it is free (no value, formula, spill output, merge
        or table cell; format-only cells count as free), else the call is refused. Turning it off
        first checks that no formula or defined name references the totals row (a range meeting it, or a
        ``[#Totals]`` / ``[#All]`` structured reference) and refuses if one does; it then shrinks the table
        by one row and shifts the cells below up. If something straddles the table's columns below, the
        old totals row is only emptied and nothing shifts. The per-column labels and functions are kept
        for the next time it is turned on. With ``range`` in the same call ``range`` is the final extent
        and only the flag changes: nothing shifts (a read record feeds back).

        Raises:
          TypeError: when a flag is not a ``bool`` or ``range`` has the wrong
            type (``None`` included).
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``range``
            names another sheet or is not an A1 range.
          PreconditionFailedError: when turning the totals row on finds the row below
            occupied while something straddles the table's columns, or the table on the last grid row; or
            when turning it off finds a formula or defined name referencing the totals row.
        """
        for key, flag in (("header_row", header_row), ("totals_row", totals_row)):
            if flag is not UNCHANGED:
                _require_bool(flag, key)
        if range is not UNCHANGED and not isinstance(range, (SheetRect, CellRect, str)):
            raise TypeError(_argument_message("range", "must be a string, a CellRect or a SheetRect"))
        h, sheet, index = self._locate()
        values = {}
        if range is not UNCHANGED:
            values["range_op"] = _PATCH_SET
        for key, flag in (("header_row", header_row), ("totals_row", totals_row)):
            if flag is not UNCHANGED:
                values[key + "_op"] = _PATCH_SET
                values[key] = int(flag)
        with _ArgScope() as scope:
            ptr = scope.struct(S.TABLE_PATCH)
            if range is not UNCHANGED:
                _pack_cell_rect_at(ptr + S.TABLE_PATCH.offsets["range"][1], self._range_rect(range, sheet))
            raw = bytearray(LIB.read_bytes(ptr, S.TABLE_PATCH.size))
            S.TABLE_PATCH.pack_into(raw, 0, values)
            LIB.write_bytes(ptr, bytes(raw))
            _check(
                LIB.fm_sheet_update_table(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position"), ptr
                )
            )

    def _range_rect(self, range: Union[SheetRect, CellRect, str], sheet: int) -> CellRect:
        """The rectangle ``range`` names, which must lie on sheet index ``sheet``."""
        if isinstance(range, CellRect):
            return range
        if isinstance(range, SheetRect):
            target, rect = range.sheet, range.rect
        else:
            ref = self._workbook.parse_reference(range)
            if ref is None:
                raise FormulonError(StatusCode.INVALID_ARGUMENT, f"{range!r} is not an A1 range", "")
            rect = ref.range
            target = sheet if ref.sheet_name is None else self._workbook.sheet(ref.sheet_name).index
        if target != sheet:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                "the range is on a different sheet than the table",
                "",
            )
        return rect

    def rename(self, name: str) -> None:
        """Rename the table.

        A name another table already has (compared case-insensitively) raises
        ``AlreadyExistsError``, an empty or malformed one
        ``InvalidArgumentError``. Structured references to the old name are
        rewritten in cell, defined-name, conditional-format, validation and
        chart formulas (a bare table name used as a reference is not).
        """
        h, sheet, index = self._locate()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_rename_table(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(index, "table_index", role="position"),
                    scope.cstr(name),
                )
            )

    # -- Style -------------------------------------------------------------
    def get_style_options(self) -> Optional[TableStyleOptions]:
        """Return the table's style name and toggles, or ``None`` without a style element."""
        h, sheet, index = self._locate()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.TABLE_STYLE_OPTIONS)
            _check(
                LIB.fm_sheet_get_table_style_options(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(index, "table_index", role="position"),
                    present,
                    ptr,
                )
            )
            if LIB.read_i32(present) == 0:
                return None
            d = S.TABLE_STYLE_OPTIONS.unpack(LIB, ptr)
            return TableStyleOptions(
                style_name=LIB.read_optional_cstr(d["style_name"]),
                show_first_column=d["show_first_column"] != 0,
                show_last_column=d["show_last_column"] != 0,
                show_row_stripes=d["show_row_stripes"] != 0,
                show_column_stripes=d["show_column_stripes"] != 0,
            )

    def update_style_options(
        self,
        *,
        style_name: Union[str, None, Unchanged] = UNCHANGED,
        show_first_column: Union[bool, Unchanged] = UNCHANGED,
        show_last_column: Union[bool, Unchanged] = UNCHANGED,
        show_row_stripes: Union[bool, Unchanged] = UNCHANGED,
        show_column_stripes: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the table's style name and toggles; an omitted argument keeps its setting.

        ``style_name=None`` removes the style name, ``""`` raises ``InvalidArgumentError``; a new ``style_name`` keeps
        the four toggles. The update is all-or-nothing.

        Raises:
          TypeError: when a toggle is not a ``bool`` (``None`` included) or
            ``style_name`` is neither a ``str`` nor ``None``.
        """
        toggles = {
            "show_first_column": show_first_column,
            "show_last_column": show_last_column,
            "show_row_stripes": show_row_stripes,
            "show_column_stripes": show_column_stripes,
        }
        for key, value in toggles.items():
            if value is not UNCHANGED:
                _require_bool(value, key)
        if style_name is not UNCHANGED and style_name is not None and not isinstance(style_name, str):
            raise TypeError(_argument_message("style_name", "must be a string"))
        h, sheet, index = self._locate()
        with _ArgScope() as scope:
            ptr = scope.struct(S.TABLE_STYLE_OPTIONS_PATCH)
            values = {}
            if style_name is None:
                values["style_name_op"] = _PATCH_CLEAR
            elif style_name is not UNCHANGED:
                values["style_name_op"] = _PATCH_SET
                values["style_name"] = scope.cstr(style_name)
            for key in _STYLE_TOGGLES:
                if toggles[key] is not UNCHANGED:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = int(_require_bool(toggles[key], key))
            S.TABLE_STYLE_OPTIONS_PATCH.pack(LIB, ptr, values)
            _check(
                LIB.fm_sheet_update_table_style_options(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position"), ptr
                )
            )

    # -- Columns -----------------------------------------------------------
    def rename_column(self, column: int, name: str) -> str:
        """Rename column ``column`` (0-based) and return the final name.

        As typing over the header cell does: an empty ``name`` takes the
        profile's ``Column<N>`` (``列<N>``) with the lowest unused N, a name
        another column already has (ignoring case under Unicode simple case
        folding, as sheet names do) gains the
        lowest unused N >= 2, and anything else is kept verbatim. Structured
        references to the column through this table are rewritten in cell,
        totals-row and defined-name formulas and chart series. Writing a
        value to a header cell (``set_text``, ``set_number``, ...) renames the
        same way.
        """
        h, sheet, index = self._locate()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_rename_table_column(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(index, "table_index", role="position"),
                    _uint(column, "column", role="position"),
                    name_ptr,
                    out,
                )
            )
            return LIB.read_cstr(LIB.read_u32(out))

    # -- Inserting rows and columns -----------------------------------------
    def insert_rows(self, first: Optional[int] = None, last: Optional[int] = None) -> None:
        """Insert data rows into the table, as ListRows.Add does.

        The cells of the table's columns from the first new row downwards move
        down, the cells of other columns stay, and the table grows by the
        span. Every calculated column's formula is written into the new rows
        with its relative references moved along; other columns stay blank.
        Structured references and the table's AutoFilter follow the new
        extent.

        Called with neither argument it appends one row, above the totals row
        when the table has one. With ``first`` alone it inserts one row at
        ``first``.

        Args:
          first: the 0-based data-row position of the first new row; the
            data-row count appends. Defaults to the data-row count.
          last: the 0-based data-row position of the last new row, not below
            ``first``; the span is inclusive. Defaults to ``first``. May be
            given only together with ``first``.

        Raises:
          TypeError: when ``first`` or ``last`` is not an integer.
          InvalidArgumentError: when ``first`` is past the data-row count,
            ``last`` is below ``first`` or given without ``first``, either is
            negative, or the span reaches past the last sheet row.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when a
            PivotTable lies in the shift path, a multi-cell array formula or
            data table is cut by the band, or a value, formula or the table
            itself would be pushed past the last sheet row; the workbook is
            unchanged.
          NotFoundError: when the table was removed.
        """
        h, sheet, index = self._locate()
        data_rows, _ = self._extent(h, sheet, index)
        first, last = self._insert_span(first, last, data_rows)
        _check(
            LIB.fm_sheet_insert_table_rows(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(index, "table_index", role="position"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    def insert_columns(self, first: Optional[int] = None, last: Optional[int] = None) -> None:
        """Insert columns into the table, as ListColumns.Add does.

        The cells of the table's rows (header through totals row) from the
        first new column rightwards move right, the cells of other rows stay,
        and the table grows by the span. Each new column is named from its
        left neighbour (``Q2`` gives ``Q3``; a name without trailing digits,
        or no left neighbour, takes the profile's ``Column<N>`` / ``列<N>``
        with the lowest free N) and the name is written to the header cell
        when the table has a header row; the data and totals cells stay blank.

        Called with neither argument it appends one column after the last.
        With ``first`` alone it inserts one column at ``first``.

        Args:
          first: the 0-based table-column position of the first new column;
            the column count appends. Defaults to the column count.
          last: the 0-based table-column position of the last new column, not
            below ``first``; the span is inclusive. Defaults to ``first``. May
            be given only together with ``first``.

        Raises:
          InvalidArgumentError: as :meth:`insert_rows`, with the last sheet
            column as the bound.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` as
            :meth:`insert_rows`.
          NotFoundError: when the table was removed.
        """
        h, sheet, index = self._locate()
        _, column_count = self._extent(h, sheet, index)
        first, last = self._insert_span(first, last, column_count)
        _check(
            LIB.fm_sheet_insert_table_columns(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(index, "table_index", role="position"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @staticmethod
    def _extent(h: int, sheet: int, index: int) -> Tuple[int, int]:
        """The table's data-row count and column count."""
        with _ArgScope() as scope:
            info_ptr = scope.struct(S.TABLE_INFO)
            _check(
                LIB.fm_sheet_get_table_at(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position"), info_ptr
                )
            )
            info = S.TABLE_INFO.unpack(LIB, info_ptr)
            rect = _decode_cell_rect_at(info_ptr + S.TABLE_INFO.offsets["range"][1])
            data_rows = rect.last_row - rect.first_row + 1 - info["header_row_count"] - info["totals_row_count"]
            return data_rows, info["column_count"]

    @staticmethod
    def _insert_span(first: Optional[int], last: Optional[int], count: int) -> Tuple[int, int]:
        """The ``(first, last)`` of an insert: both omitted appends one at ``count``, ``first`` alone is one at it."""
        if first is None:
            if last is not None:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT, _argument_message("last", "needs first"), f"last={last}"
                )
            return count, count
        return first, first if last is None else last

    def set_column_totals(self, column: int, totals: Union[TableColumnTotalsInput, TableColumn]) -> None:
        """Set column ``column``'s totals-row function, label and formula from a record.

        ``totals.totals_row_function`` is a :class:`TableTotalsRowFunction`;
        ``totals_row_formula`` is required for ``CUSTOM``. A ``None``
        label or formula means none; ``""`` raises ``InvalidArgumentError``. A :class:`TableColumn` read back is
        accepted in place of a :class:`TableColumnTotalsInput`; its ``name`` is
        ignored. While the totals row is shown its cell follows:
        ``SUBTOTAL(109,Table[Column])`` and the other 1xx codes for a built-in
        function, the formula for ``CUSTOM``, else the label as text, else
        nothing.

        Raises:
          TypeError: when ``totals`` is not one of the two record types.
        """
        if not isinstance(totals, (TableColumnTotalsInput, TableColumn)):
            raise TypeError(_argument_message("totals", "must be a TableColumnTotalsInput"))
        h, sheet, index = self._locate()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_table_column_totals(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(index, "table_index", role="position"),
                    _uint(column, "column", role="position"),
                    _sint(totals.totals_row_function, "totals_row_function"),
                    0 if totals.totals_row_label is None else scope.cstr(totals.totals_row_label),
                    0 if totals.totals_row_formula is None else scope.cstr(totals.totals_row_formula),
                )
            )

    # -- AutoFilter --------------------------------------------------------
    def get_auto_filter(self) -> Optional[AutoFilter]:
        """Return the table's AutoFilter, or ``None`` when it has none."""
        _, sheet, index = self._locate()
        return _read_auto_filter(self._workbook, LIB.fm_sheet_get_table_auto_filter, sheet, index)

    def set_auto_filter(self, filter: Optional[Union[AutoFilterInput, AutoFilter]]) -> None:
        """Replace the table's AutoFilter, or remove it with ``None`` (no-op without one). Row visibility is unchanged.

        ``range`` of the filter, of its sort state and of its sort conditions takes an A1 range too.
        """
        _, sheet, index = self._locate()
        if filter is not None:
            filter = _normalize_auto_filter(filter, self._workbook)
        self._workbook._write_auto_filter(LIB.fm_sheet_set_table_auto_filter, filter, sheet, index)

    def apply_auto_filter(self) -> None:
        """Hide the body rows the table's AutoFilter excludes and show the rest."""
        h, sheet, index = self._locate()
        _check(
            LIB.fm_sheet_apply_table_auto_filter(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position")
            )
        )

    def clear_auto_filter_criteria(self) -> None:
        """Drop every column criterion of the table's AutoFilter and show all its body rows.

        The AutoFilter itself stays (remove it with ``set_auto_filter(None)``);
        a filter without criteria is left as it is.
        """
        h, sheet, index = self._locate()
        _check(
            LIB.fm_sheet_clear_table_auto_filter_criteria(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(index, "table_index", role="position")
            )
        )

    def evaluate_auto_filter(self) -> AutoFilterEval:
        """Evaluate the table's AutoFilter without changing the sheet, as :meth:`Sheet.evaluate_auto_filter` does."""
        _, sheet, index = self._locate()
        return AutoFilterEval(
            *self._workbook._evaluate_auto_filter(LIB.fm_sheet_evaluate_table_auto_filter, sheet, index)
        )
