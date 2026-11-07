"""The :class:`Table` handle: every operation that targets one worksheet table of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`Table` per table and keeps it
pointing at the same table as the table list changes; see :class:`Table`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional, Sequence, Tuple

from . import _structs as S
from ._c import LIB, FormulonError, _sint, _uint
from ._status import StatusCode
from .workbook import (
    UNCHANGED,
    AutoFilter,
    AutoFilterEval,
    CellRect,
    SheetRect,
    TableColumn,
    _ArgScope,
    _check,
    _decode_cell_rect_at,
    _pack_cell_rect_at,
    _tristate,
)

if TYPE_CHECKING:
    from .workbook import TableProperties, Workbook

__all__ = ["Table"]


class Table:
    """One worksheet table of a :class:`Workbook`.

    Obtain a handle from :meth:`Workbook.table`, :meth:`Workbook.list_tables`
    or :meth:`Workbook.add_table`. A workbook hands out one handle per table,
    so every lookup of the same table returns the same object. The handle
    follows its table when tables are added or removed, rows or columns are
    edited, or the table is renamed. Once the table is removed (directly, or
    with its sheet), every operation except reading :attr:`workbook` raises
    :class:`NotFoundError`. Remove a table with :meth:`Workbook.remove_table`.
    """

    __slots__ = ("_workbook", "_runtime_id")

    def __init__(self, workbook: Workbook, runtime_id: int) -> None:
        self._workbook = workbook
        self._runtime_id = runtime_id

    @property
    def workbook(self) -> Workbook:
        """The workbook this table belongs to."""
        return self._workbook

    @property
    def index(self) -> int:
        """The table's current workbook-wide index (its position in :meth:`Workbook.list_tables`).

        Raises:
          NotFoundError: when the table was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._index()

    @property
    def name(self) -> str:
        """The table's current name.

        Raises:
          NotFoundError: when the table was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._names(self._index())[0]

    @property
    def display_name(self) -> str:
        """The table's current display name, the name formulas use.

        Raises:
          NotFoundError: when the table was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._names(self._index())[1]

    def __repr__(self) -> str:
        try:
            index = self._index()
        except FormulonError:
            return "Table(removed)"
        return f"Table(index={index}, display_name={self._names(index)[1]!r})"

    def _index(self) -> int:
        """The table's current index, resolved from its runtime id."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            status = LIB.fm_workbook_resolve_table_runtime_id(h, _uint(self._runtime_id, "runtime_id"), out)
            if status == StatusCode.NOT_FOUND:
                raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=("the table was removed", ""))
            _check(status)
            return LIB.read_u32(out)

    def _names(self, index: int) -> Tuple[str, str]:
        """The name and display name of the table at ``index``."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.out_ptr()
            display_ptr = scope.out_ptr()
            rect_ptr = scope.struct(S.CELL_RECT)
            sheet_ptr = scope.out_ptr()
            _check(
                LIB.fm_workbook_get_table(h, _uint(index, "table_index"), name_ptr, display_ptr, rect_ptr, sheet_ptr)
            )
            return LIB.read_cstr(LIB.read_u32(name_ptr)), LIB.read_cstr(LIB.read_u32(display_ptr))

    # -- Properties --------------------------------------------------------
    def get_properties(self) -> TableProperties:
        """Read the table's name, range, style, toggles and columns."""
        return self._properties(self._index())

    def _properties(self, index: int) -> TableProperties:
        from .workbook import TableProperties

        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.out_ptr()
            display_ptr = scope.out_ptr()
            rect_ptr = scope.struct(S.CELL_RECT)
            sheet_ptr = scope.out_ptr()
            info_ptr = scope.struct(S.TABLE_INFO)
            column_ptr = scope.struct(S.TABLE_COLUMN_INFO)
            _check(
                LIB.fm_workbook_get_table(h, _uint(index, "table_index"), name_ptr, display_ptr, rect_ptr, sheet_ptr)
            )
            name = LIB.read_cstr(LIB.read_u32(name_ptr))
            display = LIB.read_cstr(LIB.read_u32(display_ptr))
            rect = _decode_cell_rect_at(rect_ptr)
            sheet = LIB.read_u32(sheet_ptr)
            S.zero_struct(LIB, S.TABLE_INFO, info_ptr)
            _check(LIB.fm_workbook_get_table_info(h, _uint(index, "table_index"), info_ptr))
            info = S.TABLE_INFO.unpack(LIB, info_ptr)
            columns = []
            for c in range(info["column_count"]):
                S.zero_struct(LIB, S.TABLE_COLUMN_INFO, column_ptr)
                _check(LIB.fm_workbook_get_table_column(h, _uint(index, "table_index"), _uint(c, "column"), column_ptr))
                col = S.TABLE_COLUMN_INFO.unpack(LIB, column_ptr)
                columns.append(
                    TableColumn(
                        name=LIB.read_cstr(col["name"]),
                        totals_row_function=LIB.read_cstr(col["totals_row_function"]),
                        totals_row_label=LIB.read_cstr(col["totals_row_label"]),
                        totals_row_formula=LIB.read_cstr(col["totals_row_formula"]),
                    )
                )
            return TableProperties(
                name=name,
                display_name=display,
                range=SheetRect(sheet=sheet, rect=rect),
                style_name=LIB.read_cstr(info["style_name"]) or None,
                show_first_column=info["show_first_column"] != 0,
                show_last_column=info["show_last_column"] != 0,
                show_row_stripes=info["show_row_stripes"] != 0,
                show_column_stripes=info["show_column_stripes"] != 0,
                header_row_count=info["header_row_count"],
                totals_row_count=info["totals_row_count"],
                columns=columns,
            )

    def update(
        self,
        *,
        rect: CellRect = UNCHANGED,
        style_name: Optional[str] = UNCHANGED,
        header_row: bool = UNCHANGED,
        totals_row: bool = UNCHANGED,
        name: str = UNCHANGED,
        display_name: str = UNCHANGED,
        show_first_column: bool = UNCHANGED,
        show_last_column: bool = UNCHANGED,
        show_row_stripes: bool = UNCHANGED,
        show_column_stripes: bool = UNCHANGED,
    ) -> None:
        """Update the table's range, flags, names and style; an omitted argument keeps its setting.

        A new ``rect`` must keep the width the table's column list already
        describes. A new ``style_name`` keeps the four toggles, and ``None``
        (or ``""``) removes ``<tableStyleInfo>`` -- unless a ``show_*`` toggle
        is also given, which writes the element back without a style name.
        Turning the totals row on writes each column's totals cell. Names
        must be non-empty and unique case-insensitively across tables;
        ``display_name`` follows a new ``name`` unless given, and structured
        references to the old name are rewritten in cell, defined-name,
        conditional-format, validation and chart formulas (a bare table name
        used as a reference is not). The update is all-or-nothing.

        Raises:
          TypeError: when a flag is not a ``bool``.
        """
        index = self._index()
        flags = (header_row, totals_row, show_first_column, show_last_column, show_row_stripes, show_column_stripes)
        for flag in flags:
            if flag is not UNCHANGED and not isinstance(flag, bool):
                raise TypeError(f"formulon: a table flag must be a bool, got {type(flag).__name__}")
        current = self._properties(index)
        renaming = name is not UNCHANGED or display_name is not UNCHANGED
        if renaming:
            next_name = current.name if name is UNCHANGED else name
            if display_name is not UNCHANGED:
                next_display = display_name
            else:
                next_display = current.display_name if name is UNCHANGED else name
            self._rename(index, next_name, next_display)
        try:
            self._update(
                index,
                current.range.rect if rect is UNCHANGED else rect,
                None if style_name is UNCHANGED else (style_name or ""),
                header_row=None if header_row is UNCHANGED else header_row,
                totals_row=None if totals_row is UNCHANGED else totals_row,
            )
            toggles = tuple(None if t is UNCHANGED else t for t in flags[2:])
            if any(t is not None for t in toggles):
                self._apply_style_toggles(index, toggles)
        except BaseException:
            if renaming:
                self._rename(index, current.name, current.display_name)
            raise

    def _rename(self, index: int, name: str, display_name: str) -> None:
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(
                LIB.fm_workbook_rename_table(h, _uint(index, "table_index"), scope.cstr(name), scope.cstr(display_name))
            )

    def _update(
        self,
        index: int,
        rect: CellRect,
        style_name: Optional[str],
        *,
        header_row: Optional[bool],
        totals_row: Optional[bool],
    ) -> None:
        h = self._workbook._require()
        with _ArgScope() as scope:
            rect_ptr = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rect_ptr, rect)
            # NULL keeps the style; an empty name removes the element.
            style_ptr = scope.opt_cstr(style_name)
            _check(
                LIB.fm_workbook_update_table(
                    h,
                    _uint(index, "table_index"),
                    rect_ptr,
                    style_ptr,
                    _sint(_tristate(header_row), "header_row"),
                    _sint(_tristate(totals_row), "totals_row"),
                )
            )

    def _apply_style_toggles(self, index: int, toggles: Sequence[Optional[bool]]) -> None:
        """Set the supplied ``show_*`` toggles over the table's current ones."""
        current = self._properties(index)
        h = self._workbook._require()
        with _ArgScope() as scope:
            info_ptr = scope.struct(S.TABLE_INFO)
            _check(LIB.fm_workbook_get_table_info(h, _uint(index, "table_index"), info_ptr))
            style_present = S.TABLE_INFO.unpack(LIB, info_ptr)["style_present"] != 0
        existing = (
            current.show_first_column,
            current.show_last_column,
            current.show_row_stripes if style_present else True,
            current.show_column_stripes,
        )
        first, last, rows, columns = (existing[i] if t is None else t for i, t in enumerate(toggles))
        with _ArgScope() as scope:
            # Writes the style element, keeping its name (none when absent).
            _check(
                LIB.fm_workbook_set_table_style_info(
                    h,
                    _uint(index, "table_index"),
                    scope.nonempty_cstr(current.style_name),
                    1 if first else 0,
                    1 if last else 0,
                    1 if rows else 0,
                    1 if columns else 0,
                )
            )

    # -- Columns -----------------------------------------------------------
    def rename_column(self, column: int, name: str) -> str:
        """Rename column ``column`` (0-based) and return the final name.

        As typing over the header cell does: an empty ``name`` takes the
        profile's ``Column<N>`` (``列<N>``) with the lowest unused N, a name
        another column already has (ASCII case-insensitively) gains the
        lowest unused N >= 2, and anything else is kept verbatim. Structured
        references to the column through this table are rewritten in cell,
        totals-row and defined-name formulas and chart series. Writing a
        value to a header cell (``set_text``, ``set_number``, ...) renames the
        same way.
        """
        index = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(
                LIB.fm_workbook_rename_table_column(
                    h, _uint(index, "table_index"), _uint(column, "column"), name_ptr, out
                )
            )
            return LIB.read_cstr(LIB.read_u32(out))

    def set_column_totals(
        self,
        column: int,
        *,
        function: Optional[str] = None,
        label: Optional[str] = None,
        formula: Optional[str] = None,
    ) -> None:
        """Set column ``column``'s totals-row function and label.

        ``function`` is an ``ST_TotalsRowFunction`` token (``None`` means
        ``"none"``); ``formula`` is required for ``"custom"``. While the
        totals row is shown its cell follows: ``SUBTOTAL(109,Table[Column])``
        and the other 1xx codes for a built-in function, the formula for
        ``"custom"``, else the label as text, else nothing.
        """
        index = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(
                LIB.fm_workbook_set_table_column_totals(
                    h,
                    _uint(index, "table_index"),
                    _uint(column, "column"),
                    scope.nonempty_cstr(function),
                    scope.nonempty_cstr(label),
                    scope.nonempty_cstr(formula),
                )
            )

    # -- AutoFilter --------------------------------------------------------
    def get_auto_filter(self) -> Optional[AutoFilter]:
        """Return the table's AutoFilter, or ``None`` when it has none."""
        return self._workbook._read_auto_filter(LIB.fm_workbook_get_table_auto_filter, self._index(), "table_index")

    def set_auto_filter(self, auto_filter: AutoFilter) -> None:
        """Replace the table's AutoFilter. Row visibility is unchanged."""
        self._workbook._write_auto_filter(
            LIB.fm_workbook_set_table_auto_filter, self._index(), auto_filter, "table_index"
        )

    def remove_auto_filter(self) -> None:
        """Remove the table's AutoFilter (no-op when absent). Row visibility is unchanged."""
        index = self._index()
        h = self._workbook._require()
        _check(LIB.fm_workbook_remove_table_auto_filter(h, _uint(index, "table_index")))

    def apply_auto_filter(self) -> None:
        """Hide the body rows the table's AutoFilter excludes and show the rest."""
        index = self._index()
        h = self._workbook._require()
        _check(LIB.fm_workbook_apply_table_auto_filter(h, _uint(index, "table_index")))

    def clear_auto_filter(self) -> None:
        """Drop every column criterion of the table's AutoFilter and show all its body rows."""
        index = self._index()
        h = self._workbook._require()
        _check(LIB.fm_workbook_clear_table_auto_filter(h, _uint(index, "table_index")))

    def evaluate_auto_filter(self) -> AutoFilterEval:
        """Evaluate the table's AutoFilter without changing the sheet, as :meth:`Sheet.evaluate_auto_filter` does."""
        return AutoFilterEval(
            *self._workbook._evaluate_auto_filter(
                LIB.fm_workbook_evaluate_table_auto_filter, self._index(), "table_index"
            )
        )
