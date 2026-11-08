"""The :class:`PivotCache` handle: every operation that targets one pivot cache of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`PivotCache` per cache id; see
:class:`PivotCache`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Optional, Union

from . import _structs as S
from ._c import LIB, FormulonError, _sint, _uint, fm_value_t_size
from ._status import StatusCode
from ._workbook import (
    UNCHANGED,
    PivotCacheField,
    PivotCacheOptions,
    PivotCacheSharedItem,
    PivotCalculatedItem,
    PivotWorksheetSource,
    Unchanged,
    Value,
    _ArgScope,
    _check,
    _pivot_cache_value_kind,
    _PivotCacheValue,
    _read_count,
    _read_cstr_out,
)

if TYPE_CHECKING:
    from ._workbook import Workbook

__all__ = ["PivotCache"]


class PivotCache:
    """One pivot cache of a :class:`Workbook`, named by its :attr:`cache_id`.

    Obtain a handle from :meth:`Workbook.pivot_cache`,
    :meth:`Workbook.list_pivot_caches`, :meth:`Workbook.add_pivot_cache` or
    :meth:`Pivot.get_cache`. A workbook hands out one handle per cache id, so
    every lookup of the same cache returns the same object. Once the cache is
    removed (directly, or with its source sheet), every operation except
    reading :attr:`workbook` and :attr:`cache_id` raises
    :class:`NotFoundError`; a cache added again with the same id is addressed
    again. Remove a cache with :meth:`Workbook.remove_pivot_cache`.

    A cache reads its record with ``get_options()`` and
    ``get_worksheet_source()``; it has no ``get_properties()``.
    """

    __slots__ = ("_workbook", "_cache_id")

    def __init__(self, workbook: Workbook, cache_id: int) -> None:
        self._workbook = workbook
        self._cache_id = cache_id

    @property
    def workbook(self) -> Workbook:
        """The workbook this cache belongs to."""
        return self._workbook

    @property
    def cache_id(self) -> int:
        """The cache's id, the ``cache_id`` a pivot table's input takes."""
        return self._cache_id

    def __repr__(self) -> str:
        return f"PivotCache(cache_id={self._cache_id})"

    def _call(self) -> "tuple[int, int]":
        """The workbook's engine handle and the cache id as the ABI takes it.

        Raises:
          NotFoundError: when the workbook no longer has a cache with this id.
          WorkbookClosedError: when the workbook is closed.
        """
        h = self._workbook._require()
        # The raw calls report a missing cache id as an invalid argument, so look it up first.
        for index in range(_read_count(LIB.fm_workbook_get_pivot_cache_count, h)):
            if _read_count(LIB.fm_workbook_get_pivot_cache_id_at, h, _uint(index, "idx")) == self._cache_id:
                return h, _uint(self._cache_id, "cache_id")
        raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=("the pivot cache was removed", ""))

    # -- Source and options ------------------------------------------------
    def get_worksheet_source(self) -> Optional[PivotWorksheetSource]:
        """Return the cache's worksheet-source metadata, or ``None`` when absent."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            out_ref = scope.out_ptr()
            out_sheet = scope.out_ptr()
            out_name = scope.out_ptr()
            _check(LIB.fm_workbook_get_pivot_cache_worksheet_source(h, cache_id, present, out_ref, out_sheet, out_name))
            if not LIB.read_i32(present):
                return None
            return PivotWorksheetSource(
                LIB.read_cstr(LIB.read_u32(out_ref)) if LIB.read_u32(out_ref) else None,
                LIB.read_cstr(LIB.read_u32(out_sheet)) if LIB.read_u32(out_sheet) else None,
                LIB.read_cstr(LIB.read_u32(out_name)) if LIB.read_u32(out_name) else None,
            )

    def set_worksheet_source(
        self,
        source: None = UNCHANGED,
        /,
        *,
        ref: Optional[str] = None,
        sheet: Optional[str] = None,
        name: Optional[str] = None,
    ) -> None:
        """Set the cache's worksheet-source metadata, or clear it.

        ``set_worksheet_source(None)`` clears it. Otherwise the keywords set
        it: an attribute not given is left out of the source, so a call
        without arguments writes a source that has no attributes.

        Raises:
          TypeError: when the first argument is not ``None``, or is ``None``
            together with a keyword argument.
        """
        clearing = source is not UNCHANGED
        if clearing:
            if source is not None:
                raise TypeError(f"formulon: the source argument clears the source and must be None, got {source!r}")
            if ref is not None or sheet is not None or name is not None:
                raise TypeError("formulon: clearing the worksheet source takes no keyword arguments")
        h, cache_id = self._call()
        with _ArgScope() as scope:
            pointers = [0, 0, 0] if clearing else [scope.opt_cstr(value) for value in (ref, sheet, name)]
            _check(LIB.fm_workbook_set_pivot_cache_worksheet_source(h, cache_id, 0 if clearing else 1, *pointers))

    def get_options(self) -> PivotCacheOptions:
        """Read the options of the cache; see :class:`PivotCacheOptions`."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_CACHE_OPTIONS)
            _check(LIB.fm_workbook_get_pivot_cache_options(h, cache_id, ptr))
            d = S.PIVOT_CACHE_OPTIONS.unpack(LIB, ptr)
            return PivotCacheOptions(
                save_data=d["save_data"] != 0,
                missing_items_limit=d["missing_items_limit"],
                refresh_on_load=d["refresh_on_load"] != 0,
            )

    def update_options(
        self,
        *,
        save_data: Union[bool, Unchanged] = UNCHANGED,
        missing_items_limit: Union[int, Unchanged] = UNCHANGED,
        refresh_on_load: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the options of the cache; an omitted argument keeps its setting.

        See :class:`PivotCacheOptions` for the members.

        Raises:
          TypeError: when ``save_data`` or ``refresh_on_load`` is not a ``bool``.
        """
        for flag in (save_data, refresh_on_load):
            if flag is not UNCHANGED and not isinstance(flag, bool):
                raise TypeError(f"formulon: a pivot cache flag must be a bool, got {type(flag).__name__}")
        current = self.get_options()
        h, cache_id = self._call()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_CACHE_OPTIONS)
            S.PIVOT_CACHE_OPTIONS.pack(
                LIB,
                ptr,
                {
                    "save_data": 1 if (current.save_data if save_data is UNCHANGED else save_data) else 0,
                    "missing_items_limit": _sint(
                        current.missing_items_limit if missing_items_limit is UNCHANGED else missing_items_limit,
                        "missing_items_limit",
                    ),
                    "refresh_on_load": 1
                    if (current.refresh_on_load if refresh_on_load is UNCHANGED else refresh_on_load)
                    else 0,
                },
            )
            _check(LIB.fm_workbook_update_pivot_cache_options(h, cache_id, ptr))

    # -- Fields ------------------------------------------------------------
    def list_fields(self) -> List[PivotCacheField]:
        """Every field of the cache, in field order; a field's position is its ``field_index``."""
        h, cache_id = self._call()
        fields: List[PivotCacheField] = []
        for i in range(_read_count(LIB.fm_workbook_get_pivot_cache_field_count, h, cache_id)):
            fields.append(
                PivotCacheField(
                    name=_read_cstr_out(
                        LIB.fm_workbook_get_pivot_cache_field_name, h, cache_id, _uint(i, "field_index")
                    ),
                    formula=_read_cstr_out(
                        LIB.fm_workbook_get_pivot_cache_field_formula, h, cache_id, _uint(i, "field_index")
                    )
                    or None,
                )
            )
        return fields

    def add_field(self, name: str) -> int:
        """Append a field named ``name`` to the cache; return its index."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_pivot_cache_field(h, cache_id, name_ptr, out))
            return LIB.read_u32(out)

    def clear_fields(self) -> None:
        """Drop every field (and record) from the cache."""
        h, cache_id = self._call()
        _check(LIB.fm_workbook_clear_pivot_cache_fields(h, cache_id))

    # -- Calculated fields and items ---------------------------------------
    def add_calculated_field(self, name: str, formula: str) -> int:
        """Add calculated field ``name`` as the last cache field; return its index.

        The formula is in English canonical form and names cache fields only,
        quoting a name with spaces as ``'Amt Net'``. A leading ``=`` is optional.
        """
        h, cache_id = self._call()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            formula_ptr = scope.cstr(formula)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_pivot_cache_calculated_field(h, cache_id, name_ptr, formula_ptr, out))
            return LIB.read_u32(out)

    def set_calculated_field_formula(self, field_index: int, formula: str) -> None:
        """Replace the formula of calculated field ``field_index``."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            _check(
                LIB.fm_workbook_set_pivot_cache_calculated_field_formula(
                    h, cache_id, _uint(field_index, "field_index"), formula_ptr
                )
            )

    def remove_calculated_field(self, field_index: int) -> None:
        """Remove calculated field ``field_index`` and its data fields in every pivot table."""
        h, cache_id = self._call()
        _check(LIB.fm_workbook_remove_pivot_cache_calculated_field(h, cache_id, _uint(field_index, "field_index")))

    def list_calculated_items(self) -> List[PivotCalculatedItem]:
        """Every calculated item of the cache, in solve order; an item's position is its ``item_index``."""
        h, cache_id = self._call()
        count = _read_count(LIB.fm_workbook_get_pivot_cache_calculated_item_count, h, cache_id)
        items: List[PivotCalculatedItem] = []
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_CALCULATED_ITEM)
            for i in range(count):
                S.zero_struct(LIB, S.PIVOT_CALCULATED_ITEM, ptr)
                _check(LIB.fm_workbook_get_pivot_cache_calculated_item(h, cache_id, _uint(i, "item_index"), ptr))
                d = S.PIVOT_CALCULATED_ITEM.unpack(LIB, ptr)
                items.append(
                    PivotCalculatedItem(
                        field_index=d["field_index"],
                        name=LIB.read_cstr(d["name"]),
                        formula=LIB.read_cstr(d["formula"]),
                        scoped=d["scoped"] != 0,
                    )
                )
        return items

    def add_calculated_item(self, field_index: int, name: str, formula: str) -> int:
        """Add calculated item ``name`` to cache field ``field_index``; return its position, last in solve order.

        The formula names items of the same field, as ``North``, ``'North East'``,
        ``Region[North]``, ``Region[2]`` or ``Region[-1]``.
        """
        h, cache_id = self._call()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            formula_ptr = scope.cstr(formula)
            out = scope.out_ptr()
            _check(
                LIB.fm_workbook_add_pivot_cache_calculated_item(
                    h, cache_id, _uint(field_index, "field_index"), name_ptr, formula_ptr, out
                )
            )
            return LIB.read_u32(out)

    def set_calculated_item_formula(self, item_index: int, formula: str) -> None:
        """Replace the formula of calculated item ``item_index``."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            _check(
                LIB.fm_workbook_set_pivot_cache_calculated_item_formula(
                    h, cache_id, _uint(item_index, "item_index"), formula_ptr
                )
            )

    def move_calculated_item(self, item_index: int, new_index: int) -> None:
        """Move calculated item ``item_index`` to position ``new_index`` in solve order."""
        h, cache_id = self._call()
        _check(
            LIB.fm_workbook_move_pivot_cache_calculated_item(
                h, cache_id, _uint(item_index, "item_index"), _uint(new_index, "new_index")
            )
        )

    def remove_calculated_item(self, item_index: int) -> None:
        """Remove calculated item ``item_index``, its item in the cache field and in every pivot table."""
        h, cache_id = self._call()
        _check(LIB.fm_workbook_remove_pivot_cache_calculated_item(h, cache_id, _uint(item_index, "item_index")))

    # -- Shared items ------------------------------------------------------
    def list_field_shared_items(self, field_index: int) -> List[PivotCacheSharedItem]:
        """Every shared item of cache field ``field_index``, in item order.

        A date item holds its serial number as a number and has ``is_date`` set.
        """
        h, cache_id = self._call()
        field = _uint(field_index, "field_index")
        count = _read_count(LIB.fm_workbook_get_pivot_cache_field_shared_item_count, h, cache_id, field)
        items: List[PivotCacheSharedItem] = []
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            date_ptr = scope.out_ptr()
            for i in range(count):
                item = _uint(i, "item_index")
                _check(LIB.fm_workbook_get_pivot_cache_field_shared_item_value(h, cache_id, field, item, value_ptr))
                value = Value._from_wasm(value_ptr)
                _check(LIB.fm_workbook_is_pivot_cache_field_shared_item_date(h, cache_id, field, item, date_ptr))
                items.append(PivotCacheSharedItem(value=value, is_date=LIB.read_i32(date_ptr) != 0))
        return items

    def add_field_shared_item(self, field_index: int, value: _PivotCacheValue, *, date: bool = False) -> None:
        """Append a shared item to cache field ``field_index``.

        ``value`` is a number, a ``str``, a ``bool``, ``None`` (blank) or an
        :class:`ExcelError`. ``date=True`` stores a number as a date serial in
        the workbook's date system, saved to the second.

        Raises:
          TypeError: when ``value`` is none of the above, or ``date`` is set
            for a value that is not a number.
        """
        kind = _pivot_cache_value_kind(value, date)
        h, cache_id = self._call()
        field = _uint(field_index, "field_index")
        if kind == "date":
            _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_date(h, cache_id, field, float(value)))
        elif kind == "number":
            _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_number(h, cache_id, field, float(value)))
        elif kind == "bool":
            _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_bool(h, cache_id, field, 1 if value else 0))
        elif kind == "blank":
            _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_blank(h, cache_id, field))
        elif kind == "error":
            _check(
                LIB.fm_workbook_add_pivot_cache_field_shared_item_error(h, cache_id, field, _sint(value.code, "error"))
            )
        else:
            with _ArgScope() as scope:
                text = scope.cstr(value)
                _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_text(h, cache_id, field, text))

    def clear_field_shared_items(self, field_index: int) -> None:
        """Drop every shared item from cache field ``field_index``."""
        h, cache_id = self._call()
        _check(LIB.fm_workbook_clear_pivot_cache_field_shared_items(h, cache_id, _uint(field_index, "field_index")))

    # -- Records -----------------------------------------------------------
    def get_record_count(self) -> int:
        """Return the number of records on the cache."""
        h, cache_id = self._call()
        return _read_count(LIB.fm_workbook_get_pivot_cache_record_count, h, cache_id)

    def add_record(self) -> int:
        """Append an empty record to the cache; return its index."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_pivot_cache_record(h, cache_id, out))
            return LIB.read_u32(out)

    def clear_records(self) -> None:
        """Drop every record from the cache."""
        h, cache_id = self._call()
        _check(LIB.fm_workbook_clear_pivot_cache_records(h, cache_id))

    def set_record_value(
        self, record_index: int, field_index: int, value: _PivotCacheValue, *, date: bool = False
    ) -> None:
        """Set cache cell ``(record_index, field_index)``; ``value`` and ``date`` as in :meth:`add_field_shared_item`.

        Raises:
          TypeError: when ``value`` is not an accepted kind, or ``date`` is set
            for a value that is not a number.
        """
        kind = _pivot_cache_value_kind(value, date)
        h, cache_id = self._call()
        record = _uint(record_index, "record_index")
        field = _uint(field_index, "field_index")
        if kind == "date":
            _check(LIB.fm_workbook_set_pivot_cache_record_date(h, cache_id, record, field, float(value)))
        elif kind == "number":
            _check(LIB.fm_workbook_set_pivot_cache_record_number(h, cache_id, record, field, float(value)))
        elif kind == "bool":
            _check(LIB.fm_workbook_set_pivot_cache_record_bool(h, cache_id, record, field, 1 if value else 0))
        elif kind == "blank":
            _check(LIB.fm_workbook_set_pivot_cache_record_blank(h, cache_id, record, field))
        elif kind == "error":
            _check(LIB.fm_workbook_set_pivot_cache_record_error(h, cache_id, record, field, _sint(value.code, "error")))
        else:
            with _ArgScope() as scope:
                text = scope.cstr(value)
                _check(LIB.fm_workbook_set_pivot_cache_record_text(h, cache_id, record, field, text))
