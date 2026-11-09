"""The :class:`PivotCache` handle: every operation that targets one pivot cache of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`PivotCache` per cache id; see
:class:`PivotCache`.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Dict, List, Optional, Union

from . import _structs as S
from ._c import LIB, _NoCopy, _sint, _uint, fm_value_t_size
from ._inputs import date_serial
from ._workbook import (
    _PATCH_CLEAR,
    _PATCH_SET,
    UNCHANGED,
    NativeValue,
    PivotCacheField,
    PivotCacheOptions,
    PivotCacheSharedItem,
    PivotCalculatedItem,
    PivotWorksheetSource,
    PivotWorksheetSourceInput,
    Unchanged,
    Value,
    _ArgScope,
    _check,
    _pack_pivot_worksheet_source,
    _pivot_cache_value_kind,
    _pivot_worksheet_source_from,
    _PivotCacheValue,
    _read_count,
    _read_cstr_out,
    _read_optional_cstr_out,
    to_native,
)

if TYPE_CHECKING:
    from ._workbook import Workbook

__all__ = ["PivotCache"]


class PivotCache(_NoCopy):
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

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain a PivotCache from Workbook.pivot_cache() or Workbook.add_pivot_cache()")

    @classmethod
    def _new(cls, workbook: Workbook, cache_id: int) -> "PivotCache":
        handle = object.__new__(cls)
        handle._workbook = workbook
        handle._cache_id = cache_id
        return handle

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
        return self._workbook._require(), _uint(self._cache_id, "cache_id")

    # -- Source and options ------------------------------------------------
    def get_worksheet_source(self) -> Optional[PivotWorksheetSource]:
        """Return the cache's worksheet-source record, or ``None`` when absent."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.PIVOT_WORKSHEET_SOURCE)
            _check(LIB.fm_workbook_get_pivot_cache_worksheet_source(h, cache_id, present, ptr))
            if not LIB.read_i32(present):
                return None
            return _pivot_worksheet_source_from(self._workbook, S.PIVOT_WORKSHEET_SOURCE.unpack(LIB, ptr))

    def set_worksheet_source(self, source: Optional[Union[PivotWorksheetSourceInput, PivotWorksheetSource]]) -> None:
        """Set the cache's worksheet source from a record, or clear it with ``None``.

        A :class:`PivotWorksheetSource` read back is accepted in place of a
        :class:`PivotWorksheetSourceInput`.

        Metadata only: records come from :meth:`add_field` / :meth:`add_record` /
        :meth:`set_record_value`; neither :meth:`Pivot.project` nor a cell edit
        re-reads the range.

        Raises:
          TypeError: when ``source`` is neither a source record nor ``None``.
          InvalidArgumentError: when the range's sheet disagrees with ``sheet``.
        """
        h, cache_id = self._call()
        with _ArgScope() as scope:
            ptr = 0 if source is None else _pack_pivot_worksheet_source(scope, self._workbook, source)
            _check(LIB.fm_workbook_set_pivot_cache_worksheet_source(h, cache_id, ptr))

    def get_options(self) -> PivotCacheOptions:
        """Read the options of the cache; see :class:`PivotCacheOptions`."""
        h, cache_id = self._call()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_CACHE_OPTIONS)
            _check(LIB.fm_workbook_get_pivot_cache_options(h, cache_id, ptr))
            d = S.PIVOT_CACHE_OPTIONS.unpack(LIB, ptr)
            return PivotCacheOptions(
                save_data=d["save_data"] != 0,
                missing_items_limit=d["missing_items_limit"] if d["has_missing_items_limit"] else None,
                refresh_on_load=d["refresh_on_load"] != 0,
            )

    def update_options(
        self,
        *,
        save_data: Union[bool, Unchanged] = UNCHANGED,
        missing_items_limit: Union[int, None, Unchanged] = UNCHANGED,
        refresh_on_load: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the options of the cache; an omitted argument keeps its setting.

        See :class:`PivotCacheOptions` for the members. ``missing_items_limit``
        given as ``None`` makes the limit automatic.

        Raises:
          TypeError: when ``save_data`` or ``refresh_on_load`` is not a ``bool``.
          InvalidArgumentError: for a negative ``missing_items_limit``.
        """
        for flag in (save_data, refresh_on_load):
            if flag is not UNCHANGED and not isinstance(flag, bool):
                raise TypeError(f"formulon: a pivot cache flag must be a bool, got {type(flag).__name__}")
        values: Dict[str, object] = {}
        for name, flag in (("save_data", save_data), ("refresh_on_load", refresh_on_load)):
            if flag is not UNCHANGED:
                values[name + "_op"] = _PATCH_SET
                values[name] = 1 if flag else 0
        if missing_items_limit is not UNCHANGED:
            if missing_items_limit is None:
                values["missing_items_limit_op"] = _PATCH_CLEAR
            else:
                values["missing_items_limit_op"] = _PATCH_SET
                values["missing_items_limit"] = _uint(missing_items_limit, "missing_items_limit")
        h, cache_id = self._call()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_CACHE_OPTIONS_PATCH)
            S.PIVOT_CACHE_OPTIONS_PATCH.pack(LIB, ptr, values)
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
                        LIB.fm_workbook_get_pivot_cache_field_name,
                        h,
                        cache_id,
                        _uint(i, "field_index", role="position"),
                    ),
                    formula=_read_optional_cstr_out(
                        LIB.fm_workbook_get_pivot_cache_field_formula,
                        h,
                        cache_id,
                        _uint(i, "field_index", role="position"),
                    ),
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
                    h, cache_id, _uint(field_index, "field_index", role="position"), formula_ptr
                )
            )

    def remove_calculated_field(self, field_index: int) -> None:
        """Remove calculated field ``field_index`` and its data fields in every pivot table."""
        h, cache_id = self._call()
        _check(
            LIB.fm_workbook_remove_pivot_cache_calculated_field(
                h, cache_id, _uint(field_index, "field_index", role="position")
            )
        )

    def list_calculated_items(self) -> List[PivotCalculatedItem]:
        """Every calculated item of the cache, in solve order; an item's position is its ``item_index``."""
        h, cache_id = self._call()
        count = _read_count(LIB.fm_workbook_get_pivot_cache_calculated_item_count, h, cache_id)
        items: List[PivotCalculatedItem] = []
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_CALCULATED_ITEM)
            for i in range(count):
                S.zero_struct(LIB, S.PIVOT_CALCULATED_ITEM, ptr)
                _check(
                    LIB.fm_workbook_get_pivot_cache_calculated_item_at(
                        h, cache_id, _uint(i, "item_index", role="position"), ptr
                    )
                )
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
                    h, cache_id, _uint(field_index, "field_index", role="position"), name_ptr, formula_ptr, out
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
                    h, cache_id, _uint(item_index, "item_index", role="position"), formula_ptr
                )
            )

    def move_calculated_item(self, item_index: int, new_index: int) -> None:
        """Move calculated item ``item_index`` to position ``new_index`` in solve order."""
        h, cache_id = self._call()
        _check(
            LIB.fm_workbook_move_pivot_cache_calculated_item(
                h,
                cache_id,
                _uint(item_index, "item_index", role="position"),
                _uint(new_index, "new_index", role="position"),
            )
        )

    def remove_calculated_item(self, item_index: int) -> None:
        """Remove calculated item ``item_index``, its item in the cache field and in every pivot table."""
        h, cache_id = self._call()
        _check(
            LIB.fm_workbook_remove_pivot_cache_calculated_item(
                h, cache_id, _uint(item_index, "item_index", role="position")
            )
        )

    # -- Shared items ------------------------------------------------------
    def list_field_shared_items(self, field_index: int) -> List[PivotCacheSharedItem]:
        """Every shared item of cache field ``field_index``, in item order.

        A date item holds its serial number as a number and has ``is_date`` set.
        """
        h, cache_id = self._call()
        field = _uint(field_index, "field_index", role="position")
        count = _read_count(LIB.fm_workbook_get_pivot_cache_field_shared_item_count, h, cache_id, field)
        items: List[PivotCacheSharedItem] = []
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            date_ptr = scope.out_ptr()
            for i in range(count):
                item = _uint(i, "item_index", role="position")
                _check(LIB.fm_workbook_get_pivot_cache_field_shared_item_value(h, cache_id, field, item, value_ptr))
                value = Value._from_wasm(value_ptr)
                _check(LIB.fm_workbook_is_pivot_cache_field_shared_item_date(h, cache_id, field, item, date_ptr))
                items.append(PivotCacheSharedItem(value=to_native(value), is_date=LIB.read_i32(date_ptr) != 0))
        return items

    def _date_serial(self, value: object) -> float:
        """The serial of a ``date`` or ``datetime`` cache value in the workbook's date system."""
        return date_serial(value, date1904=self._workbook.get_date1904())[0]

    def add_field_shared_item(self, field_index: int, value: _PivotCacheValue) -> int:
        """Append a shared item to cache field ``field_index`` and return its index.

        ``value`` is a number, a ``str``, a ``bool``, ``None`` (blank) or an
        :class:`ExcelError`, or a ``date`` / ``datetime`` (a date item stored as
        its serial in the workbook's date system, saved to the second); a number
        is a number. The index is the field's shared-item count before the call.

        Raises:
          TypeError: when ``value`` is none of the above.
        """
        kind = _pivot_cache_value_kind(value)
        h, cache_id = self._call()
        field = _uint(field_index, "field_index", role="position")
        with _ArgScope() as scope:
            out = scope.out_ptr()
            if kind == "date":
                _check(
                    LIB.fm_workbook_add_pivot_cache_field_shared_item_date(
                        h, cache_id, field, self._date_serial(value), out
                    )
                )
            elif kind == "number":
                _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_number(h, cache_id, field, float(value), out))
            elif kind == "bool":
                _check(
                    LIB.fm_workbook_add_pivot_cache_field_shared_item_bool(h, cache_id, field, 1 if value else 0, out)
                )
            elif kind == "blank":
                _check(LIB.fm_workbook_add_pivot_cache_field_shared_item_blank(h, cache_id, field, out))
            elif kind == "error":
                _check(
                    LIB.fm_workbook_add_pivot_cache_field_shared_item_error(
                        h, cache_id, field, _sint(value.code, "error"), out
                    )
                )
            else:
                _check(
                    LIB.fm_workbook_add_pivot_cache_field_shared_item_text(h, cache_id, field, scope.cstr(value), out)
                )
            return LIB.read_u32(out)

    def clear_field_shared_items(self, field_index: int) -> None:
        """Drop every shared item from cache field ``field_index``."""
        h, cache_id = self._call()
        _check(
            LIB.fm_workbook_clear_pivot_cache_field_shared_items(
                h, cache_id, _uint(field_index, "field_index", role="position")
            )
        )

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

    def get_record_value(self, record_index: int, field_index: int) -> NativeValue:
        """Read cache cell ``(record_index, field_index)`` as a native value.

        A cell that indexes a shared item reads the item's value, a cell its record does not reach reads
        ``None``, and a date reads as its serial number, as a shared item does.

        Raises:
          IndexOutOfRangeError: when ``record_index`` is not below the record count or ``field_index`` not
            below the cache's field count.
        """
        h, cache_id = self._call()
        record = _uint(record_index, "record_index", role="position")
        field = _uint(field_index, "field_index", role="position")
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            _check(LIB.fm_workbook_get_pivot_cache_record_value(h, cache_id, record, field, value_ptr))
            return to_native(Value._from_wasm(value_ptr))

    def set_record_value(self, record_index: int, field_index: int, value: _PivotCacheValue) -> None:
        """Set cache cell ``(record_index, field_index)``; ``value`` as in :meth:`add_field_shared_item`.

        Raises:
          TypeError: when ``value`` is not an accepted kind.
        """
        kind = _pivot_cache_value_kind(value)
        h, cache_id = self._call()
        record = _uint(record_index, "record_index", role="position")
        field = _uint(field_index, "field_index", role="position")
        if kind == "date":
            _check(LIB.fm_workbook_set_pivot_cache_record_date(h, cache_id, record, field, self._date_serial(value)))
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
