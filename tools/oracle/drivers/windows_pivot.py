"""PivotTable construction and grid readout through the Excel COM object model."""

from __future__ import annotations

from typing import Any, Dict, List

from .base import CaseResult
from .case_sheet import split_sheet_qualified_addr
from .cell_result import classify_value
from .windows_com import CellAdapter, cell_displayed_text, format_com_error

# Excel `XlConsolidationFunction` constants, keyed by the declarative
# `agg` name. Mirrors the `Aggregation` enum the C++ builder accepts.
# Numeric values are the documented Office automation constants so the
# driver does not depend on the `win32com` constant cache being warm.
_XL_CONSOLIDATION = {
    "Sum": -4157,
    "Count": -4112,
    "Average": -4106,
    "Max": -4136,
    "Min": -4139,
    "Product": -4149,
    "CountNumbers": -4113,
    "StdDev": -4155,
    "StdDevP": -4156,
    "Var": -4164,
    "VarP": -4165,
}

# Excel `XlPivotFieldOrientation` constants.
_XL_ORIENT_ROW = 1
_XL_ORIENT_COLUMN = 2
_XL_ORIENT_PAGE = 3
_XL_ORIENT_DATA = 4

# Excel `XlPivotTableSourceType.xlDatabase`.
_XL_DATABASE = 1

# Excel `XlLayoutRowType` values for `RowAxisLayout`.
_XL_COMPACT_ROW = 0
_XL_TABULAR_ROW = 1
_XL_OUTLINE_ROW = 2

_PIVOT_LAYOUT_MODES = {
    "Compact": _XL_COMPACT_ROW,
    "Tabular": _XL_TABULAR_ROW,
    "Outline": _XL_OUTLINE_ROW,
}


def build_pivot_table(wb, pivot_spec: Dict[str, Any]):
    """Creates a PivotTable in ``wb`` from a declarative ``pivot`` block.

    Returns the materialised ``PivotTable`` COM object. The caller reads
    ``TableRange2`` for the rendered grid and may then run post-build
    formula probes. The ``pivot`` block shape is documented on
    ``tests/oracle/workbook_builder.h``.
    """

    source = pivot_spec.get("source")
    anchor = pivot_spec.get("anchor")
    if not isinstance(source, str) or not source:
        raise RuntimeError("pivot block missing string 'source'")
    if not isinstance(anchor, str) or not anchor:
        raise RuntimeError("pivot block missing string 'anchor'")

    # Resolve the declarative anchor "Report!A1" to the destination
    # cell's COM Range, since CreatePivotTable accepts a Range here and
    # the string form has the same fully-qualified requirement as
    # SourceData. We also drive the source resolution through a Range
    # object so PivotCaches.Create does not have to parse the address
    # itself.
    src_sheet_name, src_addr = split_sheet_qualified_addr(source)
    anchor_sheet_name, anchor_addr = split_sheet_qualified_addr(anchor)
    if src_sheet_name is None:
        raise RuntimeError(f"pivot source must be sheet-qualified (e.g. 'Data!A1:C13'); got {source!r}")
    if anchor_sheet_name is None:
        raise RuntimeError(f"pivot anchor must be sheet-qualified (e.g. 'Report!A1'); got {anchor!r}")

    def _find_sheet(name: str):
        for sht in wb.sheets:
            if sht.name.casefold() == name.casefold():
                return sht
        return None

    src_sheet = _find_sheet(src_sheet_name)
    anchor_sheet = _find_sheet(anchor_sheet_name)
    if src_sheet is None:
        raise RuntimeError(f"pivot source references unknown sheet {src_sheet_name!r}")
    if anchor_sheet is None:
        raise RuntimeError(f"pivot anchor references unknown sheet {anchor_sheet_name!r}")

    source_range_api = src_sheet.range(src_addr).api
    anchor_range_api = anchor_sheet.range(anchor_addr).api

    # Activate the anchor sheet so CreatePivotTable's destination is on
    # the active sheet -- some Excel builds reject creating a pivot
    # whose destination is on an inactive sheet with E_INVALIDARG.
    try:
        anchor_sheet.activate()
    except Exception:
        pass

    api = wb.api
    cache = api.PivotCaches().Create(_XL_DATABASE, source_range_api)
    pivot = cache.CreatePivotTable(anchor_range_api, "FormulonPivot")

    for field_name in pivot_spec.get("row_fields") or []:
        pivot.PivotFields(field_name).Orientation = _XL_ORIENT_ROW
    for field_name in pivot_spec.get("col_fields") or []:
        pivot.PivotFields(field_name).Orientation = _XL_ORIENT_COLUMN
    for field_name in pivot_spec.get("page_fields") or []:
        pivot.PivotFields(field_name).Orientation = _XL_ORIENT_PAGE

    for data_field in pivot_spec.get("data_fields") or []:
        field_name = data_field.get("field")
        agg = data_field.get("agg", "Sum")
        consolidation = _XL_CONSOLIDATION.get(agg)
        if consolidation is None:
            raise RuntimeError(f"unknown aggregation {agg!r}")
        field = pivot.PivotFields(field_name)
        field.Orientation = _XL_ORIENT_DATA
        field.Function = consolidation

    # Manual item filters: hide the named items on their field.
    for filter_spec in pivot_spec.get("filters") or []:
        field_name = filter_spec.get("field")
        for item_name in filter_spec.get("hide") or []:
            try:
                pivot.PivotFields(field_name).PivotItems(item_name).Visible = False
            except Exception:
                pass

    layout = pivot_spec.get("layout")
    if isinstance(layout, str) and layout in _PIVOT_LAYOUT_MODES:
        try:
            pivot.RowAxisLayout(_PIVOT_LAYOUT_MODES[layout])
        except Exception:
            pass

    grand = pivot_spec.get("grand_totals") or {}
    if isinstance(grand, dict):
        if "rows" in grand:
            pivot.RowGrand = bool(grand["rows"])
        if "cols" in grand:
            pivot.ColumnGrand = bool(grand["cols"])

    pivot.RefreshTable()
    return pivot


def run_formula_probes(wb, probes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Write and read post-build formulas against a materialised PivotTable.

    Formula probes are intentionally evaluated only after ``RefreshTable``.
    A rendered PivotTable grid cannot exercise GETPIVOTDATA's page/data-axis
    routing, whereas a formula cell can.  The result shape matches the
    scalar grid records: ``{"kind": ..., "value": ...}`` or
    ``{"kind": "error", "code": "#REF!"}``.
    """

    if not isinstance(probes, list) or not probes:
        return []
    out: List[Dict[str, Any]] = []
    for probe in probes:
        if not isinstance(probe, dict):
            raise RuntimeError("pivot formula_probes entries must be objects")
        probe_id = probe.get("id")
        cell_ref = probe.get("cell")
        formula = probe.get("formula")
        if not isinstance(probe_id, str) or not probe_id:
            raise RuntimeError("pivot formula probe missing string 'id'")
        if not isinstance(cell_ref, str) or "!" not in cell_ref:
            raise RuntimeError(f"formula probe {probe_id!r} needs sheet-qualified cell")
        if not isinstance(formula, str) or not formula:
            raise RuntimeError(f"formula probe {probe_id!r} missing string 'formula'")
        sheet_name, bare_addr = split_sheet_qualified_addr(cell_ref)
        target = None
        for sht in wb.sheets:
            if sht.name.casefold() == sheet_name.casefold():
                target = sht
                break
        if target is None:
            raise RuntimeError(f"formula probe {probe_id!r} references unknown sheet {sheet_name!r}")
        result_cell = target.range(bare_addr)
        try:
            result_cell.formula2 = formula
        except Exception:
            result_cell.formula = formula

    try:
        wb.app.calculate()
    except Exception as exc:
        # Do not silently emit stale values from a prior probe. A formula
        # result such as #REF! is a normal cell value and does not raise here;
        # a COM calculation failure is a driver failure and must be visible.
        raise RuntimeError(f"Excel failed to calculate formula probes: {format_com_error(exc)}") from exc

    for probe in probes:
        sheet_name, bare_addr = split_sheet_qualified_addr(probe["cell"])
        target = next(sht for sht in wb.sheets if sht.name.casefold() == sheet_name.casefold())
        result = classify_value(CellAdapter(target.range(bare_addr).api), None, cell_displayed_text)
        out.append({"id": probe["id"], "cell": probe["cell"], "result": _grid_value_record(result)})
    return out


def read_pivot_grid(wb, table_range) -> "tuple[List[Dict[str, Any]], int, int]":
    """Reads every cell of ``table_range`` into an anchor-relative grid.

    Returns ``(grid, rows, cols)`` where ``grid`` is a list of
    ``{"r", "c", "value"}`` records (``r`` / ``c`` are 0-based offsets
    from the pivot anchor) and ``value`` is the normalised ``{kind,
    value}`` record produced by ``classify_value``.
    """

    rows = int(table_range.Rows.Count)
    cols = int(table_range.Columns.Count)
    top = int(table_range.Row)
    left = int(table_range.Column)
    sht = table_range.Worksheet

    grid: List[Dict[str, Any]] = []
    for r in range(rows):
        for c in range(cols):
            com_cell = sht.Cells(top + r, left + c)
            result = classify_value(CellAdapter(com_cell), None, cell_displayed_text)
            grid.append(
                {
                    "r": r,
                    "c": c,
                    "value": _grid_value_record(result),
                }
            )
    return grid, rows, cols


def _grid_value_record(result: CaseResult) -> Dict[str, Any]:
    """Shapes a ``CaseResult`` into the golden grid's ``{kind, value}``."""

    if result.kind == "blank":
        return {"kind": "blank"}
    if result.kind == "error":
        return {"kind": "error", "code": result.error_code or "#UNKNOWN!"}
    return {"kind": result.kind, "value": result.value}
