//
// JsWorkbook sheet-view and per-sheet layout surface: zoom / freeze /
// tabHidden, the `<sheetProtection>` get/set bridge, and the column /
// row layout override iterators and mutators.

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

emscripten::val JsWorkbook::paginate(uint32_t sheet) const {
  emscripten::val result = emscripten::val::object();
  emscripten::val print_area = emscripten::val::array();
  emscripten::val horizontal_breaks = emscripten::val::array();
  emscripten::val vertical_breaks = emscripten::val::array();
  if (handle_ == nullptr) {
    result.set("status", error_status(7000));
    result.set("printArea", print_area);
    result.set("horizontalBreaks", horizontal_breaks);
    result.set("verticalBreaks", vertical_breaks);
    result.set("pageCount", 0);
    return result;
  }
  fm_pagination_t* pagination = nullptr;
  const fm_status_t rc = fm_workbook_paginate(handle_, sheet, &pagination);
  if (rc != 0) {
    result.set("status", error_status(rc));
    result.set("printArea", print_area);
    result.set("horizontalBreaks", horizontal_breaks);
    result.set("verticalBreaks", vertical_breaks);
    result.set("pageCount", 0);
    return result;
  }
  const std::size_t area_count = fm_pagination_print_area_count(pagination);
  for (std::size_t i = 0; i < area_count; ++i) {
    fm_print_range_t range{};
    if (fm_pagination_print_area_at(pagination, i, &range) != 0) {
      continue;
    }
    emscripten::val item = emscripten::val::object();
    item.set("firstRow", range.first_row);
    item.set("firstCol", range.first_col);
    item.set("lastRow", range.last_row);
    item.set("lastCol", range.last_col);
    print_area.set(i, item);
  }
  const std::size_t horizontal_count = fm_pagination_horizontal_break_count(pagination);
  for (std::size_t i = 0; i < horizontal_count; ++i) {
    uint32_t row = 0;
    if (fm_pagination_horizontal_break_at(pagination, i, &row) == 0) {
      horizontal_breaks.set(i, row);
    }
  }
  const std::size_t vertical_count = fm_pagination_vertical_break_count(pagination);
  for (std::size_t i = 0; i < vertical_count; ++i) {
    uint32_t col = 0;
    if (fm_pagination_vertical_break_at(pagination, i, &col) == 0) {
      vertical_breaks.set(i, col);
    }
  }
  result.set("status", ok_status());
  result.set("printArea", print_area);
  result.set("horizontalBreaks", horizontal_breaks);
  result.set("verticalBreaks", vertical_breaks);
  result.set("pageCount", fm_pagination_page_count(pagination));
  fm_pagination_destroy(pagination);
  return result;
}

JsSheetViewResult JsWorkbook::getSheetView(uint32_t sheet) const {
  JsSheetViewResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_sheet_view_t v{};
  fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.view.zoomScale = v.zoom_scale;
  r.view.freezeRows = v.freeze_rows;
  r.view.freezeCols = v.freeze_cols;
  r.view.tabHidden = v.tab_hidden;
  r.view.visibility = v.visibility;
  r.view.showGridLines = v.show_grid_lines;
  r.view.showRowColHeaders = v.show_row_col_headers;
  r.view.showZeros = v.show_zeros;
  r.view.rightToLeft = v.right_to_left;
  r.view.tabSelected = v.tab_selected;
  r.view.viewMode = v.view_mode == nullptr ? std::string() : v.view_mode;
  r.status = ok_status();
  return r;
}

JsSheetProtectionResult JsWorkbook::getSheetProtection(uint32_t sheet) const {
  JsSheetProtectionResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_sheet_protection_t p{};
  fm_status_t rc = fm_sheet_get_protection(handle_, sheet, &p);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.protection.enabled = p.enabled;
  r.protection.algorithmName = p.algorithm_name == nullptr ? std::string() : p.algorithm_name;
  r.protection.hashValue = p.hash_value == nullptr ? std::string() : p.hash_value;
  r.protection.saltValue = p.salt_value == nullptr ? std::string() : p.salt_value;
  r.protection.spinCount = p.spin_count;
  r.protection.legacyPassword = p.legacy_password == nullptr ? std::string() : p.legacy_password;
  r.protection.sheet = p.sheet;
  r.protection.objects = p.objects;
  r.protection.scenarios = p.scenarios;
  r.protection.formatCells = p.format_cells;
  r.protection.formatColumns = p.format_columns;
  r.protection.formatRows = p.format_rows;
  r.protection.insertColumns = p.insert_columns;
  r.protection.insertRows = p.insert_rows;
  r.protection.insertHyperlinks = p.insert_hyperlinks;
  r.protection.deleteColumns = p.delete_columns;
  r.protection.deleteRows = p.delete_rows;
  r.protection.selectLockedCells = p.select_locked_cells;
  r.protection.selectUnlockedCells = p.select_unlocked_cells;
  r.protection.sort = p.sort;
  r.protection.autoFilter = p.auto_filter;
  r.protection.pivotTables = p.pivot_tables;
  r.status = ok_status();
  return r;
}

JsStatus JsWorkbook::setSheetProtection(uint32_t sheet, JsSheetProtection in) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_sheet_protection_t p{};
  p.enabled = in.enabled;
  p.algorithm_name = in.algorithmName.c_str();
  p.hash_value = in.hashValue.c_str();
  p.salt_value = in.saltValue.c_str();
  p.spin_count = in.spinCount;
  p.legacy_password = in.legacyPassword.c_str();
  p.sheet = in.sheet;
  p.objects = in.objects;
  p.scenarios = in.scenarios;
  p.format_cells = in.formatCells;
  p.format_columns = in.formatColumns;
  p.format_rows = in.formatRows;
  p.insert_columns = in.insertColumns;
  p.insert_rows = in.insertRows;
  p.insert_hyperlinks = in.insertHyperlinks;
  p.delete_columns = in.deleteColumns;
  p.delete_rows = in.deleteRows;
  p.select_locked_cells = in.selectLockedCells;
  p.select_unlocked_cells = in.selectUnlockedCells;
  p.sort = in.sort;
  p.auto_filter = in.autoFilter;
  p.pivot_tables = in.pivotTables;
  fm_status_t rc = fm_sheet_set_protection(handle_, sheet, &p);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetZoom(uint32_t sheet, uint32_t zoomScale) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_zoom(handle_, sheet, zoomScale);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_freeze(handle_, sheet, freezeRows, freezeCols);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetTabHidden(uint32_t sheet, bool hidden) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_tab_hidden(handle_, sheet, hidden ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetVisibility(uint32_t sheet, int32_t visibility) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_visibility(handle_, sheet, visibility);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetShowGridLines(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_show_grid_lines(handle_, sheet, show ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetShowRowColHeaders(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_show_row_col_headers(handle_, sheet, show ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetShowZeros(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_show_zeros(handle_, sheet, show ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetRightToLeft(uint32_t sheet, bool rightToLeft) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_right_to_left(handle_, sheet, rightToLeft ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetTabSelected(uint32_t sheet, bool selected) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_tab_selected(handle_, sheet, selected ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setSheetViewMode(uint32_t sheet, std::string mode) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_view_mode(handle_, sheet, mode.c_str());
  return status_from_rc(rc);
}

// ---- Column layout overrides --------------------------------------------

emscripten::val JsWorkbook::getSheetColumns(uint32_t sheet) const {
  emscripten::val r = emscripten::val::object();
  emscripten::val columns = emscripten::val::array();
  if (handle_ == nullptr) {
    r.set("status", error_status(7000));
    r.set("columns", columns);
    return r;
  }
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_column_count(handle_, sheet, &count);
  if (rc != 0) {
    r.set("status", error_status(rc));
    r.set("columns", columns);
    return r;
  }
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_column_layout_t entry{};
    if (fm_sheet_get_column(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    emscripten::val out = emscripten::val::object();
    out.set("first", entry.first);
    out.set("last", entry.last);
    out.set("width", entry.width);
    out.set("hidden", entry.hidden);
    out.set("outlineLevel", static_cast<int32_t>(entry.outline_level));
    // Normalize legacy non-zero widths at the binding boundary as well as in
    // the C getter, so a mixed-version host still sees logical presence.
    out.set("hasWidth", (entry.has_width || entry.width != 0.0) ? 1 : 0);
    out.set("hasStyle", entry.has_style ? 1 : 0);
    out.set("styleXf", entry.style_xf);
    columns.set(emitted, out);
    ++emitted;
  }
  r.set("status", ok_status());
  r.set("columns", columns);
  return r;
}

JsStatus JsWorkbook::setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_column_width(handle_, sheet, first, last, width);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  if (level > 255U) {
    level = 255U;
  }
  fm_status_t rc = fm_sheet_set_column_outline(handle_, sheet, first, last, static_cast<uint8_t>(level));
  return status_from_rc(rc);
}

// ---- Row layout overrides ----------------------------------------------

emscripten::val JsWorkbook::getSheetRowOverrides(uint32_t sheet) const {
  emscripten::val r = emscripten::val::object();
  emscripten::val rows = emscripten::val::array();
  if (handle_ == nullptr) {
    r.set("status", error_status(7000));
    r.set("rows", rows);
    return r;
  }
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_row_override_count(handle_, sheet, &count);
  if (rc != 0) {
    r.set("status", error_status(rc));
    r.set("rows", rows);
    return r;
  }
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_row_layout_t entry{};
    if (fm_sheet_get_row_override(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    emscripten::val out = emscripten::val::object();
    out.set("row", entry.row);
    out.set("height", entry.height);
    out.set("hidden", entry.hidden);
    out.set("outlineLevel", static_cast<int32_t>(entry.outline_level));
    out.set("hasStyle", entry.has_style ? 1 : 0);
    out.set("styleXf", entry.style_xf);
    rows.set(emitted, out);
    ++emitted;
  }
  r.set("status", ok_status());
  r.set("rows", rows);
  return r;
}

JsStatus JsWorkbook::setRowHeight(uint32_t sheet, uint32_t row, double height) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_row_height(handle_, sheet, row, height);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setRowHidden(uint32_t sheet, uint32_t row, bool hidden) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_set_row_hidden(handle_, sheet, row, hidden ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setRowOutline(uint32_t sheet, uint32_t row, uint32_t level) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  if (level > 255U) {
    level = 255U;
  }
  fm_status_t rc = fm_sheet_set_row_outline(handle_, sheet, row, static_cast<uint8_t>(level));
  return status_from_rc(rc);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
