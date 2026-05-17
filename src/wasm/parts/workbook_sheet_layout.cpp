// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

// ---- Column layout overrides --------------------------------------------

JsColumnsResult JsWorkbook::getSheetColumns(uint32_t sheet) const {
  JsColumnsResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_column_count(handle_, sheet, &count);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.columns.reserve(count);
  for (std::size_t i = 0; i < count; ++i) {
    fm_column_layout_t entry{};
    if (fm_sheet_get_column(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    JsColumnLayout out;
    out.first = entry.first;
    out.last = entry.last;
    out.width = entry.width;
    out.hidden = entry.hidden;
    out.outlineLevel = static_cast<int32_t>(entry.outline_level);
    r.columns.push_back(out);
  }
  r.status = ok_status();
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

JsRowsResult JsWorkbook::getSheetRowOverrides(uint32_t sheet) const {
  JsRowsResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_row_override_count(handle_, sheet, &count);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.rows.reserve(count);
  for (std::size_t i = 0; i < count; ++i) {
    fm_row_layout_t entry{};
    if (fm_sheet_get_row_override(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    JsRowLayout out;
    out.row = entry.row;
    out.height = entry.height;
    out.hidden = entry.hidden;
    out.outlineLevel = static_cast<int32_t>(entry.outline_level);
    r.rows.push_back(out);
  }
  r.status = ok_status();
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
