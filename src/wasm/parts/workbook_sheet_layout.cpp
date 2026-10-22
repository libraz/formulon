//
// JsWorkbook sheet-view and per-sheet layout surface: zoom / freeze /
// tabHidden, the `<sheetProtection>` get/set bridge, and the column /
// row layout override iterators and mutators.

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

using EntryCountGetter = fm_status_t (*)(const fm_workbook_t*, size_t, size_t*);

// Reads entry `index` of `sheet` into `*out`; false skips an unreadable entry.
using EntryItemGetter = bool (*)(const fm_workbook_t*, uint32_t, std::size_t, emscripten::val*);

// Builds a `{ status, <key>: [...] }` enumeration result. Only a failed
// count fails the call; an unreadable entry is left out of the list.
emscripten::val entry_list(const fm_workbook_t* handle, uint32_t sheet, EntryCountGetter count_getter, const char* key,
                           EntryItemGetter item_getter) {
  emscripten::val r = emscripten::val::object();
  emscripten::val list = emscripten::val::array();
  std::size_t count = 0;
  const fm_status_t rc = handle == nullptr ? kBindingInvalidHandle : count_getter(handle, sheet, &count);
  if (rc != 0) {
    r.set("status", error_status(rc));
    r.set(key, list);
    return r;
  }
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    emscripten::val out = emscripten::val::undefined();
    if (!item_getter(handle, sheet, i, &out)) {
      continue;
    }
    list.set(emitted, out);
    ++emitted;
  }
  r.set("status", ok_status());
  r.set(key, list);
  return r;
}

bool column_entry(const fm_workbook_t* handle, uint32_t sheet, std::size_t index, emscripten::val* out) {
  fm_column_layout_t entry{};
  if (fm_sheet_get_column(handle, sheet, index, &entry) != 0) {
    return false;
  }
  *out = emscripten::val::object();
  out->set("first", entry.first);
  out->set("last", entry.last);
  out->set("width", entry.width);
  out->set("hidden", entry.hidden);
  out->set("outlineLevel", static_cast<int32_t>(entry.outline_level));
  out->set("collapsed", entry.collapsed ? 1 : 0);
  // Normalize legacy non-zero widths at the binding boundary as well as in
  // the C getter, so a mixed-version host still sees logical presence.
  out->set("hasWidth", (entry.has_width || entry.width != 0.0) ? 1 : 0);
  out->set("hasStyle", entry.has_style ? 1 : 0);
  out->set("styleXf", entry.style_xf);
  return true;
}

bool row_entry(const fm_workbook_t* handle, uint32_t sheet, std::size_t index, emscripten::val* out) {
  fm_row_layout_t entry{};
  if (fm_sheet_get_row_override(handle, sheet, index, &entry) != 0) {
    return false;
  }
  *out = emscripten::val::object();
  out->set("row", entry.row);
  out->set("height", entry.height);
  out->set("hidden", entry.hidden);
  out->set("outlineLevel", static_cast<int32_t>(entry.outline_level));
  out->set("collapsed", entry.collapsed ? 1 : 0);
  out->set("hasStyle", entry.has_style ? 1 : 0);
  out->set("styleXf", entry.style_xf);
  out->set("hasHeight", entry.has_height ? 1 : 0);
  out->set("customHeight", entry.custom_height ? 1 : 0);
  return true;
}

uint8_t clamp_outline_level(uint32_t level) {
  return static_cast<uint8_t>(level > 255U ? 255U : level);
}

emscripten::val js_rect_pt(const fm_rect_pt& r) {
  emscripten::val o = emscripten::val::object();
  o.set("x", r.x);
  o.set("y", r.y);
  o.set("width", r.width);
  o.set("height", r.height);
  return o;
}

/// Writes the pagination detail keys onto `result`. A null `pagination`
/// yields the all-zero record, so the key set does not depend on whether
/// the layout succeeded.
void set_pagination_detail(emscripten::val* result, const fm_pagination_t* pagination) {
  fm_paper_info paper{};
  fm_margins_pt margins{};
  fm_rect_pt printable{};
  double scale = 0.0;
  int32_t page_order = 0;
  fm_print_titles titles{};
  std::size_t page_count = 0;
  std::size_t horizontal_count = 0;
  std::size_t vertical_count = 0;
  if (pagination != nullptr) {
    fm_pagination_paper(pagination, &paper);
    fm_pagination_margins(pagination, &margins);
    fm_pagination_printable(pagination, &printable);
    fm_pagination_scale(pagination, &scale);
    fm_pagination_page_order(pagination, &page_order);
    fm_pagination_print_titles(pagination, &titles);
    page_count = fm_pagination_page_count(pagination);
    horizontal_count = fm_pagination_horizontal_break_count(pagination);
    vertical_count = fm_pagination_vertical_break_count(pagination);
  }
  emscripten::val paper_val = emscripten::val::object();
  paper_val.set("widthPt", paper.width_pt);
  paper_val.set("heightPt", paper.height_pt);
  paper_val.set("landscape", paper.landscape != 0);
  paper_val.set("known", paper.known != 0);
  result->set("paper", paper_val);
  emscripten::val margins_val = emscripten::val::object();
  margins_val.set("left", margins.left);
  margins_val.set("right", margins.right);
  margins_val.set("top", margins.top);
  margins_val.set("bottom", margins.bottom);
  margins_val.set("header", margins.header);
  margins_val.set("footer", margins.footer);
  result->set("margins", margins_val);
  result->set("printable", js_rect_pt(printable));
  result->set("scale", scale);
  result->set("pageOrder", page_order);
  emscripten::val titles_val = emscripten::val::object();
  titles_val.set("hasRows", titles.has_rows != 0);
  titles_val.set("firstRow", titles.first_row);
  titles_val.set("lastRow", titles.last_row);
  titles_val.set("hasCols", titles.has_cols != 0);
  titles_val.set("firstCol", titles.first_col);
  titles_val.set("lastCol", titles.last_col);
  result->set("printTitles", titles_val);
  emscripten::val pages = emscripten::val::array();
  for (std::size_t i = 0; i < page_count; ++i) {
    fm_page_layout page{};
    if (fm_pagination_page_at(pagination, i, &page) != 0) {
      continue;
    }
    emscripten::val item = emscripten::val::object();
    item.set("areaIndex", page.area_index);
    item.set("firstRow", page.first_row);
    item.set("lastRow", page.last_row);
    item.set("firstCol", page.first_col);
    item.set("lastCol", page.last_col);
    item.set("originXPt", page.origin_x_pt);
    item.set("originYPt", page.origin_y_pt);
    item.set("widthPt", page.width_pt);
    item.set("heightPt", page.height_pt);
    pages.call<void>("push", item);
  }
  result->set("pages", pages);
  emscripten::val horizontal_manual = emscripten::val::array();
  for (std::size_t i = 0; i < horizontal_count; ++i) {
    int32_t manual = 0;
    fm_pagination_horizontal_break_is_manual(pagination, i, &manual);
    horizontal_manual.call<void>("push", manual != 0);
  }
  result->set("horizontalBreakManual", horizontal_manual);
  emscripten::val vertical_manual = emscripten::val::array();
  for (std::size_t i = 0; i < vertical_count; ++i) {
    int32_t manual = 0;
    fm_pagination_vertical_break_is_manual(pagination, i, &manual);
    vertical_manual.call<void>("push", manual != 0);
  }
  result->set("verticalBreakManual", vertical_manual);
}

}  // namespace

emscripten::val JsWorkbook::paginate(uint32_t sheet) const {
  emscripten::val result = emscripten::val::object();
  emscripten::val print_area = emscripten::val::array();
  emscripten::val horizontal_breaks = emscripten::val::array();
  emscripten::val vertical_breaks = emscripten::val::array();
  if (handle_ == nullptr) {
    result.set("status", invalid_handle_status());
    result.set("printArea", print_area);
    result.set("horizontalBreaks", horizontal_breaks);
    result.set("verticalBreaks", vertical_breaks);
    result.set("pageCount", 0);
    set_pagination_detail(&result, nullptr);
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
    set_pagination_detail(&result, nullptr);
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
  set_pagination_detail(&result, pagination);
  fm_pagination_destroy(pagination);
  return result;
}

JsSheetViewResult JsWorkbook::getSheetView(uint32_t sheet) const {
  JsSheetViewResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
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
  r.view.viewMode = string_from_cstr(v.view_mode);
  r.view.topLeftRow = v.top_left_row;
  r.view.topLeftCol = v.top_left_col;
  r.view.paneTopLeftRow = v.pane_top_left_row;
  r.view.paneTopLeftCol = v.pane_top_left_col;
  r.view.splitX = v.split_x;
  r.view.splitY = v.split_y;
  r.view.showFormulas = v.show_formulas;
  r.view.showOutlineSymbols = v.show_outline_symbols;
  r.status = ok_status();
  return r;
}

JsSheetProtectionResult JsWorkbook::getSheetProtection(uint32_t sheet) const {
  JsSheetProtectionResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  fm_sheet_protection_t p{};
  fm_status_t rc = fm_sheet_get_protection(handle_, sheet, &p);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.protection.enabled = p.enabled;
  r.protection.algorithmName = string_from_cstr(p.algorithm_name);
  r.protection.hashValue = string_from_cstr(p.hash_value);
  r.protection.saltValue = string_from_cstr(p.salt_value);
  r.protection.spinCount = p.spin_count;
  r.protection.legacyPassword = string_from_cstr(p.legacy_password);
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

JsStatus JsWorkbook::setSheetProtection(uint32_t sheet, emscripten::val input) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setSheetProtection");
  JsSheetProtection in{};
  in.enabled = reader.i32(input, "enabled", 0, "protection.enabled");
  in.algorithmName = reader.string(input, "algorithmName", "protection.algorithmName");
  in.hashValue = reader.string(input, "hashValue", "protection.hashValue");
  in.saltValue = reader.string(input, "saltValue", "protection.saltValue");
  in.spinCount = reader.u32(input, "spinCount", 0U, "protection.spinCount");
  in.legacyPassword = reader.string(input, "legacyPassword", "protection.legacyPassword");
  in.sheet = reader.i32(input, "sheet", 0, "protection.sheet");
  in.objects = reader.i32(input, "objects", 0, "protection.objects");
  in.scenarios = reader.i32(input, "scenarios", 0, "protection.scenarios");
  in.formatCells = reader.i32(input, "formatCells", 0, "protection.formatCells");
  in.formatColumns = reader.i32(input, "formatColumns", 0, "protection.formatColumns");
  in.formatRows = reader.i32(input, "formatRows", 0, "protection.formatRows");
  in.insertColumns = reader.i32(input, "insertColumns", 0, "protection.insertColumns");
  in.insertRows = reader.i32(input, "insertRows", 0, "protection.insertRows");
  in.insertHyperlinks = reader.i32(input, "insertHyperlinks", 0, "protection.insertHyperlinks");
  in.deleteColumns = reader.i32(input, "deleteColumns", 0, "protection.deleteColumns");
  in.deleteRows = reader.i32(input, "deleteRows", 0, "protection.deleteRows");
  in.selectLockedCells = reader.i32(input, "selectLockedCells", 0, "protection.selectLockedCells");
  in.selectUnlockedCells = reader.i32(input, "selectUnlockedCells", 0, "protection.selectUnlockedCells");
  in.sort = reader.i32(input, "sort", 0, "protection.sort");
  in.autoFilter = reader.i32(input, "autoFilter", 0, "protection.autoFilter");
  in.pivotTables = reader.i32(input, "pivotTables", 0, "protection.pivotTables");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
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
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_zoom(handle_, sheet, zoomScale));
}

JsStatus JsWorkbook::setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_freeze(handle_, sheet, freezeRows, freezeCols));
}

JsStatus JsWorkbook::setSheetTabHidden(uint32_t sheet, bool hidden) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_tab_hidden(handle_, sheet, hidden ? 1 : 0));
}

JsStatus JsWorkbook::setSheetVisibility(uint32_t sheet, int32_t visibility) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_visibility(handle_, sheet, visibility));
}

JsStatus JsWorkbook::setSheetShowGridLines(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_show_grid_lines(handle_, sheet, show ? 1 : 0));
}

JsStatus JsWorkbook::setSheetShowRowColHeaders(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_show_row_col_headers(handle_, sheet, show ? 1 : 0));
}

JsStatus JsWorkbook::setSheetShowZeros(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_show_zeros(handle_, sheet, show ? 1 : 0));
}

JsStatus JsWorkbook::setSheetShowFormulas(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_show_formulas(handle_, sheet, show ? 1 : 0));
}

JsStatus JsWorkbook::setSheetShowOutlineSymbols(uint32_t sheet, bool show) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_show_outline_symbols(handle_, sheet, show ? 1 : 0));
}

JsStatus JsWorkbook::setSheetRightToLeft(uint32_t sheet, bool rightToLeft) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_right_to_left(handle_, sheet, rightToLeft ? 1 : 0));
}

JsStatus JsWorkbook::setSheetTopLeft(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_top_left(handle_, sheet, row, col));
}

JsStatus JsWorkbook::setSheetPaneTopLeft(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_pane_top_left(handle_, sheet, row, col));
}

JsStatus JsWorkbook::setSheetSplit(uint32_t sheet, double xTwips, double yTwips) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_split(handle_, sheet, xTwips, yTwips));
}

emscripten::val JsWorkbook::getSheetSelection(uint32_t sheet) const {
  emscripten::val o = emscripten::val::object();
  emscripten::val ranges = emscripten::val::array();
  uint32_t row = 0;
  uint32_t col = 0;
  size_t total = 0;
  fm_status_t rc = handle_ != nullptr ? fm_sheet_get_selection(handle_, sheet, &row, &col, nullptr, 0, &total)
                                      : kBindingInvalidHandle;
  std::vector<fm_merge_range> found(total);
  if (rc == 0) {
    rc = fm_sheet_get_selection(handle_, sheet, &row, &col, found.data(), found.size(), &total);
  }
  for (size_t i = 0; rc == 0 && i < found.size(); ++i) {
    ranges.call<void>("push", merge_range_to_val(found[i]));
  }
  o.set("status", status_from_rc(rc));
  o.set("activeRow", rc == 0 ? row : 0U);
  o.set("activeCol", rc == 0 ? col : 0U);
  o.set("ranges", ranges);
  return o;
}

JsStatus JsWorkbook::setSheetSelection(uint32_t sheet, emscripten::val selection) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setSheetSelection");
  const uint32_t row = reader.u32(selection, "activeRow", 0U, "selection.activeRow");
  const uint32_t col = reader.u32(selection, "activeCol", 0U, "selection.activeCol");
  const std::vector<fm_merge_range> ranges = js_pull_ranges(selection, "ranges", &reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_set_selection(handle_, sheet, row, col, ranges.data(), ranges.size()));
}

JsStatus JsWorkbook::setSheetTabSelected(uint32_t sheet, bool selected) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_tab_selected(handle_, sheet, selected ? 1 : 0));
}

JsStatus JsWorkbook::setSheetViewMode(uint32_t sheet, std::string mode) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_view_mode(handle_, sheet, mode.c_str()));
}

// ---- Column layout overrides --------------------------------------------

emscripten::val JsWorkbook::getSheetColumns(uint32_t sheet) const {
  return entry_list(handle_, sheet, &fm_sheet_get_column_count, "columns", &column_entry);
}

JsStatus JsWorkbook::setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_column_width(handle_, sheet, first, last, width));
}

JsStatus JsWorkbook::setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0));
}

JsStatus JsWorkbook::setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_column_outline(handle_, sheet, first, last, clamp_outline_level(level)));
}

JsStatus JsWorkbook::setColumnCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_column_collapsed(handle_, sheet, first, last, collapsed ? 1 : 0));
}

JsStatus JsWorkbook::setColumnStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_column_style(handle_, sheet, first, last, xfIndex));
}

JsStatus JsWorkbook::clearColumnStyle(uint32_t sheet, uint32_t first, uint32_t last) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_column_style(handle_, sheet, first, last));
}

// ---- Row layout overrides ----------------------------------------------

emscripten::val JsWorkbook::getSheetRowOverrides(uint32_t sheet) const {
  return entry_list(handle_, sheet, &fm_sheet_get_row_override_count, "rows", &row_entry);
}

JsStatus JsWorkbook::setRowHeight(uint32_t sheet, uint32_t row, double height) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_row_height(handle_, sheet, row, height));
}

JsStatus JsWorkbook::setRowHidden(uint32_t sheet, uint32_t row, bool hidden) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_row_hidden(handle_, sheet, row, hidden ? 1 : 0));
}

JsStatus JsWorkbook::setRowOutline(uint32_t sheet, uint32_t row, uint32_t level) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_row_outline(handle_, sheet, row, clamp_outline_level(level)));
}

JsStatus JsWorkbook::setRowCollapsed(uint32_t sheet, uint32_t row, bool collapsed) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_row_collapsed(handle_, sheet, row, collapsed ? 1 : 0));
}

JsStatus JsWorkbook::setRowStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_row_style(handle_, sheet, first, last, xfIndex));
}

JsStatus JsWorkbook::clearRowStyle(uint32_t sheet, uint32_t first, uint32_t last) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_row_style(handle_, sheet, first, last));
}

emscripten::val JsWorkbook::rangeStyleBases(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                            uint32_t lastCol) const {
  emscripten::val o = emscripten::val::object();
  o.set("xfIndices", emscripten::val::array());
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    return o;
  }
  size_t len = 0;
  fm_status_t rc = fm_sheet_range_style_bases(handle_, sheet, firstRow, firstCol, lastRow, lastCol, nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_range_style_bases(handle_, sheet, firstRow, firstCol, lastRow, lastCol, xfs.data(), xfs.size(), &len);
  }
  if (rc != 0) {
    o.set("status", status_from_rc(rc));
    return o;
  }
  emscripten::val list = emscripten::val::array();
  for (size_t i = 0; i < xfs.size(); ++i) {
    list.call<void>("push", xfs[i]);
  }
  o.set("status", ok_status());
  o.set("xfIndices", list);
  return o;
}

JsStatus JsWorkbook::remapRangeStyle(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                     uint32_t lastCol, emscripten::val fromXfs, emscripten::val toXfs) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("remapRangeStyle");
  if (!reader.is_array(fromXfs, "fromXfs") || !reader.is_array(toXfs, "toXfs")) {
    return binding_error_status(kInvalidArgument, "remapRangeStyle: `fromXfs` and `toXfs` must be arrays of numbers");
  }
  const uint32_t count = reader.length(fromXfs, "fromXfs");
  if (reader.ok() && count != reader.length(toXfs, "toXfs")) {
    return binding_error_status(kInvalidArgument, "remapRangeStyle: `fromXfs` and `toXfs` must have the same length");
  }
  std::vector<uint32_t> from(count);
  std::vector<uint32_t> to(count);
  for (uint32_t i = 0; reader.ok() && i < count; ++i) {
    from[i] = reader.u32_value(reader.array_element(fromXfs, i, "fromXfs[]"), 0U, "fromXfs[]");
    to[i] = reader.u32_value(reader.array_element(toXfs, i, "toXfs[]"), 0U, "toXfs[]");
  }
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(
      fm_sheet_remap_range_style(handle_, sheet, firstRow, firstCol, lastRow, lastCol, from.data(), to.data(), count));
}

emscripten::val JsWorkbook::snapshotRangeStyle(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                               uint32_t lastCol) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    o.set("bytes", bytes_to_val(nullptr, 0));
    return o;
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc =
      fm_sheet_snapshot_range_style(handle_, sheet, firstRow, firstCol, lastRow, lastCol, &data, &len);
  o.set("status", status_from_rc(rc));
  o.set("bytes", bytes_to_val(rc == 0 ? data : nullptr, rc == 0 ? len : 0));
  fm_buffer_free(data);
  return o;
}

JsStatus JsWorkbook::restoreRangeStyle(uint32_t sheet, emscripten::val bytes) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return binding_error_status(kInvalidArgument, bytes_result.message.c_str());
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  return status_from_rc(fm_sheet_restore_range_style(handle_, sheet, data.data(), data.size()));
}

JsStatus JsWorkbook::clearRowHeight(uint32_t sheet, uint32_t row) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_row_height(handle_, sheet, row));
}

JsStatus JsWorkbook::clearColumnWidth(uint32_t sheet, uint32_t first, uint32_t last) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_column_width(handle_, sheet, first, last));
}

// ---- Sheet format defaults and point geometry ---------------------------

emscripten::val JsWorkbook::getSheetFormatDefaults(uint32_t sheet) const {
  fm_sheet_format_defaults d{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_format_defaults(handle_, sheet, &d) : kBindingInvalidHandle;
  if (rc != 0) {
    d = fm_sheet_format_defaults{};
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("defaultColWidth", d.default_col_width);
  o.set("defaultRowHeight", d.default_row_height);
  o.set("baseColWidth", d.base_col_width);
  o.set("hasDefaultColWidth", d.has_default_col_width != 0);
  o.set("hasDefaultRowHeight", d.has_default_row_height != 0);
  o.set("customHeight", d.custom_height != 0);
  o.set("zeroHeight", d.zero_height != 0);
  o.set("effectiveDefaultColWidth", d.effective_default_col_width);
  o.set("effectiveDefaultRowHeight", d.effective_default_row_height);
  return o;
}

JsStatus JsWorkbook::setSheetFormatDefaults(uint32_t sheet, emscripten::val defaults) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setSheetFormatDefaults");
  const emscripten::val default_col_width = reader.value(defaults, "defaultColWidth", "format.defaultColWidth");
  const emscripten::val default_row_height = reader.value(defaults, "defaultRowHeight", "format.defaultRowHeight");
  const emscripten::val base_col_width = reader.value(defaults, "baseColWidth", "format.baseColWidth");
  const emscripten::val has_default_col_width =
      reader.value(defaults, "hasDefaultColWidth", "format.hasDefaultColWidth");
  const emscripten::val has_default_row_height =
      reader.value(defaults, "hasDefaultRowHeight", "format.hasDefaultRowHeight");
  const bool default_col_width_present = js_value_present(default_col_width);
  const bool default_row_height_present = js_value_present(default_row_height);
  const emscripten::val custom_height = reader.value(defaults, "customHeight", "format.customHeight");
  const emscripten::val zero_height = reader.value(defaults, "zeroHeight", "format.zeroHeight");
  // An omitted flag keeps the stored one.
  fm_sheet_format_defaults d{};
  if (handle_ != nullptr && fm_sheet_get_format_defaults(handle_, sheet, &d) != 0) {
    d = fm_sheet_format_defaults{};
  }
  d.default_col_width = reader.number_value(default_col_width, 0.0, "format.defaultColWidth");
  d.default_row_height = reader.number_value(default_row_height, 0.0, "format.defaultRowHeight");
  d.base_col_width = reader.number_value(base_col_width, 8.0, "format.baseColWidth");
  d.has_default_col_width =
      reader.boolean_value(has_default_col_width, default_col_width_present, "format.hasDefaultColWidth") ? 1 : 0;
  d.has_default_row_height =
      reader.boolean_value(has_default_row_height, default_row_height_present, "format.hasDefaultRowHeight") ? 1 : 0;
  d.custom_height = reader.boolean_value(custom_height, d.custom_height != 0, "format.customHeight") ? 1 : 0;
  d.zero_height = reader.boolean_value(zero_height, d.zero_height != 0, "format.zeroHeight") ? 1 : 0;
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_set_format_defaults(handle_, sheet, &d));
}

emscripten::val JsWorkbook::getCellRectPt(uint32_t sheet, emscripten::val range, int32_t mode) const {
  fm_rect_pt r{};
  JsStatus status;
  if (handle_ == nullptr) {
    status = invalid_handle_status();
  } else {
    JsNarrowNumericReader reader("getCellRectPt");
    const fm_merge_range b = js_pull_range(range, &reader);
    if (!reader.ok()) {
      status = binding_error_status(kInvalidArgument, reader.message().c_str());
    } else {
      const fm_status_t rc =
          fm_sheet_cell_rect_pt(handle_, sheet, b.first_row, b.first_col, b.last_row, b.last_col, mode, &r);
      status = status_from_rc(rc);
      if (rc != 0) {
        r = fm_rect_pt{};
      }
    }
  }
  emscripten::val o = js_rect_pt(r);
  o.set("status", status);
  return o;
}

JsNumberResult JsWorkbook::getColumnWidthPt(uint32_t sheet, uint32_t col, int32_t mode) const {
  double pt = 0.0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_column_width_pt(handle_, sheet, col, mode, &pt) : kBindingInvalidHandle;
  return number_result(rc, pt);
}

JsNumberResult JsWorkbook::getRowHeightPt(uint32_t sheet, uint32_t row) const {
  double pt = 0.0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_row_height_pt(handle_, sheet, row, &pt) : kBindingInvalidHandle;
  return number_result(rc, pt);
}

emscripten::val JsWorkbook::getWidthModel(uint32_t sheet, int32_t mode) const {
  fm_width_model m{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_width_model(handle_, sheet, mode, &m) : kBindingInvalidHandle;
  if (rc != 0) {
    m = fm_width_model{};
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("pointsPerChar", m.points_per_char);
  o.set("paddingPt", m.padding_pt);
  o.set("normalFontSize", m.normal_font_size);
  o.set("calibrated", m.calibrated != 0);
  js_set_cstr(o, "normalFontName", m.normal_font_name);
  js_set_cstr(o, "platform", m.platform);
  return o;
}

JsNumberResult JsWorkbook::columnCharsToPt(uint32_t sheet, int32_t mode, double chars) const {
  double pt = 0.0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_column_chars_to_pt(handle_, sheet, mode, chars, &pt) : kBindingInvalidHandle;
  return number_result(rc, pt);
}

JsNumberResult JsWorkbook::columnPtToChars(uint32_t sheet, int32_t mode, double pt) const {
  double chars = 0.0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_column_pt_to_chars(handle_, sheet, mode, pt, &chars) : kBindingInvalidHandle;
  return number_result(rc, chars);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
