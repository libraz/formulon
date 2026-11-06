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

// Builds an enumeration array; `undefined` after recording the failure.
emscripten::val entry_list(const fm_workbook_t* handle, uint32_t sheet, JsListCount count_getter,
                           JsListItem item_getter) {
  emscripten::val list = emscripten::val::array();
  if (!succeeded(js_collect_list(handle, sheet, count_getter, item_getter, list))) {
    return emscripten::val::undefined();
  }
  return list;
}

fm_status_t column_entry(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_column_layout_t entry{};
  if (const fm_status_t rc = fm_sheet_get_column_override_at(handle, sheet, index, &entry); rc != 0) {
    return rc;
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
  return 0;
}

fm_status_t row_entry(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_row_layout_t entry{};
  if (const fm_status_t rc = fm_sheet_get_row_override_at(handle, sheet, index, &entry); rc != 0) {
    return rc;
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
  return 0;
}

uint8_t clamp_outline_level(uint32_t level) {
  return static_cast<uint8_t>(level > 255U ? 255U : level);
}

emscripten::val js_rect_pt(const fm_rect_pt_t& r) {
  emscripten::val o = emscripten::val::object();
  o.set("x", r.x);
  o.set("y", r.y);
  o.set("width", r.width);
  o.set("height", r.height);
  return o;
}

/// Writes the pagination detail keys onto `result`.
void set_pagination_detail(emscripten::val* result, const fm_pagination_t* pagination) {
  fm_paper_info_t paper{};
  fm_margins_pt_t margins{};
  fm_rect_pt_t printable{};
  double scale = 0.0;
  int32_t page_order = 0;
  fm_print_titles_t titles{};
  std::size_t page_count = 0;
  std::size_t horizontal_count = 0;
  std::size_t vertical_count = 0;
  fm_pagination_get_paper(pagination, &paper);
  fm_pagination_get_margins(pagination, &margins);
  fm_pagination_get_printable(pagination, &printable);
  fm_pagination_get_scale(pagination, &scale);
  fm_pagination_get_page_order(pagination, &page_order);
  fm_pagination_get_print_titles(pagination, &titles);
  fm_pagination_get_page_count(pagination, &page_count);
  fm_pagination_get_row_break_count(pagination, &horizontal_count);
  fm_pagination_get_column_break_count(pagination, &vertical_count);
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
    fm_page_layout_t page{};
    if (fm_pagination_get_page_at(pagination, i, &page) != 0) {
      continue;
    }
    emscripten::val item = emscripten::val::object();
    item.set("areaIndex", page.area_index);
    item.set("rect", cell_rect_to_val(page.rect));
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
    fm_pagination_is_row_break_manual(pagination, i, &manual);
    horizontal_manual.call<void>("push", manual != 0);
  }
  result->set("horizontalBreakManual", horizontal_manual);
  emscripten::val vertical_manual = emscripten::val::array();
  for (std::size_t i = 0; i < vertical_count; ++i) {
    int32_t manual = 0;
    fm_pagination_is_column_break_manual(pagination, i, &manual);
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
  fm_pagination_t* pagination = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_pagination(handle_, sheet, &pagination) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  std::size_t area_count = 0;
  fm_pagination_get_print_area_count(pagination, &area_count);
  for (std::size_t i = 0; i < area_count; ++i) {
    fm_cell_rect_t range{};
    if (fm_pagination_get_print_area_at(pagination, i, &range) != 0) {
      continue;
    }
    print_area.set(i, cell_rect_to_val(range));
  }
  std::size_t horizontal_count = 0;
  fm_pagination_get_row_break_count(pagination, &horizontal_count);
  for (std::size_t i = 0; i < horizontal_count; ++i) {
    uint32_t row = 0;
    if (fm_pagination_get_row_break_at(pagination, i, &row) == 0) {
      horizontal_breaks.set(i, row);
    }
  }
  std::size_t vertical_count = 0;
  fm_pagination_get_column_break_count(pagination, &vertical_count);
  for (std::size_t i = 0; i < vertical_count; ++i) {
    uint32_t col = 0;
    if (fm_pagination_get_column_break_at(pagination, i, &col) == 0) {
      vertical_breaks.set(i, col);
    }
  }
  result.set("printArea", print_area);
  result.set("horizontalBreaks", horizontal_breaks);
  result.set("verticalBreaks", vertical_breaks);
  std::size_t page_count = 0;
  fm_pagination_get_page_count(pagination, &page_count);
  result.set("pageCount", page_count);
  set_pagination_detail(&result, pagination);
  fm_pagination_destroy(pagination);
  return result;
}

JsSheetView JsWorkbook::getSheetView(uint32_t sheet) const {
  JsSheetView r;
  fm_sheet_view_t v{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_view(handle_, sheet, &v) : kBindingInvalidHandle)) {
    return r;
  }
  r.zoomScale = v.zoom_scale;
  r.freezeRows = v.freeze_rows;
  r.freezeCols = v.freeze_cols;
  r.tabHidden = v.tab_hidden;
  r.visibility = v.visibility;
  r.showGridLines = v.show_grid_lines;
  r.showRowColHeaders = v.show_row_col_headers;
  r.showZeros = v.show_zeros;
  r.rightToLeft = v.right_to_left;
  r.tabSelected = v.tab_selected;
  r.viewMode = string_from_cstr(v.view_mode);
  r.topLeftRow = v.top_left_row;
  r.topLeftCol = v.top_left_col;
  r.paneTopLeftRow = v.pane_top_left_row;
  r.paneTopLeftCol = v.pane_top_left_col;
  r.splitX = v.split_x;
  r.splitY = v.split_y;
  r.showFormulas = v.show_formulas;
  r.showOutlineSymbols = v.show_outline_symbols;
  return r;
}

JsSheetProtection JsWorkbook::getSheetProtection(uint32_t sheet) const {
  JsSheetProtection r;
  fm_sheet_protection_t p{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_sheet_protection(handle_, sheet, &p) : kBindingInvalidHandle)) {
    return r;
  }
  r.enabled = p.enabled;
  r.algorithmName = string_from_cstr(p.algorithm_name);
  r.hashValue = string_from_cstr(p.hash_value);
  r.saltValue = string_from_cstr(p.salt_value);
  r.spinCount = p.spin_count;
  r.legacyPassword = string_from_cstr(p.legacy_password);
  r.sheet = p.sheet;
  r.objects = p.objects;
  r.scenarios = p.scenarios;
  r.formatCells = p.format_cells;
  r.formatColumns = p.format_columns;
  r.formatRows = p.format_rows;
  r.insertColumns = p.insert_columns;
  r.insertRows = p.insert_rows;
  r.insertHyperlinks = p.insert_hyperlinks;
  r.deleteColumns = p.delete_columns;
  r.deleteRows = p.delete_rows;
  r.selectLockedCells = p.select_locked_cells;
  r.selectUnlockedCells = p.select_unlocked_cells;
  r.sort = p.sort;
  r.autoFilter = p.auto_filter;
  r.pivotTables = p.pivot_tables;
  return r;
}

void JsWorkbook::setSheetProtection(uint32_t sheet, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
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
    return;
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
  fm_status_t rc = fm_sheet_set_sheet_protection(handle_, sheet, &p);
  succeeded(rc);
}

void JsWorkbook::setSheetZoom(uint32_t sheet, uint32_t zoomScale) {
  succeeded(handle_ != nullptr ? fm_sheet_set_zoom(handle_, sheet, zoomScale) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols) {
  succeeded(handle_ != nullptr ? fm_sheet_set_freeze(handle_, sheet, freezeRows, freezeCols) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetTabHidden(uint32_t sheet, bool hidden) {
  succeeded(handle_ != nullptr ? fm_sheet_set_tab_hidden(handle_, sheet, hidden ? 1 : 0) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetVisibility(uint32_t sheet, int32_t visibility) {
  succeeded(handle_ != nullptr ? fm_sheet_set_visibility(handle_, sheet, visibility) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetShowGridLines(uint32_t sheet, bool show) {
  succeeded(handle_ != nullptr ? fm_sheet_set_show_grid_lines(handle_, sheet, show ? 1 : 0) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetShowRowColHeaders(uint32_t sheet, bool show) {
  succeeded(handle_ != nullptr ? fm_sheet_set_show_row_col_headers(handle_, sheet, show ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setSheetShowZeros(uint32_t sheet, bool show) {
  succeeded(handle_ != nullptr ? fm_sheet_set_show_zeros(handle_, sheet, show ? 1 : 0) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetShowFormulas(uint32_t sheet, bool show) {
  succeeded(handle_ != nullptr ? fm_sheet_set_show_formulas(handle_, sheet, show ? 1 : 0) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetShowOutlineSymbols(uint32_t sheet, bool show) {
  succeeded(handle_ != nullptr ? fm_sheet_set_show_outline_symbols(handle_, sheet, show ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setSheetRightToLeft(uint32_t sheet, bool rightToLeft) {
  succeeded(handle_ != nullptr ? fm_sheet_set_right_to_left(handle_, sheet, rightToLeft ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setSheetTopLeft(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_set_top_left(handle_, sheet, row, col) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetPaneTopLeft(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_set_pane_top_left(handle_, sheet, row, col) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetSplit(uint32_t sheet, double xTwips, double yTwips) {
  succeeded(handle_ != nullptr ? fm_sheet_set_split(handle_, sheet, xTwips, yTwips) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getSheetSelection(uint32_t sheet) const {
  uint32_t row = 0;
  uint32_t col = 0;
  size_t total = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_selection(handle_, sheet, &row, &col, nullptr, 0, &total)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  std::vector<fm_cell_rect_t> found(total);
  if (!succeeded(fm_sheet_get_selection(handle_, sheet, &row, &col, found.data(), found.size(), &total))) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("activeRow", row);
  o.set("activeCol", col);
  o.set("ranges", cell_rects_to_val(found.data(), found.size()));
  return o;
}

void JsWorkbook::setSheetSelection(uint32_t sheet, emscripten::val selection) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setSelection");
  const uint32_t row = reader.u32(selection, "activeRow", 0U, "selection.activeRow");
  const uint32_t col = reader.u32(selection, "activeCol", 0U, "selection.activeCol");
  const std::vector<fm_cell_rect_t> ranges = js_pull_ranges(selection, "ranges", &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_selection(handle_, sheet, row, col, ranges.data(), ranges.size()));
}

void JsWorkbook::setSheetTabSelected(uint32_t sheet, bool selected) {
  succeeded(handle_ != nullptr ? fm_sheet_set_tab_selected(handle_, sheet, selected ? 1 : 0) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetViewMode(uint32_t sheet, std::string mode) {
  succeeded(handle_ != nullptr ? fm_sheet_set_view_mode(handle_, sheet, mode.c_str()) : kBindingInvalidHandle);
}

// ---- Column layout overrides --------------------------------------------

emscripten::val JsWorkbook::getSheetColumns(uint32_t sheet) const {
  return entry_list(handle_, sheet, &fm_sheet_get_column_override_count, &column_entry);
}

void JsWorkbook::setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_width(handle_, sheet, first, last, width) : kBindingInvalidHandle);
}

void JsWorkbook::setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_outline(handle_, sheet, first, last, clamp_outline_level(level))
                               : kBindingInvalidHandle);
}

void JsWorkbook::setColumnCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_collapsed(handle_, sheet, first, last, collapsed ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setColumnStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_style(handle_, sheet, first, last, xfIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::clearColumnStyle(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_column_style(handle_, sheet, first, last) : kBindingInvalidHandle);
}

// ---- Row layout overrides ----------------------------------------------

emscripten::val JsWorkbook::getSheetRowOverrides(uint32_t sheet) const {
  return entry_list(handle_, sheet, &fm_sheet_get_row_override_count, &row_entry);
}

void JsWorkbook::setRowHeight(uint32_t sheet, uint32_t first, uint32_t last, double height) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_height(handle_, sheet, first, last, height) : kBindingInvalidHandle);
}

void JsWorkbook::setRowAutoHeight(uint32_t sheet, uint32_t first, uint32_t last, double height) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_auto_height(handle_, sheet, first, last, height)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setRowHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_hidden(handle_, sheet, first, last, hidden ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setRowOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_outline(handle_, sheet, first, last, clamp_outline_level(level))
                               : kBindingInvalidHandle);
}

void JsWorkbook::setRowCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_collapsed(handle_, sheet, first, last, collapsed ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setRowStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_style(handle_, sheet, first, last, xfIndex) : kBindingInvalidHandle);
}

void JsWorkbook::clearRowStyle(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_row_style(handle_, sheet, first, last) : kBindingInvalidHandle);
}

namespace {

emscripten::val range_style_bases_result(fm_workbook_t* handle, uint32_t sheet, const fm_cell_rect_t& rect,
                                         const std::vector<uint32_t>& rows, const std::vector<uint32_t>& cols) {
  size_t len = 0;
  fm_status_t rc = fm_sheet_get_range_style_bases_excluding(handle, sheet, &rect, rows.data(), cols.data(), rows.size(),
                                                            nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_range_style_bases_excluding(handle, sheet, &rect, rows.data(), cols.data(), rows.size(),
                                                  xfs.data(), xfs.size(), &len);
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val list = emscripten::val::array();
  for (size_t i = 0; i < xfs.size(); ++i) {
    list.call<void>("push", xfs[i]);
  }
  return list;
}

}  // namespace

emscripten::val JsWorkbook::rangeStyleBases(uint32_t sheet, emscripten::val range) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("getRangeStyleBases");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  return range_style_bases_result(handle_, sheet, rect, {}, {});
}

emscripten::val JsWorkbook::rangeStyleBasesExcluding(uint32_t sheet, emscripten::val range,
                                                     emscripten::val excludedRows, emscripten::val excludedCols) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("getRangeStyleBasesExcluding");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  if (!reader.is_array(excludedRows, "excludedRows") || !reader.is_array(excludedCols, "excludedCols")) {
    record_type_error(
        "getRangeStyleBasesExcluding: `excludedRows` and `excludedCols` must be arrays "
        "of numbers");
    return emscripten::val::undefined();
  }
  const uint32_t count = reader.length(excludedRows, "excludedRows");
  if (reader.ok() && count != reader.length(excludedCols, "excludedCols")) {
    record_type_error("getRangeStyleBasesExcluding: `excludedRows` and `excludedCols` must have the same length");
    return emscripten::val::undefined();
  }
  std::vector<uint32_t> rows(count);
  std::vector<uint32_t> cols(count);
  for (uint32_t i = 0; reader.ok() && i < count; ++i) {
    rows[i] = reader.u32_value(reader.array_element(excludedRows, i, "excludedRows[]"), 0U, "excludedRows[]");
    cols[i] = reader.u32_value(reader.array_element(excludedCols, i, "excludedCols[]"), 0U, "excludedCols[]");
  }
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  return range_style_bases_result(handle_, sheet, rect, rows, cols);
}

void JsWorkbook::remapRangeStyle(uint32_t sheet, emscripten::val range, emscripten::val fromXfs,
                                 emscripten::val toXfs) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("remapRangeStyle");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  if (!reader.is_array(fromXfs, "fromXfs") || !reader.is_array(toXfs, "toXfs")) {
    record_type_error("remapRangeStyle: `fromXfs` and `toXfs` must be arrays of numbers");
    return;
  }
  const uint32_t count = reader.length(fromXfs, "fromXfs");
  if (reader.ok() && count != reader.length(toXfs, "toXfs")) {
    record_type_error("remapRangeStyle: `fromXfs` and `toXfs` must have the same length");
    return;
  }
  std::vector<uint32_t> from(count);
  std::vector<uint32_t> to(count);
  for (uint32_t i = 0; reader.ok() && i < count; ++i) {
    from[i] = reader.u32_value(reader.array_element(fromXfs, i, "fromXfs[]"), 0U, "fromXfs[]");
    to[i] = reader.u32_value(reader.array_element(toXfs, i, "toXfs[]"), 0U, "toXfs[]");
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_remap_range_style(handle_, sheet, &rect, from.data(), to.data(), count));
}

emscripten::val JsWorkbook::snapshotRangeStyle(uint32_t sheet, emscripten::val range) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("snapshotRangeStyle");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_range_style(handle_, sheet, &rect, &data, &len);
  return snapshot_bytes(rc, data, len);
}

void JsWorkbook::restoreRangeStyle(uint32_t sheet, emscripten::val bytes) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return;
  }
  succeeded(fm_sheet_restore_range_style(handle_, sheet, data.data(), data.size()));
}

void JsWorkbook::clearRowHeight(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_row_height(handle_, sheet, first, last) : kBindingInvalidHandle);
}

void JsWorkbook::clearColumnWidth(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_column_width(handle_, sheet, first, last) : kBindingInvalidHandle);
}

// ---- Sheet format defaults and point geometry ---------------------------

emscripten::val JsWorkbook::getSheetFormatDefaults(uint32_t sheet) const {
  fm_sheet_format_defaults_t d{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_format_defaults(handle_, sheet, &d) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
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

void JsWorkbook::setSheetFormatDefaults(uint32_t sheet, emscripten::val defaults) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setFormatDefaults");
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
  fm_sheet_format_defaults_t d{};
  if (handle_ != nullptr && fm_sheet_get_format_defaults(handle_, sheet, &d) != 0) {
    d = fm_sheet_format_defaults_t{};
  }
  d.default_col_width = reader.number_value(default_col_width, 0.0, "format.defaultColWidth");
  d.default_row_height = reader.number_value(default_row_height, 0.0, "format.defaultRowHeight");
  d.base_col_width = reader.number_value(base_col_width, 8.0, "format.baseColWidth");
  d.has_default_col_width = reader.boolean_value(has_default_col_width, default_col_width_present) ? 1 : 0;
  d.has_default_row_height = reader.boolean_value(has_default_row_height, default_row_height_present) ? 1 : 0;
  d.custom_height = reader.boolean_value(custom_height, d.custom_height != 0) ? 1 : 0;
  d.zero_height = reader.boolean_value(zero_height, d.zero_height != 0) ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_format_defaults(handle_, sheet, &d));
}

emscripten::val JsWorkbook::getCellRectPt(uint32_t sheet, emscripten::val range, int32_t mode) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("getCellRectPt");
  const fm_cell_rect_t b = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  fm_rect_pt_t r{};
  if (!succeeded(fm_sheet_get_cell_rect_pt(handle_, sheet, &b, mode, &r))) {
    return emscripten::val::undefined();
  }
  return js_rect_pt(r);
}

double JsWorkbook::getColumnWidthPt(uint32_t sheet, uint32_t col, int32_t mode) const {
  double pt = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_column_width_pt(handle_, sheet, col, mode, &pt)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return pt;
}

double JsWorkbook::getRowHeightPt(uint32_t sheet, uint32_t row) const {
  double pt = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_row_height_pt(handle_, sheet, row, &pt) : kBindingInvalidHandle)) {
    return 0.0;
  }
  return pt;
}

emscripten::val JsWorkbook::getWidthModel(uint32_t sheet, int32_t mode) const {
  fm_width_model_t m{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_width_model(handle_, sheet, mode, &m) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("pointsPerChar", m.points_per_char);
  o.set("paddingPt", m.padding_pt);
  o.set("normalFontSize", m.normal_font_size);
  o.set("calibrated", m.calibrated != 0);
  js_set_cstr(o, "normalFontName", m.normal_font_name);
  js_set_cstr(o, "platform", m.platform);
  return o;
}

double JsWorkbook::columnCharsToPt(uint32_t sheet, int32_t mode, double chars) const {
  double pt = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_convert_column_chars_to_pt(handle_, sheet, mode, chars, &pt)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return pt;
}

double JsWorkbook::columnPtToChars(uint32_t sheet, int32_t mode, double pt) const {
  double chars = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_convert_column_pt_to_chars(handle_, sheet, mode, pt, &chars)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return chars;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
