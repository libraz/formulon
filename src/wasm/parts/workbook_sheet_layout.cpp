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
  fm_column_layout_t entry = FM_STRUCT_INIT(fm_column_layout_t);
  if (const fm_status_t rc = fm_sheet_get_column_override_at(handle, sheet, index, &entry); rc != 0) {
    return rc;
  }
  *out = emscripten::val::object();
  out->set("first", entry.first);
  out->set("last", entry.last);
  js_set_present_or_null(*out, "width", entry.has_width, entry.width);
  out->set("hidden", entry.hidden != 0);
  out->set("outlineLevel", static_cast<int32_t>(entry.outline_level));
  out->set("collapsed", entry.collapsed != 0);
  js_set_present_or_null(*out, "styleXf", entry.has_style, entry.style_xf);
  return 0;
}

fm_status_t row_entry(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_row_layout_t entry = FM_STRUCT_INIT(fm_row_layout_t);
  if (const fm_status_t rc = fm_sheet_get_row_override_at(handle, sheet, index, &entry); rc != 0) {
    return rc;
  }
  *out = emscripten::val::object();
  out->set("row", entry.row);
  js_set_present_or_null(*out, "height", entry.has_height, entry.height);
  out->set("hidden", entry.hidden != 0);
  out->set("outlineLevel", static_cast<int32_t>(entry.outline_level));
  out->set("collapsed", entry.collapsed != 0);
  js_set_present_or_null(*out, "styleXf", entry.has_style, entry.style_xf);
  out->set("customHeight", entry.custom_height != 0);
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
  fm_paper_info_t paper = FM_STRUCT_INIT(fm_paper_info_t);
  fm_margins_pt_t margins = FM_STRUCT_INIT(fm_margins_pt_t);
  fm_rect_pt_t printable = FM_STRUCT_INIT(fm_rect_pt_t);
  double scale = 0.0;
  int32_t page_order = 0;
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
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
  result->set("printTitles", print_titles_to_val(titles));
  emscripten::val pages = emscripten::val::array();
  for (std::size_t i = 0; i < page_count; ++i) {
    fm_page_layout_t page = FM_STRUCT_INIT(fm_page_layout_t);
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
  result->set("rowBreakManual", horizontal_manual);
  emscripten::val vertical_manual = emscripten::val::array();
  for (std::size_t i = 0; i < vertical_count; ++i) {
    int32_t manual = 0;
    fm_pagination_is_column_break_manual(pagination, i, &manual);
    vertical_manual.call<void>("push", manual != 0);
  }
  result->set("columnBreakManual", vertical_manual);
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
  result.set("rowBreaks", horizontal_breaks);
  result.set("columnBreaks", vertical_breaks);
  std::size_t page_count = 0;
  fm_pagination_get_page_count(pagination, &page_count);
  result.set("pageCount", page_count);
  set_pagination_detail(&result, pagination);
  fm_pagination_destroy(pagination);
  return result;
}

JsSheetView JsWorkbook::getSheetView(uint32_t sheet) const {
  JsSheetView r;
  fm_sheet_view_t v = FM_STRUCT_INIT(fm_sheet_view_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_view(handle_, sheet, &v) : kBindingInvalidHandle)) {
    return r;
  }
  r.zoomScale = v.zoom_scale;
  r.freezeRows = v.freeze_rows;
  r.freezeCols = v.freeze_cols;
  r.visibility = v.visibility;
  r.showGridLines = v.show_grid_lines != 0;
  r.showRowColHeaders = v.show_row_col_headers != 0;
  r.showZeros = v.show_zeros != 0;
  r.rightToLeft = v.right_to_left != 0;
  r.tabSelected = v.tab_selected != 0;
  r.viewMode = v.view_mode;
  r.topLeftRow = v.top_left_row;
  r.topLeftCol = v.top_left_col;
  r.paneTopLeftRow = v.pane_top_left_row;
  r.paneTopLeftCol = v.pane_top_left_col;
  r.splitX = v.split_x;
  r.splitY = v.split_y;
  r.showFormulas = v.show_formulas != 0;
  r.showOutlineSymbols = v.show_outline_symbols != 0;
  return r;
}

JsSheetProtection JsWorkbook::getSheetProtection(uint32_t sheet) const {
  JsSheetProtection r;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_sheet_protection(handle_, sheet, &p) : kBindingInvalidHandle)) {
    return r;
  }
  r.enabled = p.enabled != 0;
  r.algorithmName = cstr_or_null(p.algorithm_name);
  r.hashValue = cstr_or_null(p.hash_value);
  r.saltValue = cstr_or_null(p.salt_value);
  r.spinCount = p.spin_count;
  r.legacyPassword = cstr_or_null(p.legacy_password);
  r.sheet = p.sheet != 0;
  r.objects = p.objects != 0;
  r.scenarios = p.scenarios != 0;
  r.formatCells = p.format_cells != 0;
  r.formatColumns = p.format_columns != 0;
  r.formatRows = p.format_rows != 0;
  r.insertColumns = p.insert_columns != 0;
  r.insertRows = p.insert_rows != 0;
  r.insertHyperlinks = p.insert_hyperlinks != 0;
  r.deleteColumns = p.delete_columns != 0;
  r.deleteRows = p.delete_rows != 0;
  r.selectLockedCells = p.select_locked_cells != 0;
  r.selectUnlockedCells = p.select_unlocked_cells != 0;
  r.sort = p.sort != 0;
  r.autoFilter = p.auto_filter != 0;
  r.pivotTables = p.pivot_tables != 0;
  return r;
}

void JsWorkbook::setSheetProtection(uint32_t sheet, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // A missing or `null` string is no stored value.
  JsNarrowNumericReader reader("Sheet.setSheetProtection");
  std::string algorithm_name;
  std::string hash_value;
  std::string salt_value;
  std::string legacy_password;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  p.algorithm_name = reader.optional_string(input, "algorithmName", algorithm_name, "algorithmName");
  p.hash_value = reader.optional_string(input, "hashValue", hash_value, "hashValue");
  p.salt_value = reader.optional_string(input, "saltValue", salt_value, "saltValue");
  p.spin_count = reader.u32(input, "spinCount", 0U, "spinCount");
  p.legacy_password = reader.optional_string(input, "legacyPassword", legacy_password, "legacyPassword");
  p.enabled = reader.boolean(input, "enabled", false, "enabled") ? 1 : 0;
  p.sheet = reader.boolean(input, "sheet", false, "sheet") ? 1 : 0;
  p.objects = reader.boolean(input, "objects", false, "objects") ? 1 : 0;
  p.scenarios = reader.boolean(input, "scenarios", false, "scenarios") ? 1 : 0;
  p.format_cells = reader.boolean(input, "formatCells", false, "formatCells") ? 1 : 0;
  p.format_columns = reader.boolean(input, "formatColumns", false, "formatColumns") ? 1 : 0;
  p.format_rows = reader.boolean(input, "formatRows", false, "formatRows") ? 1 : 0;
  p.insert_columns = reader.boolean(input, "insertColumns", false, "insertColumns") ? 1 : 0;
  p.insert_rows = reader.boolean(input, "insertRows", false, "insertRows") ? 1 : 0;
  p.insert_hyperlinks = reader.boolean(input, "insertHyperlinks", false, "insertHyperlinks") ? 1 : 0;
  p.delete_columns = reader.boolean(input, "deleteColumns", false, "deleteColumns") ? 1 : 0;
  p.delete_rows = reader.boolean(input, "deleteRows", false, "deleteRows") ? 1 : 0;
  p.select_locked_cells = reader.boolean(input, "selectLockedCells", false, "selectLockedCells") ? 1 : 0;
  p.select_unlocked_cells = reader.boolean(input, "selectUnlockedCells", false, "selectUnlockedCells") ? 1 : 0;
  p.sort = reader.boolean(input, "sort", false, "sort") ? 1 : 0;
  p.auto_filter = reader.boolean(input, "autoFilter", false, "autoFilter") ? 1 : 0;
  p.pivot_tables = reader.boolean(input, "pivotTables", false, "pivotTables") ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  fm_status_t rc = fm_sheet_set_sheet_protection(handle_, sheet, &p);
  succeeded(rc);
}

void JsWorkbook::setSheetZoom(uint32_t sheet, uint32_t zoomScale) {
  succeeded(handle_ != nullptr ? fm_sheet_set_zoom(handle_, sheet, zoomScale) : kBindingInvalidHandle);
}

void JsWorkbook::setSheetFreeze(uint32_t sheet, uint32_t freezeRows, uint32_t freezeCols) {
  succeeded(handle_ != nullptr ? fm_sheet_set_freeze(handle_, sheet, freezeRows, freezeCols) : kBindingInvalidHandle);
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

void JsWorkbook::setSheetViewMode(uint32_t sheet, int32_t mode) {
  succeeded(handle_ != nullptr ? fm_sheet_set_view_mode(handle_, sheet, mode) : kBindingInvalidHandle);
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
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_format_defaults(handle_, sheet, &d) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_present_or_null(o, "defaultColWidth", d.has_default_col_width, d.default_col_width);
  js_set_present_or_null(o, "defaultRowHeight", d.has_default_row_height, d.default_row_height);
  o.set("baseColWidth", d.base_col_width);
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
  // A patch: an omitted key keeps the stored value, `null` clears a nullable one.
  JsNarrowNumericReader reader("Sheet.setFormatDefaults");
  if (!reader.require_object(defaults, "defaults")) {
    return;
  }
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  if (!succeeded(fm_sheet_get_format_defaults(handle_, sheet, &d))) {
    return;
  }
  const emscripten::val col_width = reader.value(defaults, "defaultColWidth", "defaultColWidth");
  switch (js_patch_key(reader, col_width, "defaultColWidth", true)) {
    case JsPatchKey::kClear:
      d.has_default_col_width = 0;
      d.default_col_width = 0.0;
      break;
    case JsPatchKey::kSet:
      d.has_default_col_width = 1;
      d.default_col_width = reader.number_value(col_width, 0.0, "defaultColWidth");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  const emscripten::val row_height = reader.value(defaults, "defaultRowHeight", "defaultRowHeight");
  switch (js_patch_key(reader, row_height, "defaultRowHeight", true)) {
    case JsPatchKey::kClear:
      d.has_default_row_height = 0;
      d.default_row_height = 0.0;
      break;
    case JsPatchKey::kSet:
      d.has_default_row_height = 1;
      d.default_row_height = reader.number_value(row_height, 0.0, "defaultRowHeight");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  const emscripten::val base = reader.value(defaults, "baseColWidth", "baseColWidth");
  if (js_patch_key(reader, base, "baseColWidth", false) == JsPatchKey::kSet) {
    d.base_col_width = reader.number_value(base, d.base_col_width, "baseColWidth");
  }
  const emscripten::val custom_height = reader.value(defaults, "customHeight", "customHeight");
  if (js_patch_key(reader, custom_height, "customHeight", false) == JsPatchKey::kSet) {
    d.custom_height = reader.boolean_value(custom_height, false) ? 1 : 0;
  }
  const emscripten::val zero_height = reader.value(defaults, "zeroHeight", "zeroHeight");
  if (js_patch_key(reader, zero_height, "zeroHeight", false) == JsPatchKey::kSet) {
    d.zero_height = reader.boolean_value(zero_height, false) ? 1 : 0;
  }
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
  fm_rect_pt_t r = FM_STRUCT_INIT(fm_rect_pt_t);
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
  fm_width_model_t m = FM_STRUCT_INIT(fm_width_model_t);
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
