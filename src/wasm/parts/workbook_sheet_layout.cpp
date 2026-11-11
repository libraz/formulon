//
// JsWorkbook sheet-view and per-sheet layout surface: zoom / freeze /
// tabHidden, the `<sheetProtection>` get/set bridge, and the column /
// row layout override iterators and mutators.

#include <cstddef>
#include <cstdint>
#include <cstring>
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
  fm_column_layout_info_t entry = FM_STRUCT_INIT(fm_column_layout_info_t);
  if (const fm_status_t rc = fm_sheet_get_column_layout_at(handle, sheet, index, &entry); rc != 0) {
    return rc;
  }
  *out = emscripten::val::object();
  out->set("first", entry.first);
  out->set("last", entry.last);
  js_set_present_or_null(*out, "widthChars", entry.has_width_chars, entry.width_chars);
  out->set("hidden", entry.hidden != 0);
  out->set("outlineLevel", static_cast<int32_t>(entry.outline_level));
  out->set("collapsed", entry.collapsed != 0);
  js_set_present_or_null(*out, "xfIndex", entry.has_xf_index, entry.xf_index);
  return 0;
}

fm_status_t row_entry(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_row_layout_info_t entry = FM_STRUCT_INIT(fm_row_layout_info_t);
  if (const fm_status_t rc = fm_sheet_get_row_layout_at(handle, sheet, index, &entry); rc != 0) {
    return rc;
  }
  *out = emscripten::val::object();
  out->set("row", entry.row);
  js_set_present_or_null(*out, "heightPt", entry.has_height_pt, entry.height_pt);
  out->set("hidden", entry.hidden != 0);
  out->set("outlineLevel", static_cast<int32_t>(entry.outline_level));
  out->set("collapsed", entry.collapsed != 0);
  js_set_present_or_null(*out, "xfIndex", entry.has_xf_index, entry.xf_index);
  out->set("customHeight", entry.custom_height != 0);
  return 0;
}

emscripten::val js_rect_pt(const fm_rect_pt_info_t& r) {
  emscripten::val o = emscripten::val::object();
  o.set("x", r.x);
  o.set("y", r.y);
  o.set("widthPt", r.width_pt);
  o.set("heightPt", r.height_pt);
  return o;
}

/// Writes the pagination detail keys onto `result`.
void set_pagination_detail(emscripten::val* result, const fm_pagination_t* pagination) {
  fm_paper_info_t paper = FM_STRUCT_INIT(fm_paper_info_t);
  fm_margins_pt_info_t margins = FM_STRUCT_INIT(fm_margins_pt_info_t);
  fm_rect_pt_info_t printable = FM_STRUCT_INIT(fm_rect_pt_info_t);
  double scale = 0.0;
  int32_t page_order = 0;
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  std::size_t page_count = 0;
  fm_pagination_get_paper(pagination, &paper);
  fm_pagination_get_margins(pagination, &margins);
  fm_pagination_get_printable(pagination, &printable);
  fm_pagination_get_scale(pagination, &scale);
  fm_pagination_get_page_order(pagination, &page_order);
  fm_pagination_get_print_titles(pagination, &titles);
  fm_pagination_get_page_count(pagination, &page_count);
  emscripten::val paper_val = emscripten::val::object();
  paper_val.set("widthPt", paper.width_pt);
  paper_val.set("heightPt", paper.height_pt);
  paper_val.set("landscape", paper.landscape != 0);
  paper_val.set("known", paper.known != 0);
  result->set("paper", paper_val);
  emscripten::val margins_val = emscripten::val::object();
  margins_val.set("leftIn", margins.left_in);
  margins_val.set("rightIn", margins.right_in);
  margins_val.set("topIn", margins.top_in);
  margins_val.set("bottomIn", margins.bottom_in);
  margins_val.set("headerIn", margins.header_in);
  margins_val.set("footerIn", margins.footer_in);
  result->set("margins", margins_val);
  result->set("printable", js_rect_pt(printable));
  result->set("scalePct", scale);
  result->set("pageOrder", page_order);
  result->set("printTitles", print_titles_to_val(titles));
  emscripten::val pages = emscripten::val::array();
  for (std::size_t i = 0; i < page_count; ++i) {
    fm_page_layout_info_t page = FM_STRUCT_INIT(fm_page_layout_info_t);
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
}

/// The breaks in effect as `RowBreak` / `ColumnBreak` records, in ascending position. A break the page walk
/// computed covers the whole sheet (`first` / `last` of `0`).
emscripten::val pagination_breaks(const fm_pagination_t* pagination, bool rows) {
  emscripten::val out = emscripten::val::array();
  size_t count = 0;
  if (rows) {
    fm_pagination_get_row_break_count(pagination, &count);
  } else {
    fm_pagination_get_column_break_count(pagination, &count);
  }
  for (size_t i = 0; i < count; ++i) {
    emscripten::val entry = emscripten::val::object();
    if (rows) {
      fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
      if (fm_pagination_get_row_break_at(pagination, i, &brk) != 0) {
        continue;
      }
      entry.set("row", brk.row);
      entry.set("first", brk.first);
      entry.set("last", brk.last);
      entry.set("manual", brk.manual == FM_TRISTATE_TRUE);
    } else {
      fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
      if (fm_pagination_get_column_break_at(pagination, i, &brk) != 0) {
        continue;
      }
      entry.set("col", brk.col);
      entry.set("first", brk.first);
      entry.set("last", brk.last);
      entry.set("manual", brk.manual == FM_TRISTATE_TRUE);
    }
    out.call<void>("push", entry);
  }
  return out;
}

}  // namespace

emscripten::val JsWorkbook::paginate(uint32_t sheet) const {
  emscripten::val result = emscripten::val::object();
  emscripten::val print_area = emscripten::val::array();
  fm_pagination_t* pagination = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_sheet_paginate(handle_, sheet, &pagination) : kBindingInvalidHandle)) {
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
  result.set("printArea", print_area);
  result.set("rowBreaks", pagination_breaks(pagination, /*rows=*/true));
  result.set("columnBreaks", pagination_breaks(pagination, /*rows=*/false));
  std::size_t page_count = 0;
  fm_pagination_get_page_count(pagination, &page_count);
  result.set("pageCount", page_count);
  set_pagination_detail(&result, pagination);
  fm_pagination_destroy(pagination);
  return result;
}

int32_t JsWorkbook::getVisibility(uint32_t sheet) const {
  int32_t visibility = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_visibility(handle_, sheet, &visibility) : kBindingInvalidHandle)) {
    return 0;
  }
  return visibility;
}

void JsWorkbook::setVisibility(uint32_t sheet, int32_t visibility) {
  succeeded(handle_ != nullptr ? fm_sheet_set_visibility(handle_, sheet, visibility) : kBindingInvalidHandle);
}

JsSheetView JsWorkbook::getSheetView(uint32_t sheet) const {
  JsSheetView r;
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_view(handle_, sheet, &v) : kBindingInvalidHandle)) {
    return r;
  }
  r.zoomPct = v.zoom_pct;
  r.freezeRows = v.freeze_rows;
  r.freezeCols = v.freeze_cols;
  r.showGridLines = v.show_grid_lines != 0;
  r.showRowColumnHeaders = v.show_row_column_headers != 0;
  r.showZeros = v.show_zeros != 0;
  r.rightToLeft = v.right_to_left != 0;
  r.tabSelected = v.tab_selected != 0;
  r.viewMode = v.view_mode;
  r.topLeftRow = v.top_left_row;
  r.topLeftCol = v.top_left_col;
  r.paneTopLeftRow = v.pane_top_left_row;
  r.paneTopLeftCol = v.pane_top_left_col;
  r.splitXTwips = v.split_x_twips;
  r.splitYTwips = v.split_y_twips;
  r.showFormulas = v.show_formulas != 0;
  r.showOutlineSymbols = v.show_outline_symbols != 0;
  return r;
}

namespace {

/// One `allow*` flag of the sheet protection record: its JS key, its slot in the record and the slots of its
/// patch companion. The two select flags are tristates in the record (an omitted one allows).
struct ProtectionFlag {
  const char* key;
  std::size_t record;
  std::size_t patch_op;
  std::size_t patch_value;
  bool tristate;
};

// clang-format off
#define PROTECTION_FLAG(key, member, tristate) \
  {key, offsetof(fm_sheet_protection_t, member), offsetof(fm_sheet_protection_patch_t, member##_op), \
   offsetof(fm_sheet_protection_patch_t, member), tristate}
constexpr ProtectionFlag kProtectionFlags[] = {
    PROTECTION_FLAG("allowEditObjects", allow_edit_objects, false),
    PROTECTION_FLAG("allowEditScenarios", allow_edit_scenarios, false),
    PROTECTION_FLAG("allowFormatCells", allow_format_cells, false),
    PROTECTION_FLAG("allowFormatColumns", allow_format_columns, false),
    PROTECTION_FLAG("allowFormatRows", allow_format_rows, false),
    PROTECTION_FLAG("allowInsertColumns", allow_insert_columns, false),
    PROTECTION_FLAG("allowInsertRows", allow_insert_rows, false),
    PROTECTION_FLAG("allowInsertHyperlinks", allow_insert_hyperlinks, false),
    PROTECTION_FLAG("allowDeleteColumns", allow_delete_columns, false),
    PROTECTION_FLAG("allowDeleteRows", allow_delete_rows, false),
    PROTECTION_FLAG("allowSelectLockedCells", allow_select_locked_cells, true),
    PROTECTION_FLAG("allowSelectUnlockedCells", allow_select_unlocked_cells, true),
    PROTECTION_FLAG("allowSort", allow_sort, false),
    PROTECTION_FLAG("allowAutoFilter", allow_auto_filter, false),
    PROTECTION_FLAG("allowPivotTables", allow_pivot_tables, false),
};
#undef PROTECTION_FLAG
// clang-format on

template <class T>
T field_at(const void* base, std::size_t offset) {
  T value;
  std::memcpy(&value, static_cast<const char*>(base) + offset, sizeof(T));
  return value;
}

template <class T>
void set_field_at(void* base, std::size_t offset, T value) {
  std::memcpy(static_cast<char*>(base) + offset, &value, sizeof(T));
}

}  // namespace

emscripten::val JsWorkbook::getSheetProtection(uint32_t sheet) const {
  int32_t present = 0;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_sheet_protection(handle_, sheet, &present, &p)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  if (present == 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_or_null_fields(o, {{"algorithmName", p.algorithm_name},
                                 {"hashValue", p.hash_value},
                                 {"saltValue", p.salt_value},
                                 {"legacyPasswordHash", p.legacy_password_hash}});
  o.set("spinCount", p.spin_count);
  for (const ProtectionFlag& flag : kProtectionFlags) {
    const int32_t value = field_at<int32_t>(&p, flag.record);
    o.set(flag.key, flag.tristate ? value == FM_TRISTATE_TRUE : value != 0);
  }
  return o;
}

// `null` unprotects the sheet; a record protects it with every omitted key at Excel's dialog default.
void JsWorkbook::setSheetProtection(uint32_t sheet, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (input.isNull()) {
    succeeded(fm_sheet_set_sheet_protection(handle_, sheet, nullptr));
    return;
  }
  // A missing or `null` string is no stored value.
  JsNarrowNumericReader reader("Sheet.setSheetProtection");
  std::string algorithm_name;
  std::string hash_value;
  std::string salt_value;
  std::string legacy_password_hash;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  p.algorithm_name = reader.optional_string(input, "algorithmName", algorithm_name, "algorithmName");
  p.hash_value = reader.optional_string(input, "hashValue", hash_value, "hashValue");
  p.salt_value = reader.optional_string(input, "saltValue", salt_value, "saltValue");
  p.spin_count = reader.u32(input, "spinCount", 0U, "spinCount");
  p.legacy_password_hash =
      reader.optional_string(input, "legacyPasswordHash", legacy_password_hash, "legacyPasswordHash");
  for (const ProtectionFlag& flag : kProtectionFlags) {
    if (flag.tristate) {
      set_field_at<int32_t>(&p, flag.record, js_tristate(reader, input, flag.key, flag.key));
    } else {
      set_field_at<int32_t>(&p, flag.record, reader.boolean(input, flag.key, false, flag.key) ? 1 : 0);
    }
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_sheet_protection(handle_, sheet, &p));
}

// Only the keys the patch gives change; `null` clears a string or `spinCount`, and an `allow*` flag takes no `null`.
void JsWorkbook::updateSheetProtection(uint32_t sheet, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateSheetProtection");
  fm_sheet_protection_patch_t p = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  std::string text[4];
  const auto string_key = [&](const char* key, std::string& storage, fm_patch_op_t* op, const char** value) {
    js_patch_string(reader, patch, key, key, true, storage, op, value);
  };
  string_key("algorithmName", text[0], &p.algorithm_name_op, &p.algorithm_name);
  string_key("hashValue", text[1], &p.hash_value_op, &p.hash_value);
  string_key("saltValue", text[2], &p.salt_value_op, &p.salt_value);
  string_key("legacyPasswordHash", text[3], &p.legacy_password_hash_op, &p.legacy_password_hash);
  js_patch_u32(reader, patch, "spinCount", "spinCount", true, &p.spin_count_op, &p.spin_count);
  for (const ProtectionFlag& flag : kProtectionFlags) {
    const emscripten::val v = reader.value(patch, flag.key, flag.key);
    if (js_patch_key(reader, v, flag.key, false) == JsPatchKey::kSet) {
      set_field_at<int32_t>(&p, flag.patch_op, FM_PATCH_OP_SET);
      const bool allowed = reader.boolean_value(v, false, flag.key);
      // A patch value is a plain 0 / 1; only the record's select flags carry a tristate.
      set_field_at<int32_t>(&p, flag.patch_value, allowed ? 1 : 0);
    }
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_sheet_protection(handle_, sheet, &p));
}

void JsWorkbook::updateSheetView(uint32_t sheet, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateView");
  fm_sheet_view_patch_t p = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  js_patch_u32(reader, patch, "zoomPct", "zoomPct", false, &p.zoom_pct_op, &p.zoom_pct);
  js_patch_u32(reader, patch, "freezeRows", "freezeRows", false, &p.freeze_rows_op, &p.freeze_rows);
  js_patch_u32(reader, patch, "freezeCols", "freezeCols", false, &p.freeze_cols_op, &p.freeze_cols);
  js_patch_flag(reader, patch, "showGridLines", "showGridLines", &p.show_grid_lines_op, &p.show_grid_lines);
  js_patch_flag(reader, patch, "showRowColumnHeaders", "showRowColumnHeaders", &p.show_row_column_headers_op,
                &p.show_row_column_headers);
  js_patch_flag(reader, patch, "showZeros", "showZeros", &p.show_zeros_op, &p.show_zeros);
  js_patch_flag(reader, patch, "rightToLeft", "rightToLeft", &p.right_to_left_op, &p.right_to_left);
  js_patch_flag(reader, patch, "tabSelected", "tabSelected", &p.tab_selected_op, &p.tab_selected);
  js_patch_int(reader, patch, "viewMode", "viewMode", false, &p.view_mode_op, &p.view_mode);
  js_patch_u32(reader, patch, "topLeftRow", "topLeftRow", false, &p.top_left_row_op, &p.top_left_row);
  js_patch_u32(reader, patch, "topLeftCol", "topLeftCol", false, &p.top_left_col_op, &p.top_left_col);
  js_patch_u32(reader, patch, "paneTopLeftRow", "paneTopLeftRow", false, &p.pane_top_left_row_op, &p.pane_top_left_row);
  js_patch_u32(reader, patch, "paneTopLeftCol", "paneTopLeftCol", false, &p.pane_top_left_col_op, &p.pane_top_left_col);
  js_patch_number(reader, patch, "splitXTwips", "splitXTwips", false, &p.split_x_twips_op, &p.split_x_twips);
  js_patch_number(reader, patch, "splitYTwips", "splitYTwips", false, &p.split_y_twips_op, &p.split_y_twips);
  js_patch_flag(reader, patch, "showFormulas", "showFormulas", &p.show_formulas_op, &p.show_formulas);
  js_patch_flag(reader, patch, "showOutlineSymbols", "showOutlineSymbols", &p.show_outline_symbols_op,
                &p.show_outline_symbols);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_view(handle_, sheet, &p));
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
  emscripten::val active = emscripten::val::object();
  active.set("row", row);
  active.set("col", col);
  o.set("active", active);
  o.set("ranges", cell_rects_to_val(found.data(), found.size()));
  return o;
}

void JsWorkbook::setSheetSelection(uint32_t sheet, emscripten::val selection) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.setSelection");
  const emscripten::val active = reader.value(selection, "active", "selection.active");
  uint32_t row = 0;
  uint32_t col = 0;
  if (reader.require_object(active, "selection.active")) {
    row = reader.required_u32(active, "row", "selection.active.row");
    col = reader.required_u32(active, "col", "selection.active.col");
  }
  const emscripten::val ranges_value = reader.value(selection, "ranges", "selection.ranges");
  std::vector<fm_cell_rect_t> ranges;
  if (reader.ok()) {
    ranges = js_pull_range_array(ranges_value, "selection.ranges", &reader);
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_selection(handle_, sheet, row, col, ranges.data(), ranges.size()));
}

// ---- Column layout overrides --------------------------------------------

emscripten::val JsWorkbook::getSheetColumns(uint32_t sheet) const {
  return entry_list(handle_, sheet, &fm_sheet_get_column_layout_count, &column_entry);
}

void JsWorkbook::setColumnWidth(uint32_t sheet, uint32_t first, uint32_t last, double width) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_width(handle_, sheet, first, last, width) : kBindingInvalidHandle);
}

void JsWorkbook::setColumnHidden(uint32_t sheet, uint32_t first, uint32_t last, bool hidden) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setColumnOutline(uint32_t sheet, uint32_t first, uint32_t last, uint32_t level) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_outline(handle_, sheet, first, last, level)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setColumnCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_collapsed(handle_, sheet, first, last, collapsed ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setColumnStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_set_column_xf_index(handle_, sheet, first, last, xfIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::clearColumnStyle(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_reset_column_xf_index(handle_, sheet, first, last) : kBindingInvalidHandle);
}

// ---- Row layout overrides ----------------------------------------------

emscripten::val JsWorkbook::getSheetRowOverrides(uint32_t sheet) const {
  return entry_list(handle_, sheet, &fm_sheet_get_row_layout_count, &row_entry);
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
  succeeded(handle_ != nullptr ? fm_sheet_set_row_outline(handle_, sheet, first, last, level) : kBindingInvalidHandle);
}

void JsWorkbook::setRowCollapsed(uint32_t sheet, uint32_t first, uint32_t last, bool collapsed) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_collapsed(handle_, sheet, first, last, collapsed ? 1 : 0)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setRowStyle(uint32_t sheet, uint32_t first, uint32_t last, uint32_t xfIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_set_row_xf_index(handle_, sheet, first, last, xfIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::clearRowStyle(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_reset_row_xf_index(handle_, sheet, first, last) : kBindingInvalidHandle);
}

namespace {

emscripten::val range_style_bases_result(fm_workbook_t* handle, uint32_t sheet, const fm_cell_rect_t& rect,
                                         const std::vector<fm_cell_address_t>& excluded) {
  size_t len = 0;
  fm_status_t rc =
      fm_sheet_get_range_xf_bases_excluding(handle, sheet, &rect, excluded.data(), excluded.size(), nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_range_xf_bases_excluding(handle, sheet, &rect, excluded.data(), excluded.size(), xfs.data(),
                                               xfs.size(), &len);
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
  JsNarrowNumericReader reader("Sheet.getRangeXfBases");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  return range_style_bases_result(handle_, sheet, rect, {});
}

emscripten::val JsWorkbook::listCellXfIndices(uint32_t sheet, emscripten::val range) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.listCellXfIndices");
  const fm_cell_rect_t rect = js_value_present(range) ? js_pull_range(range, &reader) : kWholeSheetRect;
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  const fm_cell_rect_t* where = &rect;
  size_t len = 0;
  fm_status_t rc = fm_sheet_get_cell_xf_indices(handle_, sheet, where, nullptr, 0, &len);
  std::vector<fm_xf_cell_t> cells(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_cell_xf_indices(handle_, sheet, where, cells.data(), cells.size(), &len);
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val list = emscripten::val::array();
  for (const fm_xf_cell_t& cell : cells) {
    emscripten::val item = emscripten::val::object();
    item.set("row", cell.row);
    item.set("col", cell.col);
    item.set("xfIndex", cell.xf_index);
    list.call<void>("push", item);
  }
  return list;
}

emscripten::val JsWorkbook::listShowPhoneticCells(uint32_t sheet, emscripten::val range) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.listShowPhoneticCells");
  const fm_cell_rect_t rect = js_value_present(range) ? js_pull_range(range, &reader) : kWholeSheetRect;
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  const fm_cell_rect_t* where = &rect;
  size_t len = 0;
  fm_status_t rc = fm_sheet_get_show_phonetic_cells(handle_, sheet, where, nullptr, 0, &len);
  std::vector<fm_cell_address_t> cells(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_show_phonetic_cells(handle_, sheet, where, cells.data(), cells.size(), &len);
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val list = emscripten::val::array();
  for (const fm_cell_address_t& cell : cells) {
    emscripten::val item = emscripten::val::object();
    item.set("row", cell.row);
    item.set("col", cell.col);
    list.call<void>("push", item);
  }
  return list;
}

emscripten::val JsWorkbook::rangeStyleBasesExcluding(uint32_t sheet, emscripten::val range,
                                                     emscripten::val excludedRows, emscripten::val excludedCols) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.getRangeXfBasesExcluding");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  if (!reader.is_array(excludedRows, "excludedRows") || !reader.is_array(excludedCols, "excludedCols")) {
    record_type_error("Sheet.getRangeXfBasesExcluding excludedRows and excludedCols must be arrays of numbers");
    return emscripten::val::undefined();
  }
  const uint32_t count = reader.length(excludedRows, "excludedRows");
  if (reader.ok() && count != reader.length(excludedCols, "excludedCols")) {
    record_type_error("Sheet.getRangeXfBasesExcluding excludedRows and excludedCols must have the same length");
    return emscripten::val::undefined();
  }
  std::vector<fm_cell_address_t> excluded(count);
  for (uint32_t i = 0; reader.ok() && i < count; ++i) {
    excluded[i].row = reader.u32_value(reader.array_element(excludedRows, i, "excludedRows[]"), 0U, "excludedRows[]");
    excluded[i].col = reader.u32_value(reader.array_element(excludedCols, i, "excludedCols[]"), 0U, "excludedCols[]");
  }
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  return range_style_bases_result(handle_, sheet, rect, excluded);
}

void JsWorkbook::remapRangeStyle(uint32_t sheet, emscripten::val range, emscripten::val fromXfs,
                                 emscripten::val toXfs) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.remapRangeXfs");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  if (!reader.is_array(fromXfs, "fromXfs") || !reader.is_array(toXfs, "toXfs")) {
    record_type_error("Sheet.remapRangeXfs fromXfs and toXfs must be arrays of numbers");
    return;
  }
  const uint32_t count = reader.length(fromXfs, "fromXfs");
  if (reader.ok() && count != reader.length(toXfs, "toXfs")) {
    record_type_error("Sheet.remapRangeXfs fromXfs and toXfs must have the same length");
    return;
  }
  std::vector<fm_xf_remap_t> pairs(count);
  for (uint32_t i = 0; reader.ok() && i < count; ++i) {
    pairs[i].from_xf = reader.u32_value(reader.array_element(fromXfs, i, "fromXfs[]"), 0U, "fromXfs[]");
    pairs[i].to_xf = reader.u32_value(reader.array_element(toXfs, i, "toXfs[]"), 0U, "toXfs[]");
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_remap_range_xfs(handle_, sheet, &rect, pairs.data(), pairs.size()));
}

emscripten::val JsWorkbook::snapshotRangeStyle(uint32_t sheet, emscripten::val range) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.snapshotRangeXfs");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_range_xfs(handle_, sheet, &rect, &data, &len);
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
  succeeded(fm_sheet_restore_range_xfs(handle_, sheet, data.data(), data.size()));
}

void JsWorkbook::clearRowHeight(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_reset_row_height(handle_, sheet, first, last) : kBindingInvalidHandle);
}

void JsWorkbook::clearColumnWidth(uint32_t sheet, uint32_t first, uint32_t last) {
  succeeded(handle_ != nullptr ? fm_sheet_reset_column_width(handle_, sheet, first, last) : kBindingInvalidHandle);
}

// ---- Sheet format defaults and point geometry ---------------------------

emscripten::val JsWorkbook::getSheetFormatDefaults(uint32_t sheet) const {
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_format_defaults(handle_, sheet, &d) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_present_or_null(o, "defaultColumnWidthChars", d.has_default_column_width_chars, d.default_column_width_chars);
  js_set_present_or_null(o, "defaultRowHeightPt", d.has_default_row_height_pt, d.default_row_height_pt);
  o.set("baseColumnWidthChars", d.base_column_width_chars);
  o.set("customHeight", d.custom_height != 0);
  o.set("zeroHeight", d.zero_height != 0);
  o.set("effectiveDefaultColumnWidthChars", d.effective_default_column_width_chars);
  o.set("effectiveDefaultRowHeightPt", d.effective_default_row_height_pt);
  return o;
}

void JsWorkbook::setSheetFormatDefaults(uint32_t sheet, emscripten::val defaults) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // An omitted key keeps the stored value, `null` clears a nullable one.
  JsNarrowNumericReader reader("Sheet.updateFormatDefaults");
  if (!reader.require_object(defaults, "defaults")) {
    return;
  }
  fm_sheet_format_defaults_patch_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  js_patch_number(reader, defaults, "defaultColumnWidthChars", "defaultColumnWidthChars", true,
                  &d.default_column_width_chars_op, &d.default_column_width_chars);
  js_patch_number(reader, defaults, "defaultRowHeightPt", "defaultRowHeightPt", true, &d.default_row_height_pt_op,
                  &d.default_row_height_pt);
  js_patch_number(reader, defaults, "baseColumnWidthChars", "baseColumnWidthChars", false,
                  &d.base_column_width_chars_op, &d.base_column_width_chars);
  js_patch_flag(reader, defaults, "customHeight", "customHeight", &d.custom_height_op, &d.custom_height);
  js_patch_flag(reader, defaults, "zeroHeight", "zeroHeight", &d.zero_height_op, &d.zero_height);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_format_defaults(handle_, sheet, &d));
}

emscripten::val JsWorkbook::getCellRectPt(uint32_t sheet, emscripten::val range, int32_t mode) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.getCellRectPt");
  const fm_cell_rect_t b = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  fm_rect_pt_info_t r = FM_STRUCT_INIT(fm_rect_pt_info_t);
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
  fm_width_model_info_t m = FM_STRUCT_INIT(fm_width_model_info_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_width_model(handle_, sheet, mode, &m) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("pointsPerChar", m.points_per_char);
  o.set("paddingPt", m.padding_pt);
  o.set("normalFontSizePt", m.normal_font_size_pt);
  o.set("calibrated", m.calibrated != 0);
  js_set_cstr(o, "normalFontName", m.normal_font_name);
  o.set("platform", m.platform);
  return o;
}

double JsWorkbook::columnCharsToPt(uint32_t sheet, int32_t mode, double chars) const {
  double pt = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_convert_column_chars_to_pt(handle_, sheet, chars, mode, &pt)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return pt;
}

double JsWorkbook::columnPtToChars(uint32_t sheet, int32_t mode, double pt) const {
  double chars = 0.0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_convert_column_pt_to_chars(handle_, sheet, pt, mode, &chars)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return chars;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
