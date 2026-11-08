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
  fm_column_layout_t entry = FM_STRUCT_INIT(fm_column_layout_t);
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
  fm_row_layout_t entry = FM_STRUCT_INIT(fm_row_layout_t);
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

uint8_t clamp_outline_level(uint32_t level) {
  return static_cast<uint8_t>(level > 255U ? 255U : level);
}

emscripten::val js_rect_pt(const fm_rect_pt_t& r) {
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
  fm_margins_pt_t margins = FM_STRUCT_INIT(fm_margins_pt_t);
  fm_rect_pt_t printable = FM_STRUCT_INIT(fm_rect_pt_t);
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
}

/// The span a declared break at `position` covers on the perpendicular axis; a break the sheet does not
/// declare (computed by the page walk) covers the whole sheet, which is `0..0`.
void declared_break_span(const fm_workbook_t* handle, uint32_t sheet, bool rows, uint32_t position, uint32_t* first,
                         uint32_t* last) {
  *first = 0;
  *last = 0;
  size_t count = 0;
  if (rows ? fm_sheet_get_row_break_count(handle, sheet, &count) != 0
           : fm_sheet_get_column_break_count(handle, sheet, &count) != 0) {
    return;
  }
  for (size_t i = 0; i < count; ++i) {
    if (rows) {
      fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
      if (fm_sheet_get_row_break_at(handle, sheet, i, &brk) == 0 && brk.row == position) {
        *first = brk.first;
        *last = brk.last;
        return;
      }
    } else {
      fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
      if (fm_sheet_get_column_break_at(handle, sheet, i, &brk) == 0 && brk.col == position) {
        *first = brk.first;
        *last = brk.last;
        return;
      }
    }
  }
}

/// The breaks in effect as `RowBreak` / `ColumnBreak` records, in ascending position.
emscripten::val pagination_breaks(const fm_workbook_t* handle, uint32_t sheet, const fm_pagination_t* pagination,
                                  bool rows) {
  emscripten::val out = emscripten::val::array();
  size_t count = 0;
  if (rows) {
    fm_pagination_get_row_break_count(pagination, &count);
  } else {
    fm_pagination_get_column_break_count(pagination, &count);
  }
  for (size_t i = 0; i < count; ++i) {
    uint32_t position = 0;
    int32_t manual = 0;
    const fm_status_t at_rc = rows ? fm_pagination_get_row_break_at(pagination, i, &position)
                                   : fm_pagination_get_column_break_at(pagination, i, &position);
    if (at_rc != 0) {
      continue;
    }
    if (rows) {
      fm_pagination_is_row_break_manual(pagination, i, &manual);
    } else {
      fm_pagination_is_column_break_manual(pagination, i, &manual);
    }
    uint32_t first = 0;
    uint32_t last = 0;
    declared_break_span(handle, sheet, rows, position, &first, &last);
    emscripten::val entry = emscripten::val::object();
    entry.set(rows ? "row" : "col", position);
    entry.set("first", first);
    entry.set("last", last);
    entry.set("manual", manual != 0);
    out.call<void>("push", entry);
  }
  return out;
}

}  // namespace

emscripten::val JsWorkbook::paginate(uint32_t sheet) const {
  emscripten::val result = emscripten::val::object();
  emscripten::val print_area = emscripten::val::array();
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
  result.set("printArea", print_area);
  result.set("rowBreaks", pagination_breaks(handle_, sheet, pagination, /*rows=*/true));
  result.set("columnBreaks", pagination_breaks(handle_, sheet, pagination, /*rows=*/false));
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
  r.zoomPct = v.zoom_pct;
  r.freezeRows = v.freeze_rows;
  r.freezeCols = v.freeze_cols;
  r.visibility = v.visibility;
  r.showGridLines = v.show_grid_lines != 0;
  r.showRowColHeaders = v.show_row_column_headers != 0;
  r.showZeros = v.show_zeros != 0;
  r.rightToLeft = v.right_to_left != 0;
  r.tabSelected = v.tab_selected != 0;
  r.viewMode = v.view_mode;
  r.topLeftRow = v.top_left_row;
  r.topLeftCol = v.top_left_col;
  r.paneTopLeftRow = v.pane_top_left_row;
  r.paneTopLeftCol = v.pane_top_left_col;
  r.splitXTwips = v.split_x;
  r.splitYTwips = v.split_y;
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
  const auto string_key = [&](const char* key, std::string& storage, int32_t* op, const char** value) {
    const emscripten::val v = reader.value(patch, key, key);
    switch (js_patch_key(reader, v, key, true)) {
      case JsPatchKey::kClear:
        *op = FM_CHART_PATCH_OP_CLEAR;
        break;
      case JsPatchKey::kSet:
        *op = FM_CHART_PATCH_OP_SET;
        *value = reader.optional_string_value(v, storage, key);
        break;
      case JsPatchKey::kKeep:
        break;
    }
  };
  string_key("algorithmName", text[0], &p.algorithm_name_op, &p.algorithm_name);
  string_key("hashValue", text[1], &p.hash_value_op, &p.hash_value);
  string_key("saltValue", text[2], &p.salt_value_op, &p.salt_value);
  string_key("legacyPasswordHash", text[3], &p.legacy_password_hash_op, &p.legacy_password_hash);
  const emscripten::val spin = reader.value(patch, "spinCount", "spinCount");
  switch (js_patch_key(reader, spin, "spinCount", true)) {
    case JsPatchKey::kClear:
      p.spin_count_op = FM_CHART_PATCH_OP_CLEAR;
      break;
    case JsPatchKey::kSet:
      p.spin_count_op = FM_CHART_PATCH_OP_SET;
      p.spin_count = reader.u32_value(spin, 0U, "spinCount");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  for (const ProtectionFlag& flag : kProtectionFlags) {
    const emscripten::val v = reader.value(patch, flag.key, flag.key);
    if (js_patch_key(reader, v, flag.key, false) == JsPatchKey::kSet) {
      set_field_at<int32_t>(&p, flag.patch_op, FM_CHART_PATCH_OP_SET);
      const bool allowed = reader.boolean_value(v, false);
      // A patch value is a plain 0 / 1; only the record's select flags carry a tristate.
      set_field_at<int32_t>(&p, flag.patch_value, allowed ? 1 : 0);
    }
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_sheet_protection(handle_, sheet, &p));
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
  succeeded(handle_ != nullptr ? fm_sheet_set_show_row_column_headers(handle_, sheet, show ? 1 : 0)
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
  succeeded(handle_ != nullptr ? fm_sheet_set_column_outline(handle_, sheet, first, last, clamp_outline_level(level))
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
  succeeded(handle_ != nullptr ? fm_sheet_set_row_outline(handle_, sheet, first, last, clamp_outline_level(level))
                               : kBindingInvalidHandle);
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
                                         const std::vector<uint32_t>& rows, const std::vector<uint32_t>& cols) {
  size_t len = 0;
  fm_status_t rc = fm_sheet_get_range_xf_bases_excluding(handle, sheet, &rect, rows.data(), cols.data(), rows.size(),
                                                         nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_range_xf_bases_excluding(handle, sheet, &rect, rows.data(), cols.data(), rows.size(), xfs.data(),
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
  JsNarrowNumericReader reader("getRangeXfBases");
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
  JsNarrowNumericReader reader("getRangeXfBasesExcluding");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  if (!reader.is_array(excludedRows, "excludedRows") || !reader.is_array(excludedCols, "excludedCols")) {
    record_type_error(
        "getRangeXfBasesExcluding: `excludedRows` and `excludedCols` must be arrays "
        "of numbers");
    return emscripten::val::undefined();
  }
  const uint32_t count = reader.length(excludedRows, "excludedRows");
  if (reader.ok() && count != reader.length(excludedCols, "excludedCols")) {
    record_type_error("getRangeXfBasesExcluding: `excludedRows` and `excludedCols` must have the same length");
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
  succeeded(fm_sheet_remap_range_xfs(handle_, sheet, &rect, from.data(), to.data(), count));
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
  js_set_present_or_null(o, "defaultColWidthChars", d.has_default_column_width_chars, d.default_column_width_chars);
  js_set_present_or_null(o, "defaultRowHeightPt", d.has_default_row_height_pt, d.default_row_height_pt);
  o.set("baseColWidthChars", d.base_column_width_chars);
  o.set("customHeight", d.custom_height != 0);
  o.set("zeroHeight", d.zero_height != 0);
  o.set("effectiveDefaultColWidthChars", d.effective_default_column_width_chars);
  o.set("effectiveDefaultRowHeightPt", d.effective_default_row_height_pt);
  return o;
}

void JsWorkbook::setSheetFormatDefaults(uint32_t sheet, emscripten::val defaults) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // A patch: an omitted key keeps the stored value, `null` clears a nullable one.
  JsNarrowNumericReader reader("Sheet.updateFormatDefaults");
  if (!reader.require_object(defaults, "defaults")) {
    return;
  }
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  if (!succeeded(fm_sheet_get_format_defaults(handle_, sheet, &d))) {
    return;
  }
  const emscripten::val col_width = reader.value(defaults, "defaultColWidthChars", "defaultColWidthChars");
  switch (js_patch_key(reader, col_width, "defaultColWidthChars", true)) {
    case JsPatchKey::kClear:
      d.has_default_column_width_chars = 0;
      d.default_column_width_chars = 0.0;
      break;
    case JsPatchKey::kSet:
      d.has_default_column_width_chars = 1;
      d.default_column_width_chars = reader.number_value(col_width, 0.0, "defaultColWidthChars");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  const emscripten::val row_height = reader.value(defaults, "defaultRowHeightPt", "defaultRowHeightPt");
  switch (js_patch_key(reader, row_height, "defaultRowHeightPt", true)) {
    case JsPatchKey::kClear:
      d.has_default_row_height_pt = 0;
      d.default_row_height_pt = 0.0;
      break;
    case JsPatchKey::kSet:
      d.has_default_row_height_pt = 1;
      d.default_row_height_pt = reader.number_value(row_height, 0.0, "defaultRowHeightPt");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  const emscripten::val base = reader.value(defaults, "baseColWidthChars", "baseColWidthChars");
  if (js_patch_key(reader, base, "baseColWidthChars", false) == JsPatchKey::kSet) {
    d.base_column_width_chars = reader.number_value(base, d.base_column_width_chars, "baseColWidthChars");
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
  succeeded(fm_sheet_update_format_defaults(handle_, sheet, &d));
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
  o.set("normalFontSizePt", m.normal_font_size_pt);
  o.set("calibrated", m.calibrated != 0);
  js_set_cstr(o, "normalFontName", m.normal_font_name);
  o.set("platform", m.platform);
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
