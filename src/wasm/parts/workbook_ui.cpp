//
// JsWorkbook UI-feature surface: merges, hyperlinks, comments, and
// data-validation rules. Each accessor returns a JS-friendly value
// (`Array<...>` or `null`) so JS callers don't have to step through
// count + getter pairs.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- AutoFilter --------------------------------------------------------

emscripten::val JsWorkbook::getSheetAutoFilterXml(uint32_t sheet) const {
  emscripten::val out = emscripten::val::object();
  if (handle_ == nullptr) {
    out.set("status", error_status(7000));
    out.set("xml", std::string());
    return out;
  }
  const char* xml = nullptr;
  fm_status_t rc = fm_sheet_get_auto_filter_xml(handle_, sheet, &xml);
  if (rc != 0) {
    out.set("status", error_status(rc));
    out.set("xml", std::string());
    return out;
  }
  out.set("status", ok_status());
  js_set_cstr(out, "xml", xml);
  return out;
}

JsStatus JsWorkbook::setSheetAutoFilterXml(uint32_t sheet, const std::string& xml) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_set_auto_filter_xml(handle_, sheet, xml.c_str()));
}

// ---- Merges ------------------------------------------------------------

JsStatus JsWorkbook::addMerge(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const fm_merge_range m = js_pull_range(range);
  fm_status_t rc = fm_sheet_add_merge(handle_, sheet, m);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeMerge(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const fm_merge_range m = js_pull_range(range);
  fm_status_t rc = fm_sheet_remove_merge(handle_, sheet, m);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeMergeAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_merge_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearMerges(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_clear_merges(handle_, sheet);
  return status_from_rc(rc);
}

namespace {

// Reads list element `index` of `sheet` into `*out`; the per-kind half of
// `sheet_list`.
using SheetListItemFn = fm_status_t (*)(fm_workbook_t*, uint32_t, uint32_t, emscripten::val*);

// Enumerates a per-sheet list through its count / at-index pair into a
// `ListResult` array, stopping at the first failed element read.
emscripten::val sheet_list(fm_workbook_t* wb, uint32_t sheet,
                           fm_status_t (*count_fn)(fm_workbook_t*, uint32_t, uint32_t*), SheetListItemFn item_fn) {
  emscripten::val arr = emscripten::val::array();
  if (wb == nullptr) {
    arr.set("status", error_status(7000));
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = count_fn(wb, sheet, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  for (uint32_t i = 0; i < count; ++i) {
    emscripten::val item = emscripten::val::undefined();
    rc = item_fn(wb, sheet, i, &item);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
    }
    arr.set(i, item);
  }
  arr.set("status", ok_status());
  return arr;
}

fm_status_t merge_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_merge_range m{};
  const fm_status_t rc = fm_sheet_get_merge_at(wb, sheet, index, &m);
  if (rc != 0) {
    return rc;
  }
  *out = merge_range_to_val(m);
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getMerges(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_merge_count, &merge_item);
}

emscripten::val JsWorkbook::getMergesInRange(uint32_t sheet, emscripten::val range) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(7000));
    return arr;
  }
  const fm_merge_range query = js_pull_range(range);
  uint32_t total = 0;
  fm_status_t rc = fm_sheet_merges_in_range(handle_, sheet, query, nullptr, 0, &total);
  std::vector<fm_merge_range> found(rc == 0 ? total : 0);
  if (rc == 0 && total > 0) {
    uint32_t written = 0;
    rc = fm_sheet_merges_in_range(handle_, sheet, query, found.data(), total, &written);
    found.resize(rc == 0 ? written : 0);
  }
  if (rc != 0) {
    arr.set("status", error_status(rc));
    return arr;
  }
  for (uint32_t i = 0; i < found.size(); ++i) {
    arr.set(i, merge_range_to_val(found[i]));
  }
  arr.set("status", ok_status());
  return arr;
}

// ---- Comments ----------------------------------------------------------

emscripten::val JsWorkbook::getComment(uint32_t sheet, uint32_t row, uint32_t col) const {
  if (handle_ == nullptr) {
    return emscripten::val::null();
  }
  fm_comment c{};
  if (fm_sheet_get_comment_at(handle_, sheet, row, col, &c) != 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_fields(o, {{"author", c.author}, {"text", c.text}});
  return o;
}

emscripten::val JsWorkbook::getCommentResult(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val out = emscripten::val::object();
  if (handle_ == nullptr) {
    out.set("status", error_status(7000));
    out.set("comment", emscripten::val::null());
    return out;
  }
  fm_comment c{};
  const fm_status_t rc = fm_sheet_get_comment_at(handle_, sheet, row, col, &c);
  if (rc != 0) {
    out.set("status", status_from_rc(rc));
    out.set("comment", emscripten::val::null());
    return out;
  }
  emscripten::val comment = emscripten::val::object();
  js_set_cstr_fields(comment, {{"author", c.author}, {"text", c.text}});
  out.set("status", ok_status());
  out.set("comment", comment);
  return out;
}

namespace {

fm_status_t comment_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_comment c{};
  const fm_status_t rc = fm_sheet_get_comment_at_index(wb, sheet, index, &c);
  if (rc != 0) {
    return rc;
  }
  emscripten::val o = emscripten::val::object();
  o.set("row", c.row);
  o.set("col", c.col);
  js_set_cstr_fields(o, {{"author", c.author}, {"text", c.text}});
  *out = o;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getComments(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_comment_count, &comment_item);
}

JsStatus JsWorkbook::setComment(uint32_t sheet, uint32_t row, uint32_t col, const std::string& author,
                                const std::string& text) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const char* author_c = author.empty() ? nullptr : author.c_str();
  const char* text_c = text.empty() ? nullptr : text.c_str();
  fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
  return status_from_rc(rc);
}

// ---- Hyperlinks --------------------------------------------------------

JsStatus JsWorkbook::addHyperlink(uint32_t sheet, uint32_t row, uint32_t col, const std::string& target,
                                  const std::string& display, const std::string& tooltip, const std::string& location) {
  return addHyperlinkRange(sheet, row, col, row, col, target, display, tooltip, location);
}

JsStatus JsWorkbook::addHyperlinkRange(uint32_t sheet, uint32_t row, uint32_t col, uint32_t lastRow, uint32_t lastCol,
                                       const std::string& target, const std::string& display,
                                       const std::string& tooltip, const std::string& location) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_hyperlink hl{};
  hl.row = row;
  hl.col = col;
  hl.last_row = lastRow;
  hl.last_col = lastCol;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  fm_status_t rc = fm_sheet_add_hyperlink(handle_, sheet, hl);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeHyperlink(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_hyperlink(handle_, sheet, row, col);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeHyperlinkAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_hyperlink_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearHyperlinks(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_clear_hyperlinks(handle_, sheet);
  return status_from_rc(rc);
}

namespace {

fm_status_t hyperlink_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_hyperlink h{};
  const fm_status_t rc = fm_sheet_get_hyperlink_at(wb, sheet, index, &h);
  if (rc != 0) {
    return rc;
  }
  emscripten::val item = emscripten::val::object();
  item.set("row", h.row);
  item.set("col", h.col);
  item.set("lastRow", h.last_row);
  item.set("lastCol", h.last_col);
  js_set_cstr_fields(item,
                     {{"target", h.target}, {"location", h.location}, {"display", h.display}, {"tooltip", h.tooltip}});
  *out = item;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getHyperlinks(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_hyperlink_count, &hyperlink_item);
}

// ---- Data validations --------------------------------------------------

namespace {

fm_status_t validation_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_data_validation v{};
  const fm_status_t rc = fm_sheet_get_validation_at(wb, sheet, index, &v);
  if (rc != 0) {
    return rc;
  }
  emscripten::val item = emscripten::val::object();
  emscripten::val ranges = emscripten::val::array();
  for (uint32_t r = 0; r < v.range_count; ++r) {
    ranges.set(r, merge_range_to_val(v.ranges[r]));
  }
  item.set("ranges", ranges);
  item.set("type", v.type);
  item.set("op", v.op);
  item.set("errorStyle", v.error_style);
  item.set("allowBlank", v.allow_blank != 0);
  item.set("showInputMessage", v.show_input_message != 0);
  item.set("showErrorMessage", v.show_error_message != 0);
  item.set("showDropDown", v.show_dropdown != 0);
  js_set_cstr_fields(item, {{"formula1", v.formula1},
                            {"formula2", v.formula2},
                            {"errorTitle", v.error_title},
                            {"errorMessage", v.error_message},
                            {"promptTitle", v.prompt_title},
                            {"promptMessage", v.prompt_message}});
  *out = item;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getValidations(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_validation_count, &validation_item);
}

JsStatus JsWorkbook::addValidation(uint32_t sheet, emscripten::val v) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  const std::vector<fm_merge_range> ranges_buf = js_pull_ranges(v, "ranges");
  const std::string formula1 = js_pull_string(v, "formula1");
  const std::string formula2 = js_pull_string(v, "formula2");
  const std::string error_title = js_pull_string(v, "errorTitle");
  const std::string error_message = js_pull_string(v, "errorMessage");
  const std::string prompt_title = js_pull_string(v, "promptTitle");
  const std::string prompt_message = js_pull_string(v, "promptMessage");

  fm_data_validation dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = static_cast<uint32_t>(ranges_buf.size());
  dv.type = js_pull_u8(v, "type", 0U);
  dv.op = js_pull_u8(v, "op", 0U);
  dv.error_style = js_pull_u8(v, "errorStyle", 0U);
  dv.allow_blank = js_pull_bool(v, "allowBlank", false) ? 1 : 0;
  dv.show_input_message = js_pull_bool(v, "showInputMessage", false) ? 1 : 0;
  dv.show_error_message = js_pull_bool(v, "showErrorMessage", false) ? 1 : 0;
  dv.show_dropdown = js_pull_bool(v, "showDropDown", true) ? 1 : 0;
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, dv);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeValidationAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_validation_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearValidations(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_clear_validations(handle_, sheet);
  return status_from_rc(rc);
}

// ---- Typed AutoFilter --------------------------------------------------

namespace {

/// The six AutoFilter entry points, once for sheets and once for tables, so
/// every JS method below is written a single time.
struct AutoFilterOps {
  fm_status_t (*get)(const fm_workbook_t*, size_t, fm_auto_filter*, int32_t*);
  fm_status_t (*set)(fm_workbook_t*, size_t, const fm_auto_filter*);
  fm_status_t (*remove)(fm_workbook_t*, size_t);
  fm_status_t (*apply)(fm_workbook_t*, size_t);
  fm_status_t (*clear)(fm_workbook_t*, size_t);
  fm_status_t (*evaluate)(const fm_workbook_t*, size_t, uint8_t*, size_t, size_t*, uint32_t*);
};

constexpr AutoFilterOps kSheetAutoFilter = {fm_sheet_get_auto_filter,    fm_sheet_set_auto_filter,
                                            fm_sheet_remove_auto_filter, fm_sheet_apply_auto_filter,
                                            fm_sheet_clear_auto_filter,  fm_sheet_evaluate_auto_filter};
constexpr AutoFilterOps kTableAutoFilter = {fm_table_get_auto_filter,    fm_table_set_auto_filter,
                                            fm_table_remove_auto_filter, fm_table_apply_auto_filter,
                                            fm_table_clear_auto_filter,  fm_table_evaluate_auto_filter};

emscripten::val date_group_to_val(const fm_date_group_item& d) {
  emscripten::val o = emscripten::val::object();
  o.set("year", static_cast<uint32_t>(d.year));
  o.set("month", static_cast<uint32_t>(d.month));
  o.set("day", static_cast<uint32_t>(d.day));
  o.set("hour", static_cast<uint32_t>(d.hour));
  o.set("minute", static_cast<uint32_t>(d.minute));
  o.set("second", static_cast<uint32_t>(d.second));
  o.set("grouping", static_cast<uint32_t>(d.grouping));
  return o;
}

emscripten::val filter_column_to_val(const fm_filter_column& c) {
  emscripten::val o = emscripten::val::object();
  o.set("colId", c.col_id);
  o.set("hiddenButton", c.hidden_button != 0);
  o.set("showButton", c.show_button != 0);
  o.set("kind", c.kind);
  o.set("filterBlank", c.filter_blank != 0);
  emscripten::val values = emscripten::val::array();
  for (uint32_t i = 0; i < c.value_count; ++i) {
    values.set(i, std::string(c.values[i]));
  }
  o.set("values", values);
  emscripten::val groups = emscripten::val::array();
  for (uint32_t i = 0; i < c.date_group_count; ++i) {
    groups.set(i, date_group_to_val(c.date_groups[i]));
  }
  o.set("dateGroups", groups);
  o.set("customAnd", c.custom_and != 0);
  o.set("customCount", c.custom_count);
  o.set("op1", c.op1);
  o.set("op2", c.op2);
  o.set("top", c.top != 0);
  o.set("percent", c.percent != 0);
  o.set("hasFilterVal", c.has_filter_val != 0);
  o.set("topVal", c.top_val);
  o.set("filterVal", c.filter_val);
  o.set("dynamicType", c.dynamic_type);
  o.set("hasDynVal", c.has_dyn_val != 0);
  o.set("hasDynMaxVal", c.has_dyn_max_val != 0);
  o.set("dynVal", c.dyn_val);
  o.set("dynMaxVal", c.dyn_max_val);
  o.set("dxfId", c.dxf_id);
  o.set("cellColor", c.cell_color != 0);
  o.set("iconSet", c.icon_set);
  o.set("iconId", c.icon_id);
  o.set("hasIconId", c.has_icon_id != 0);
  js_set_cstr_fields(o, {{"val1", c.val1}, {"val2", c.val2}, {"valIso", c.val_iso}, {"maxValIso", c.max_val_iso}});
  return o;
}

emscripten::val sort_condition_to_val(const fm_sort_condition& c) {
  emscripten::val o = emscripten::val::object();
  o.set("ref", merge_range_to_val(c.ref));
  o.set("descending", c.descending != 0);
  o.set("sortBy", c.sort_by);
  o.set("dxfId", c.dxf_id);
  o.set("hasDxfId", c.has_dxf_id != 0);
  o.set("iconSet", c.icon_set);
  o.set("iconId", c.icon_id);
  o.set("hasIconId", c.has_icon_id != 0);
  js_set_cstr(o, "customList", c.custom_list);
  return o;
}

emscripten::val auto_filter_to_val(const fm_auto_filter& f) {
  emscripten::val o = emscripten::val::object();
  o.set("range", merge_range_to_val(f.range));
  emscripten::val columns = emscripten::val::array();
  for (uint32_t i = 0; i < f.column_count; ++i) {
    columns.set(i, filter_column_to_val(f.columns[i]));
  }
  o.set("columns", columns);
  if (f.has_sort == 0) {
    o.set("sort", emscripten::val::null());
    return o;
  }
  emscripten::val sort = emscripten::val::object();
  sort.set("ref", merge_range_to_val(f.sort_ref));
  sort.set("columnSort", f.column_sort != 0);
  sort.set("caseSensitive", f.case_sensitive != 0);
  sort.set("sortMethod", f.sort_method);
  emscripten::val conditions = emscripten::val::array();
  for (uint32_t i = 0; i < f.condition_count; ++i) {
    conditions.set(i, sort_condition_to_val(f.conditions[i]));
  }
  sort.set("conditions", conditions);
  o.set("sort", sort);
  return o;
}

/// Owns the strings and arrays one `fm_filter_column` points into.
struct FilterColumnStore {
  std::vector<std::string> values;
  std::vector<const char*> value_ptrs;
  std::vector<fm_date_group_item> date_groups;
  std::string val1;
  std::string val2;
  std::string val_iso;
  std::string max_val_iso;
};

/// Reads the JS array `v[key]`; absent or non-array reads as empty.
emscripten::val js_pull_array(const emscripten::val& v, const char* key) {
  emscripten::val a = v[key];
  return a.isArray() ? a : emscripten::val::array();
}

fm_date_group_item pull_date_group(const emscripten::val& v) {
  fm_date_group_item d{};
  d.year = js_pull_u16(v, "year", 0U);
  d.month = js_pull_u8(v, "month", 0U);
  d.day = js_pull_u8(v, "day", 0U);
  d.hour = js_pull_u8(v, "hour", 0U);
  d.minute = js_pull_u8(v, "minute", 0U);
  d.second = js_pull_u8(v, "second", 0U);
  d.grouping = js_pull_u8(v, "grouping", 0U);
  return d;
}

int32_t pull_flag(const emscripten::val& v, const char* key) {
  return js_pull_bool(v, key, false) ? 1 : 0;
}

void pull_filter_column(const emscripten::val& v, FilterColumnStore& st, fm_filter_column& c) {
  c.col_id = js_pull_u32(v, "colId", 0U);
  c.hidden_button = pull_flag(v, "hiddenButton");
  c.show_button = js_pull_bool(v, "showButton", true) ? 1 : 0;
  c.kind = js_pull_i32(v, "kind", 0);
  c.filter_blank = pull_flag(v, "filterBlank");
  const emscripten::val values = js_pull_array(v, "values");
  const uint32_t value_count = js_length(values);
  st.values.reserve(value_count);
  for (uint32_t i = 0; i < value_count; ++i) {
    st.values.push_back(values[i].as<std::string>());
  }
  for (const std::string& s : st.values) {
    st.value_ptrs.push_back(s.c_str());
  }
  c.values = st.value_ptrs.empty() ? nullptr : st.value_ptrs.data();
  c.value_count = static_cast<uint32_t>(st.value_ptrs.size());
  const emscripten::val groups = js_pull_array(v, "dateGroups");
  const uint32_t group_count = js_length(groups);
  for (uint32_t i = 0; i < group_count; ++i) {
    st.date_groups.push_back(pull_date_group(groups[i]));
  }
  c.date_groups = st.date_groups.empty() ? nullptr : st.date_groups.data();
  c.date_group_count = static_cast<uint32_t>(st.date_groups.size());
  c.custom_and = pull_flag(v, "customAnd");
  c.custom_count = js_pull_i32(v, "customCount", 0);
  c.op1 = js_pull_i32(v, "op1", 0);
  c.op2 = js_pull_i32(v, "op2", 0);
  st.val1 = js_pull_string(v, "val1");
  st.val2 = js_pull_string(v, "val2");
  st.val_iso = js_pull_string(v, "valIso");
  st.max_val_iso = js_pull_string(v, "maxValIso");
  c.val1 = st.val1.c_str();
  c.val2 = st.val2.c_str();
  c.val_iso = st.val_iso.c_str();
  c.max_val_iso = st.max_val_iso.c_str();
  c.top = pull_flag(v, "top");
  c.percent = pull_flag(v, "percent");
  c.has_filter_val = pull_flag(v, "hasFilterVal");
  c.top_val = js_pull_double(v, "topVal", 0.0);
  c.filter_val = js_pull_double(v, "filterVal", 0.0);
  c.dynamic_type = js_pull_i32(v, "dynamicType", 0);
  c.has_dyn_val = pull_flag(v, "hasDynVal");
  c.has_dyn_max_val = pull_flag(v, "hasDynMaxVal");
  c.dyn_val = js_pull_double(v, "dynVal", 0.0);
  c.dyn_max_val = js_pull_double(v, "dynMaxVal", 0.0);
  c.dxf_id = js_pull_u32(v, "dxfId", UINT32_MAX);
  c.cell_color = pull_flag(v, "cellColor");
  c.icon_set = js_pull_i32(v, "iconSet", 0);
  c.icon_id = js_pull_i32(v, "iconId", 0);
  c.has_icon_id = pull_flag(v, "hasIconId");
}

void pull_sort_condition(const emscripten::val& v, std::string& custom_list, fm_sort_condition& c) {
  c.ref = js_pull_range(v["ref"]);
  c.descending = pull_flag(v, "descending");
  c.sort_by = js_pull_i32(v, "sortBy", 0);
  custom_list = js_pull_string(v, "customList");
  c.custom_list = custom_list.c_str();
  c.dxf_id = js_pull_u32(v, "dxfId", 0U);
  c.has_dxf_id = pull_flag(v, "hasDxfId");
  c.icon_set = js_pull_i32(v, "iconSet", 0);
  c.icon_id = js_pull_i32(v, "iconId", 0);
  c.has_icon_id = pull_flag(v, "hasIconId");
}

emscripten::val get_auto_filter(const fm_workbook_t* wb, const AutoFilterOps& ops, uint32_t index) {
  emscripten::val o = emscripten::val::object();
  o.set("autoFilter", emscripten::val::null());
  if (wb == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  fm_auto_filter f{};
  int32_t present = 0;
  const fm_status_t rc = ops.get(wb, index, &f, &present);
  o.set("status", status_from_rc(rc));
  if (rc == 0 && present != 0) {
    o.set("autoFilter", auto_filter_to_val(f));
  }
  return o;
}

JsStatus set_auto_filter(fm_workbook_t* wb, const AutoFilterOps& ops, uint32_t index, const emscripten::val& filter) {
  if (wb == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  fm_auto_filter f{};
  f.range = js_pull_range(filter["range"]);
  const emscripten::val columns = js_pull_array(filter, "columns");
  const uint32_t column_count = js_length(columns);
  std::vector<FilterColumnStore> stores(column_count);
  std::vector<fm_filter_column> cols(column_count);
  for (uint32_t i = 0; i < column_count; ++i) {
    pull_filter_column(columns[i], stores[i], cols[i]);
  }
  f.columns = cols.empty() ? nullptr : cols.data();
  f.column_count = column_count;
  const emscripten::val sort = filter["sort"];
  std::vector<std::string> custom_lists;
  std::vector<fm_sort_condition> conditions;
  if (!sort.isUndefined() && !sort.isNull()) {
    f.has_sort = 1;
    f.sort_ref = js_pull_range(sort["ref"]);
    f.column_sort = pull_flag(sort, "columnSort");
    f.case_sensitive = pull_flag(sort, "caseSensitive");
    f.sort_method = js_pull_i32(sort, "sortMethod", 0);
    const emscripten::val list = js_pull_array(sort, "conditions");
    const uint32_t n = js_length(list);
    custom_lists.resize(n);
    conditions.resize(n);
    for (uint32_t i = 0; i < n; ++i) {
      pull_sort_condition(list[i], custom_lists[i], conditions[i]);
    }
    f.conditions = conditions.empty() ? nullptr : conditions.data();
    f.condition_count = n;
  }
  return status_from_rc(ops.set(wb, index, &f));
}

JsStatus run_auto_filter(fm_workbook_t* wb, fm_status_t (*fn)(fm_workbook_t*, size_t), uint32_t index) {
  return wb == nullptr ? error_status(kBindingInvalidHandle) : status_from_rc(fn(wb, index));
}

emscripten::val evaluate_auto_filter(const fm_workbook_t* wb, const AutoFilterOps& ops, uint32_t index) {
  emscripten::val o = emscripten::val::object();
  o.set("firstRow", 0U);
  o.set("match", emscripten::val::array());
  if (wb == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  size_t len = 0;
  uint32_t first_row = 0;
  fm_status_t rc = ops.evaluate(wb, index, nullptr, 0, &len, &first_row);
  std::vector<uint8_t> flags(len);
  if (rc == 0 && len > 0) {
    rc = ops.evaluate(wb, index, flags.data(), flags.size(), &len, &first_row);
  }
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  emscripten::val match = emscripten::val::array();
  for (size_t i = 0; i < flags.size(); ++i) {
    match.call<void>("push", flags[i] != 0);
  }
  o.set("status", ok_status());
  o.set("firstRow", first_row);
  o.set("match", match);
  return o;
}

}  // namespace

emscripten::val JsWorkbook::getAutoFilter(uint32_t sheet) const {
  return get_auto_filter(handle_, kSheetAutoFilter, sheet);
}

JsStatus JsWorkbook::setAutoFilter(uint32_t sheet, emscripten::val filter) {
  return set_auto_filter(handle_, kSheetAutoFilter, sheet, filter);
}

JsStatus JsWorkbook::removeAutoFilter(uint32_t sheet) {
  return run_auto_filter(handle_, fm_sheet_remove_auto_filter, sheet);
}

JsStatus JsWorkbook::applyAutoFilter(uint32_t sheet) {
  return run_auto_filter(handle_, fm_sheet_apply_auto_filter, sheet);
}

JsStatus JsWorkbook::clearAutoFilter(uint32_t sheet) {
  return run_auto_filter(handle_, fm_sheet_clear_auto_filter, sheet);
}

emscripten::val JsWorkbook::evaluateAutoFilter(uint32_t sheet) const {
  return evaluate_auto_filter(handle_, kSheetAutoFilter, sheet);
}

emscripten::val JsWorkbook::getTableAutoFilter(uint32_t table) const {
  return get_auto_filter(handle_, kTableAutoFilter, table);
}

JsStatus JsWorkbook::setTableAutoFilter(uint32_t table, emscripten::val filter) {
  return set_auto_filter(handle_, kTableAutoFilter, table, filter);
}

JsStatus JsWorkbook::removeTableAutoFilter(uint32_t table) {
  return run_auto_filter(handle_, fm_table_remove_auto_filter, table);
}

JsStatus JsWorkbook::applyTableAutoFilter(uint32_t table) {
  return run_auto_filter(handle_, fm_table_apply_auto_filter, table);
}

JsStatus JsWorkbook::clearTableAutoFilter(uint32_t table) {
  return run_auto_filter(handle_, fm_table_clear_auto_filter, table);
}

emscripten::val JsWorkbook::evaluateTableAutoFilter(uint32_t table) const {
  return evaluate_auto_filter(handle_, kTableAutoFilter, table);
}

// ---- Threaded comments and persons -------------------------------------

namespace {

/// Owns the strings behind a list of `fm_mention`.
struct MentionStore {
  std::vector<std::string> person_ids;
  std::vector<std::string> mention_ids;
  std::vector<fm_mention> items;
};

void pull_mentions(const emscripten::val& list, MentionStore& st) {
  const uint32_t n = (list.isArray() ? js_length(list) : 0U);
  st.person_ids.resize(n);
  st.mention_ids.resize(n);
  st.items.resize(n);
  for (uint32_t i = 0; i < n; ++i) {
    const emscripten::val m = list[i];
    st.person_ids[i] = js_pull_string(m, "personId");
    st.mention_ids[i] = js_pull_string(m, "mentionId");
    st.items[i] = {st.person_ids[i].c_str(), st.mention_ids[i].c_str(), js_pull_u32(m, "start", 0U),
                   js_pull_u32(m, "length", 0U)};
  }
}

emscripten::val mention_to_val(const fm_mention& m) {
  emscripten::val o = emscripten::val::object();
  o.set("start", m.start);
  o.set("length", m.length);
  js_set_cstr_fields(o, {{"personId", m.person_id}, {"mentionId", m.mention_id}});
  return o;
}

emscripten::val threaded_comment_to_val(const fm_threaded_comment& c) {
  emscripten::val o = emscripten::val::object();
  o.set("row", c.row);
  o.set("col", c.col);
  o.set("done", c.done != 0);
  js_set_cstr_fields(
      o,
      {{"id", c.id}, {"personId", c.person_id}, {"created", c.created}, {"text", c.text}, {"parentId", c.parent_id}});
  emscripten::val mentions = emscripten::val::array();
  for (uint32_t i = 0; i < c.mention_count; ++i) {
    mentions.set(i, mention_to_val(c.mentions[i]));
  }
  o.set("mentions", mentions);
  return o;
}

emscripten::val person_to_val(const fm_person& p) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr_fields(
      o, {{"id", p.id}, {"displayName", p.display_name}, {"userId", p.user_id}, {"providerId", p.provider_id}});
  return o;
}

}  // namespace

emscripten::val JsWorkbook::getThreadedComments(uint32_t sheet) const {
  emscripten::val o = emscripten::val::object();
  emscripten::val comments = emscripten::val::array();
  o.set("comments", comments);
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  size_t count = 0;
  fm_status_t rc = fm_sheet_threaded_comment_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_threaded_comment c{};
    rc = fm_sheet_threaded_comment_at(handle_, sheet, i, &c);
    if (rc == 0) {
      comments.call<void>("push", threaded_comment_to_val(c));
    }
  }
  o.set("status", status_from_rc(rc));
  return o;
}

JsStatus JsWorkbook::addThreadedComment(uint32_t sheet, emscripten::val comment) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  const std::string id = js_pull_string(comment, "id");
  const std::string person_id = js_pull_string(comment, "personId");
  const std::string created = js_pull_string(comment, "created");
  const std::string text = js_pull_string(comment, "text");
  const std::string parent_id = js_pull_string(comment, "parentId");
  MentionStore mentions;
  pull_mentions(comment["mentions"], mentions);
  fm_threaded_comment c{};
  c.id = id.c_str();
  c.row = js_pull_u32(comment, "row", 0U);
  c.col = js_pull_u32(comment, "col", 0U);
  c.person_id = person_id.c_str();
  c.created = created.c_str();
  c.text = text.c_str();
  c.parent_id = parent_id.c_str();
  c.done = pull_flag(comment, "done");
  c.mentions = mentions.items.empty() ? nullptr : mentions.items.data();
  c.mention_count = static_cast<uint32_t>(mentions.items.size());
  return status_from_rc(fm_sheet_add_threaded_comment(handle_, sheet, &c));
}

JsStatus JsWorkbook::editThreadedComment(uint32_t sheet, const std::string& id, const std::string& text,
                                         emscripten::val mentions) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  MentionStore store;
  pull_mentions(mentions, store);
  return status_from_rc(fm_sheet_edit_threaded_comment(handle_, sheet, id.c_str(), text.c_str(),
                                                       store.items.empty() ? nullptr : store.items.data(),
                                                       static_cast<uint32_t>(store.items.size())));
}

JsStatus JsWorkbook::setThreadResolved(uint32_t sheet, const std::string& threadId, bool done) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_set_thread_resolved(handle_, sheet, threadId.c_str(), done ? 1 : 0));
}

JsStatus JsWorkbook::removeThreadedComment(uint32_t sheet, const std::string& id) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_remove_threaded_comment(handle_, sheet, id.c_str()));
}

emscripten::val JsWorkbook::getPersons() const {
  emscripten::val o = emscripten::val::object();
  emscripten::val persons = emscripten::val::array();
  o.set("persons", persons);
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  size_t count = 0;
  fm_status_t rc = fm_workbook_person_count(handle_, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_person p{};
    rc = fm_workbook_person_at(handle_, i, &p);
    if (rc == 0) {
      persons.call<void>("push", person_to_val(p));
    }
  }
  o.set("status", status_from_rc(rc));
  return o;
}

JsStatus JsWorkbook::addPerson(emscripten::val person) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  const std::string id = js_pull_string(person, "id");
  const std::string display_name = js_pull_string(person, "displayName");
  const std::string user_id = js_pull_string(person, "userId");
  const std::string provider_id = js_pull_string(person, "providerId");
  const fm_person p = {id.c_str(), display_name.c_str(), user_id.c_str(), provider_id.c_str()};
  return status_from_rc(fm_workbook_add_person(handle_, &p));
}

JsStatus JsWorkbook::removePerson(const std::string& id) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_workbook_remove_person(handle_, id.c_str()));
}

// ---- Drawing images ----------------------------------------------------

namespace {

/// Envelope for the image calls: `status` plus the header info, zeroed on failure.
emscripten::val image_info_to_val(fm_status_t rc, const fm_image_info& info) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("format", rc == 0 ? info.format : 0);
  o.set("pxWidth", rc == 0 ? info.px_width : 0U);
  o.set("pxHeight", rc == 0 ? info.px_height : 0U);
  return o;
}

emscripten::val drawing_object_to_val(const fm_drawing_object& d) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", d.object_id);
  o.set("kind", d.kind);
  o.set("anchorKind", d.anchor_kind);
  o.set("editAs", d.edit_as);
  o.set("fromRow", d.from_row);
  o.set("fromCol", d.from_col);
  o.set("fromRowOff", static_cast<double>(d.from_row_off));
  o.set("fromColOff", static_cast<double>(d.from_col_off));
  o.set("toRow", d.to_row);
  o.set("toCol", d.to_col);
  o.set("toRowOff", static_cast<double>(d.to_row_off));
  o.set("toColOff", static_cast<double>(d.to_col_off));
  o.set("cx", static_cast<double>(d.cx));
  o.set("cy", static_cast<double>(d.cy));
  o.set("imageFormat", d.image_format);
  js_set_cstr_fields(o, {{"name", d.name}, {"descr", d.descr}, {"mediaPath", d.media_path}});
  return o;
}

}  // namespace

emscripten::val JsWorkbook::probeImage(emscripten::val bytes) const {
  if (handle_ == nullptr) {
    return image_info_to_val(kBindingInvalidHandle, fm_image_info{});
  }
  const std::vector<uint8_t> data = val_to_bytes(bytes);
  fm_image_info info{};
  const fm_status_t rc = fm_workbook_probe_image(handle_, data.data(), data.size(), &info);
  return image_info_to_val(rc, info);
}

emscripten::val JsWorkbook::listDrawingObjects(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(kBindingInvalidHandle));
    return arr;
  }
  size_t count = 0;
  fm_status_t rc = fm_sheet_drawing_object_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_drawing_object d{};
    rc = fm_sheet_drawing_object_at(handle_, sheet, i, &d);
    if (rc == 0) {
      arr.call<void>("push", drawing_object_to_val(d));
    }
  }
  arr.set("status", status_from_rc(rc));
  return arr;
}

emscripten::val JsWorkbook::getImage(uint32_t sheet, uint32_t objectId) const {
  emscripten::val o;
  if (handle_ == nullptr) {
    o = image_info_to_val(kBindingInvalidHandle, fm_image_info{});
    o.set("bytes", bytes_to_val(nullptr, 0));
    return o;
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  fm_image_info info{};
  const fm_status_t rc = fm_sheet_get_image(handle_, sheet, objectId, &data, &len, &info);
  o = image_info_to_val(rc, info);
  o.set("bytes", bytes_to_val(rc == 0 ? data : nullptr, rc == 0 ? len : 0));
  return o;
}

emscripten::val JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", 0U);
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  const std::vector<uint8_t> data = val_to_bytes(bytes);
  std::string name_store;
  std::string descr_store;
  fm_image_insert ins{};
  ins.name = js_pull_optional_string(opts, "name", name_store);
  ins.descr = js_pull_optional_string(opts, "descr", descr_store);
  ins.anchor_kind = js_pull_i32(opts, "anchorKind", FM_ANCHOR_KIND_ONE_CELL);
  ins.edit_as = js_pull_i32(opts, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL);
  ins.row = js_pull_u32(opts, "row", 0U);
  ins.col = js_pull_u32(opts, "col", 0U);
  ins.row_off_emu = static_cast<int64_t>(js_pull_double(opts, "rowOffEmu", 0.0));
  ins.col_off_emu = static_cast<int64_t>(js_pull_double(opts, "colOffEmu", 0.0));
  ins.width_emu = static_cast<int64_t>(js_pull_double(opts, "widthEmu", 0.0));
  ins.height_emu = static_cast<int64_t>(js_pull_double(opts, "heightEmu", 0.0));
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_insert_image(handle_, sheet, data.data(), data.size(), &ins, &id);
  o.set("status", status_from_rc(rc));
  o.set("objectId", rc == 0 ? id : 0U);
  return o;
}

JsStatus JsWorkbook::removeImage(uint32_t sheet, uint32_t objectId) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_remove_image(handle_, sheet, objectId));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
