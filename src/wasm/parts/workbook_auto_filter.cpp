//
// JsWorkbook AutoFilter surface: the raw `<autoFilter>` XML accessors and
// the typed AutoFilter API, written once over an `AutoFilterOps` table and
// exposed for both sheets and tables.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/binding_message.h"
#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- AutoFilter --------------------------------------------------------

emscripten::val JsWorkbook::getSheetAutoFilterXml(uint32_t sheet) const {
  const char* xml = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_auto_filter_xml(handle_, sheet, &xml) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return cstr_or_null(xml);
}

void JsWorkbook::setSheetAutoFilterXml(uint32_t sheet, const std::string& xml) {
  succeeded(handle_ != nullptr ? fm_sheet_set_auto_filter_xml(handle_, sheet, xml.c_str()) : kBindingInvalidHandle);
}

// ---- Typed AutoFilter --------------------------------------------------

namespace {

/// The five AutoFilter entry points, once for sheets and once for tables, so
/// every JS method below is written a single time. A table is addressed by
/// its sheet and sheet-local index; the sheet entry points ignore `table`.
struct AutoFilterOps {
  const char* label;  ///< Public `Class.method` that reports argument failures of the set call.
  fm_status_t (*get)(const fm_workbook_t*, size_t, size_t, int32_t*, fm_auto_filter_t*);
  fm_status_t (*set)(fm_workbook_t*, size_t, size_t, const fm_auto_filter_t*);
  fm_status_t (*apply)(fm_workbook_t*, size_t, size_t);
  fm_status_t (*clear)(fm_workbook_t*, size_t, size_t);
  fm_status_t (*evaluate)(const fm_workbook_t*, size_t, size_t, uint8_t*, size_t, size_t*, uint32_t*);
};

fm_status_t sheet_get(const fm_workbook_t* wb, size_t sheet, size_t, int32_t* present, fm_auto_filter_t* out) {
  return fm_sheet_get_auto_filter(wb, sheet, present, out);
}
fm_status_t sheet_set(fm_workbook_t* wb, size_t sheet, size_t, const fm_auto_filter_t* filter) {
  return fm_sheet_set_auto_filter(wb, sheet, filter);
}
fm_status_t sheet_apply(fm_workbook_t* wb, size_t sheet, size_t) {
  return fm_sheet_apply_auto_filter(wb, sheet);
}
fm_status_t sheet_clear(fm_workbook_t* wb, size_t sheet, size_t) {
  return fm_sheet_clear_auto_filter_criteria(wb, sheet);
}
fm_status_t sheet_evaluate(const fm_workbook_t* wb, size_t sheet, size_t, uint8_t* buf, size_t cap, size_t* count,
                           uint32_t* first_row) {
  return fm_sheet_evaluate_auto_filter(wb, sheet, buf, cap, count, first_row);
}

constexpr AutoFilterOps kSheetAutoFilter = {"Sheet.setAutoFilter", sheet_get,   sheet_set,
                                            sheet_apply,           sheet_clear, sheet_evaluate};
constexpr AutoFilterOps kTableAutoFilter = {"Table.setAutoFilter",
                                            fm_sheet_get_table_auto_filter,
                                            fm_sheet_set_table_auto_filter,
                                            fm_sheet_apply_table_auto_filter,
                                            fm_sheet_clear_table_auto_filter_criteria,
                                            fm_sheet_evaluate_table_auto_filter};

emscripten::val date_group_to_val(const fm_date_group_item_t& d) {
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

emscripten::val filter_column_to_val(const fm_filter_column_t& c) {
  emscripten::val o = emscripten::val::object();
  o.set("columnOffset", c.column_offset);
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
  emscripten::val conditions = emscripten::val::array();
  for (uint32_t i = 0; i < c.condition_count; ++i) {
    emscripten::val condition = emscripten::val::object();
    condition.set("operator", c.conditions[i].op);
    condition.set("value", std::string(c.conditions[i].value != nullptr ? c.conditions[i].value : ""));
    conditions.set(i, condition);
  }
  o.set("conditions", conditions);
  o.set("top", c.top != 0);
  o.set("percent", c.percent != 0);
  o.set("topValue", c.top_val);
  js_set_present_or_null(o, "filterValue", c.has_filter_val, c.filter_val);
  o.set("dynamicType", c.dynamic_type);
  js_set_present_or_null(o, "dynamicValue", c.has_dyn_val, c.dyn_val);
  js_set_present_or_null(o, "dynamicMaxValue", c.has_dyn_max_val, c.dyn_max_val);
  js_set_present_or_null(o, "dxfIndex", c.has_dxf_index, c.dxf_index);
  o.set("cellColor", c.cell_color != 0);
  o.set("iconSet", c.icon_set);
  js_set_present_or_null(o, "iconIndex", c.has_icon_index, c.icon_index);
  js_set_cstr_fields(o, {{"valueIso", c.val_iso}, {"maxValueIso", c.max_val_iso}});
  return o;
}

emscripten::val sort_condition_to_val(const fm_sort_condition_t& c) {
  emscripten::val o = emscripten::val::object();
  o.set("ref", cell_rect_to_val(c.ref));
  o.set("descending", c.descending != 0);
  o.set("sortBy", c.sort_by);
  js_set_present_or_null(o, "dxfIndex", c.has_dxf_index, c.dxf_index);
  o.set("iconSet", c.icon_set);
  js_set_present_or_null(o, "iconIndex", c.has_icon_index, c.icon_index);
  js_set_cstr(o, "customList", c.custom_list);
  return o;
}

emscripten::val auto_filter_to_val(const fm_auto_filter_t& f) {
  emscripten::val o = emscripten::val::object();
  o.set("range", cell_rect_to_val(f.range));
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
  sort.set("ref", cell_rect_to_val(f.sort_ref));
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

/// Owns the strings and arrays one `fm_filter_column_t` points into.
struct FilterColumnStore {
  std::vector<std::string> values;
  std::vector<const char*> value_ptrs;
  std::vector<fm_date_group_item_t> date_groups;
  std::vector<std::string> condition_values;
  std::vector<fm_filter_condition_t> conditions;
  std::string val_iso;
  std::string max_val_iso;
};

// Every reader below takes `path`, the path of `v` under the filter record (`columns[2]`, `sort`), and names a
// rejected field by it.

/// Reads the JS array `v[key]`; absent or non-array reads as empty. Property
/// access and Array.isArray both go through the reader's JS try/catch path.
emscripten::val js_pull_array(const emscripten::val& v, std::string_view path, const char* key,
                              JsNarrowNumericReader& reader) {
  if (!js_value_present(v)) {
    return emscripten::val::array();
  }
  const std::string field = binding::field_path(path, key);
  const emscripten::val a = reader.value(v, key, field.c_str());
  if (!reader.ok() || !js_value_present(a)) {
    return emscripten::val::array();
  }
  return reader.is_array(a, field.c_str()) ? a : emscripten::val::array();
}

fm_date_group_item_t pull_date_group(const emscripten::val& v, std::string_view path, JsNarrowNumericReader& reader) {
  const auto at = [&](const char* key) { return binding::field_path(path, key); };
  fm_date_group_item_t d{};
  d.year = reader.u16(v, "year", 0U, at("year").c_str());
  d.month = reader.u8(v, "month", 0U, at("month").c_str());
  d.day = reader.u8(v, "day", 0U, at("day").c_str());
  d.hour = reader.u8(v, "hour", 0U, at("hour").c_str());
  d.minute = reader.u8(v, "minute", 0U, at("minute").c_str());
  d.second = reader.u8(v, "second", 0U, at("second").c_str());
  d.grouping = reader.u8(v, "grouping", 0U, at("grouping").c_str());
  return d;
}

int32_t pull_flag(const emscripten::val& v, std::string_view path, const char* key, JsNarrowNumericReader& reader,
                  bool dflt = false) {
  return reader.boolean(v, key, dflt, binding::field_path(path, key).c_str()) ? 1 : 0;
}

void pull_filter_column(const emscripten::val& v, std::string_view path, FilterColumnStore& st, fm_filter_column_t& c,
                        JsNarrowNumericReader& reader) {
  const auto at = [&](const char* key) { return binding::field_path(path, key); };
  c = fm_filter_column_t{};
  c.column_offset = reader.u32(v, "columnOffset", 0U, at("columnOffset").c_str());
  if (!reader.ok()) {
    return;
  }
  c.hidden_button = pull_flag(v, path, "hiddenButton", reader);
  c.show_button = pull_flag(v, path, "showButton", reader, true);
  if (!reader.ok()) {
    return;
  }
  c.kind = reader.i32(v, "kind", 0, at("kind").c_str());
  if (!reader.ok()) {
    return;
  }
  c.filter_blank = pull_flag(v, path, "filterBlank", reader);
  if (!reader.ok()) {
    return;
  }
  const emscripten::val values = js_pull_array(v, path, "values", reader);
  const uint32_t value_count = reader.length(values, at("values").c_str());
  if (!reader.ok()) {
    return;
  }
  st.values.reserve(value_count);
  for (uint32_t i = 0; i < value_count; ++i) {
    const std::string element = binding::element_path(at("values"), i);
    const emscripten::val value = reader.array_element(values, i, element.c_str());
    if (!reader.ok()) {
      return;
    }
    st.values.push_back(reader.string_value(value, element.c_str()));
    if (!reader.ok()) {
      return;
    }
  }
  for (const std::string& s : st.values) {
    st.value_ptrs.push_back(s.c_str());
  }
  c.values = st.value_ptrs.empty() ? nullptr : st.value_ptrs.data();
  c.value_count = static_cast<uint32_t>(st.value_ptrs.size());
  const emscripten::val groups = js_pull_array(v, path, "dateGroups", reader);
  const uint32_t group_count = reader.length(groups, at("dateGroups").c_str());
  if (!reader.ok()) {
    return;
  }
  for (uint32_t i = 0; i < group_count; ++i) {
    if (!reader.ok()) {
      return;
    }
    const std::string element = binding::element_path(at("dateGroups"), i);
    const emscripten::val group = reader.array_element(groups, i, element.c_str());
    if (!reader.ok()) {
      return;
    }
    if (!js_value_present(group)) {
      continue;
    }
    st.date_groups.push_back(pull_date_group(group, element, reader));
  }
  if (!reader.ok()) {
    return;
  }
  c.date_groups = st.date_groups.empty() ? nullptr : st.date_groups.data();
  c.date_group_count = static_cast<uint32_t>(st.date_groups.size());
  c.custom_and = pull_flag(v, path, "customAnd", reader);
  const emscripten::val conditions = js_pull_array(v, path, "conditions", reader);
  const uint32_t condition_count = reader.length(conditions, at("conditions").c_str());
  if (!reader.ok()) {
    return;
  }
  st.conditions.reserve(condition_count);
  st.condition_values.reserve(condition_count);
  for (uint32_t i = 0; i < condition_count; ++i) {
    const std::string element = binding::element_path(at("conditions"), i);
    const emscripten::val condition = reader.array_element(conditions, i, element.c_str());
    if (!reader.ok()) {
      return;
    }
    fm_filter_condition_t entry{};
    entry.op = reader.i32(condition, "operator", 0, binding::field_path(element, "operator").c_str());
    st.condition_values.push_back(reader.string(condition, "value", binding::field_path(element, "value").c_str()));
    if (!reader.ok()) {
      return;
    }
    st.conditions.push_back(entry);
  }
  // The value pointers are taken after both vectors stop growing.
  for (uint32_t i = 0; i < condition_count; ++i) {
    st.conditions[i].value = st.condition_values[i].c_str();
  }
  c.conditions = st.conditions.empty() ? nullptr : st.conditions.data();
  c.condition_count = condition_count;
  st.val_iso = reader.string(v, "valueIso", at("valueIso").c_str());
  st.max_val_iso = reader.string(v, "maxValueIso", at("maxValueIso").c_str());
  if (!reader.ok()) {
    return;
  }
  c.val_iso = st.val_iso.c_str();
  c.max_val_iso = st.max_val_iso.c_str();
  c.top = pull_flag(v, path, "top", reader, true);
  c.percent = pull_flag(v, path, "percent", reader);
  c.top_val = reader.number(v, "topValue", 0.0, at("topValue").c_str());
  const emscripten::val filter_val = reader.value(v, "filterValue", at("filterValue").c_str());
  c.has_filter_val = js_value_present(filter_val) ? 1 : 0;
  c.filter_val = reader.number_value(filter_val, 0.0, at("filterValue").c_str());
  c.dynamic_type = reader.i32(v, "dynamicType", 0, at("dynamicType").c_str());
  if (!reader.ok()) {
    return;
  }
  const emscripten::val dyn_val = reader.value(v, "dynamicValue", at("dynamicValue").c_str());
  c.has_dyn_val = js_value_present(dyn_val) ? 1 : 0;
  c.dyn_val = reader.number_value(dyn_val, 0.0, at("dynamicValue").c_str());
  const emscripten::val dyn_max_val = reader.value(v, "dynamicMaxValue", at("dynamicMaxValue").c_str());
  c.has_dyn_max_val = js_value_present(dyn_max_val) ? 1 : 0;
  c.dyn_max_val = reader.number_value(dyn_max_val, 0.0, at("dynamicMaxValue").c_str());
  const emscripten::val dxf_index = reader.value(v, "dxfIndex", at("dxfIndex").c_str());
  c.has_dxf_index = js_value_present(dxf_index) ? 1 : 0;
  c.dxf_index = reader.u32_value(dxf_index, 0U, at("dxfIndex").c_str());
  if (!reader.ok()) {
    return;
  }
  c.cell_color = pull_flag(v, path, "cellColor", reader);
  c.icon_set = reader.i32(v, "iconSet", 0, at("iconSet").c_str());
  const emscripten::val icon_index = reader.value(v, "iconIndex", at("iconIndex").c_str());
  c.has_icon_index = js_value_present(icon_index) ? 1 : 0;
  c.icon_index = reader.i32_value(icon_index, 0, at("iconIndex").c_str());
}

void pull_sort_condition(const emscripten::val& v, std::string_view path, std::string& custom_list,
                         fm_sort_condition_t& c, JsNarrowNumericReader& reader) {
  const auto at = [&](const char* key) { return binding::field_path(path, key); };
  c = fm_sort_condition_t{};
  c.ref = js_pull_range(reader.value(v, "ref", at("ref").c_str()), &reader, at("ref"));
  if (!reader.ok()) {
    return;
  }
  c.descending = pull_flag(v, path, "descending", reader);
  c.sort_by = reader.i32(v, "sortBy", 0, at("sortBy").c_str());
  if (!reader.ok()) {
    return;
  }
  custom_list = reader.string(v, "customList", at("customList").c_str());
  c.custom_list = custom_list.c_str();
  const emscripten::val dxf_index = reader.value(v, "dxfIndex", at("dxfIndex").c_str());
  c.has_dxf_index = js_value_present(dxf_index) ? 1 : 0;
  c.dxf_index = reader.u32_value(dxf_index, 0U, at("dxfIndex").c_str());
  c.icon_set = reader.i32(v, "iconSet", 0, at("iconSet").c_str());
  const emscripten::val icon_index = reader.value(v, "iconIndex", at("iconIndex").c_str());
  c.has_icon_index = js_value_present(icon_index) ? 1 : 0;
  c.icon_index = reader.i32_value(icon_index, 0, at("iconIndex").c_str());
}

emscripten::val get_auto_filter(const fm_workbook_t* wb, const AutoFilterOps& ops, uint32_t sheet, uint32_t table) {
  fm_auto_filter_t f = FM_STRUCT_INIT(fm_auto_filter_t);
  int32_t present = 0;
  const fm_status_t rc = wb != nullptr ? ops.get(wb, sheet, table, &present, &f) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? auto_filter_to_val(f) : emscripten::val::null());
}

void set_auto_filter(fm_workbook_t* wb, const AutoFilterOps& ops, uint32_t sheet, uint32_t table,
                     const emscripten::val& filter) {
  if (wb == nullptr) {
    record_invalid_handle();
    return;
  }
  // `null` removes the filter.
  if (filter.isNull()) {
    succeeded(ops.set(wb, sheet, table, nullptr));
    return;
  }
  JsNarrowNumericReader reader(ops.label);
  fm_auto_filter_t f = FM_STRUCT_INIT(fm_auto_filter_t);
  f.range = js_pull_range(reader.value(filter, "range", "range"), &reader, "range");
  if (!reader.ok()) {
    return;
  }
  const emscripten::val columns = js_pull_array(filter, {}, "columns", reader);
  const uint32_t column_count = reader.length(columns, "columns");
  if (!reader.ok()) {
    return;
  }
  std::vector<FilterColumnStore> stores(column_count);
  std::vector<fm_filter_column_t> cols(column_count);
  for (uint32_t i = 0; i < column_count; ++i) {
    if (!reader.ok()) {
      break;
    }
    const std::string element = binding::element_path("columns", i);
    const emscripten::val column = reader.array_element(columns, i, element.c_str());
    if (!reader.ok()) {
      break;
    }
    if (!js_value_present(column)) {
      continue;
    }
    pull_filter_column(column, element, stores[i], cols[i], reader);
  }
  if (!reader.ok()) {
    return;
  }
  f.columns = cols.empty() ? nullptr : cols.data();
  f.column_count = column_count;
  const emscripten::val sort = reader.value(filter, "sort", "sort");
  std::vector<std::string> custom_lists;
  std::vector<fm_sort_condition_t> conditions;
  if (js_value_present(sort)) {
    f.has_sort = 1;
    f.sort_ref = js_pull_range(reader.value(sort, "ref", "sort.ref"), &reader, "sort.ref");
    if (!reader.ok()) {
      return;
    }
    f.column_sort = pull_flag(sort, "sort", "columnSort", reader);
    f.case_sensitive = pull_flag(sort, "sort", "caseSensitive", reader);
    f.sort_method = reader.i32(sort, "sortMethod", 0, "sort.sortMethod");
    if (!reader.ok()) {
      return;
    }
    const emscripten::val list = js_pull_array(sort, "sort", "conditions", reader);
    const uint32_t n = reader.length(list, "sort.conditions");
    if (!reader.ok()) {
      return;
    }
    custom_lists.resize(n);
    conditions.resize(n);
    for (uint32_t i = 0; i < n; ++i) {
      if (!reader.ok()) {
        break;
      }
      const std::string element = binding::element_path("sort.conditions", i);
      const emscripten::val condition = reader.array_element(list, i, element.c_str());
      if (!reader.ok()) {
        break;
      }
      if (!js_value_present(condition)) {
        continue;
      }
      pull_sort_condition(condition, element, custom_lists[i], conditions[i], reader);
    }
    if (!reader.ok()) {
      return;
    }
    f.conditions = conditions.empty() ? nullptr : conditions.data();
    f.condition_count = n;
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(ops.set(wb, sheet, table, &f));
}

void run_auto_filter(fm_workbook_t* wb, fm_status_t (*fn)(fm_workbook_t*, size_t, size_t), uint32_t sheet,
                     uint32_t table) {
  succeeded(wb != nullptr ? fn(wb, sheet, table) : kBindingInvalidHandle);
}

emscripten::val evaluate_auto_filter(const fm_workbook_t* wb, const AutoFilterOps& ops, uint32_t sheet,
                                     uint32_t table) {
  if (wb == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  size_t len = 0;
  uint32_t first_row = 0;
  fm_status_t rc = ops.evaluate(wb, sheet, table, nullptr, 0, &len, &first_row);
  std::vector<uint8_t> flags(len);
  if (rc == 0 && len > 0) {
    rc = ops.evaluate(wb, sheet, table, flags.data(), flags.size(), &len, &first_row);
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val match = emscripten::val::array();
  for (size_t i = 0; i < flags.size(); ++i) {
    match.call<void>("push", flags[i] != 0);
  }
  emscripten::val o = emscripten::val::object();
  o.set("firstRow", first_row);
  o.set("match", match);
  return o;
}

}  // namespace

emscripten::val JsWorkbook::getAutoFilter(uint32_t sheet) const {
  return get_auto_filter(handle_, kSheetAutoFilter, sheet, 0U);
}

void JsWorkbook::setAutoFilter(uint32_t sheet, emscripten::val filter) {
  set_auto_filter(handle_, kSheetAutoFilter, sheet, 0U, filter);
}

void JsWorkbook::applyAutoFilter(uint32_t sheet) {
  run_auto_filter(handle_, kSheetAutoFilter.apply, sheet, 0U);
}

void JsWorkbook::clearAutoFilter(uint32_t sheet) {
  run_auto_filter(handle_, kSheetAutoFilter.clear, sheet, 0U);
}

emscripten::val JsWorkbook::evaluateAutoFilter(uint32_t sheet) const {
  return evaluate_auto_filter(handle_, kSheetAutoFilter, sheet, 0U);
}

emscripten::val JsWorkbook::getTableAutoFilter(uint32_t sheet, uint32_t table) const {
  return get_auto_filter(handle_, kTableAutoFilter, sheet, table);
}

void JsWorkbook::setTableAutoFilter(uint32_t sheet, uint32_t table, emscripten::val filter) {
  set_auto_filter(handle_, kTableAutoFilter, sheet, table, filter);
}

void JsWorkbook::applyTableAutoFilter(uint32_t sheet, uint32_t table) {
  run_auto_filter(handle_, kTableAutoFilter.apply, sheet, table);
}

void JsWorkbook::clearTableAutoFilter(uint32_t sheet, uint32_t table) {
  run_auto_filter(handle_, kTableAutoFilter.clear, sheet, table);
}

emscripten::val JsWorkbook::evaluateTableAutoFilter(uint32_t sheet, uint32_t table) const {
  return evaluate_auto_filter(handle_, kTableAutoFilter, sheet, table);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
