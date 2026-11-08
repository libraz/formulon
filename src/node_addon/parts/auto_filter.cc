// Typed AutoFilter bindings for sheets and tables: read, evaluate, set
// (a null filter removes), apply and clear.

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- AutoFilter XML fragment ----------------------------------------

Napi::Value Workbook::GetAutoFilterXml(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const char* xml = nullptr;
  const fm_status_t rc = fm_sheet_get_auto_filter_xml(handle_, ArgU32(info, 0), &xml);
  return ValueOrThrow(env, rc, JsStringOrNull(env, xml));
}

Napi::Value Workbook::SetAutoFilterXml(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string xml = ArgString(info, 1);
    return fm_sheet_set_auto_filter_xml(handle, ArgU32(info, 0), xml.c_str());
  });
}

// ---- Typed AutoFilter -----------------------------------------------

namespace {

constexpr uint32_t kAbsentDxfId = UINT32_MAX;

bool PullRange(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, fm_cell_rect_t* out) {
  *out = fm_cell_rect_t{};
  Napi::Object object;
  if (!reader.Object(spec, key, &object)) {
    return reader.ok();
  }
  return ReadCellRect(reader, object, out);
}

Napi::Object FilterColumnToJs(Napi::Env env, const fm_filter_column_t& c) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("columnOffset", JsNumber(env, c.column_offset));
  o.Set("hiddenButton", JsBool(env, c.hidden_button));
  o.Set("showButton", JsBool(env, c.show_button));
  o.Set("kind", JsNumber(env, c.kind));
  o.Set("filterBlank", JsBool(env, c.filter_blank));
  Napi::Array values = Napi::Array::New(env, c.value_count);
  for (uint32_t i = 0; i < c.value_count; ++i) {
    values.Set(i, JsString(env, c.values[i]));
  }
  o.Set("values", values);
  Napi::Array groups = Napi::Array::New(env, c.date_group_count);
  for (uint32_t i = 0; i < c.date_group_count; ++i) {
    const fm_date_group_item_t& g = c.date_groups[i];
    Napi::Object item = Napi::Object::New(env);
    item.Set("year", JsNumber(env, g.year));
    item.Set("month", JsNumber(env, g.month));
    item.Set("day", JsNumber(env, g.day));
    item.Set("hour", JsNumber(env, g.hour));
    item.Set("minute", JsNumber(env, g.minute));
    item.Set("second", JsNumber(env, g.second));
    item.Set("grouping", JsNumber(env, g.grouping));
    groups.Set(i, item);
  }
  o.Set("dateGroups", groups);
  o.Set("customAnd", JsBool(env, c.custom_and));
  Napi::Array conditions = Napi::Array::New(env, c.condition_count);
  for (uint32_t i = 0; i < c.condition_count; ++i) {
    Napi::Object condition = Napi::Object::New(env);
    condition.Set("operator", JsNumber(env, c.conditions[i].op));
    condition.Set("value", JsString(env, c.conditions[i].value));
    conditions.Set(i, condition);
  }
  o.Set("conditions", conditions);
  o.Set("top", JsBool(env, c.top));
  o.Set("percent", JsBool(env, c.percent));
  o.Set("topValue", JsNumber(env, c.top_val));
  o.Set("filterValue", JsNumberOrNull(env, c.has_filter_val, c.filter_val));
  o.Set("dynamicType", JsNumber(env, c.dynamic_type));
  o.Set("dynamicValue", JsNumberOrNull(env, c.has_dyn_val, c.dyn_val));
  o.Set("dynamicMaxValue", JsNumberOrNull(env, c.has_dyn_max_val, c.dyn_max_val));
  o.Set("valueIso", JsString(env, c.val_iso));
  o.Set("maxValueIso", JsString(env, c.max_val_iso));
  o.Set("dxfIndex", JsNumber(env, c.dxf_index));
  o.Set("cellColor", JsBool(env, c.cell_color));
  o.Set("iconSet", JsNumber(env, c.icon_set));
  o.Set("iconIndex", JsNumberOrNull(env, c.has_icon_index, c.icon_index));
  return o;
}

Napi::Object AutoFilterToJs(Napi::Env env, const fm_auto_filter_t& f) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("range", RangeToJs(env, f.range));
  Napi::Array columns = Napi::Array::New(env, f.column_count);
  for (uint32_t i = 0; i < f.column_count; ++i) {
    columns.Set(i, FilterColumnToJs(env, f.columns[i]));
  }
  o.Set("columns", columns);
  if (f.has_sort == 0) {
    o.Set("sort", env.Null());
    return o;
  }
  Napi::Object sort = Napi::Object::New(env);
  sort.Set("ref", RangeToJs(env, f.sort_ref));
  sort.Set("columnSort", JsBool(env, f.column_sort));
  sort.Set("caseSensitive", JsBool(env, f.case_sensitive));
  sort.Set("sortMethod", JsNumber(env, f.sort_method));
  Napi::Array conditions = Napi::Array::New(env, f.condition_count);
  for (uint32_t i = 0; i < f.condition_count; ++i) {
    const fm_sort_condition_t& c = f.conditions[i];
    Napi::Object item = Napi::Object::New(env);
    item.Set("ref", RangeToJs(env, c.ref));
    item.Set("descending", JsBool(env, c.descending));
    item.Set("sortBy", JsNumber(env, c.sort_by));
    item.Set("customList", JsString(env, c.custom_list));
    item.Set("dxfIndex", JsNumberOrNull(env, c.has_dxf_index, c.dxf_index));
    item.Set("iconSet", JsNumber(env, c.icon_set));
    item.Set("iconIndex", JsNumberOrNull(env, c.has_icon_index, c.icon_index));
    conditions.Set(i, item);
  }
  sort.Set("conditions", conditions);
  o.Set("sort", sort);
  return o;
}

// Owns the strings and arrays an `fm_auto_filter_t` borrows from.
struct AutoFilterInput {
  fm_auto_filter_t filter = FM_STRUCT_INIT(fm_auto_filter_t);
  std::vector<fm_filter_column_t> columns;
  std::vector<fm_sort_condition_t> conditions;
  // A deque never relocates its elements, so the `c_str()` views stay valid.
  std::deque<std::string> strings;
  std::vector<std::vector<const char*>> value_ptrs;
  std::vector<std::vector<fm_date_group_item_t>> groups;
  std::vector<std::vector<fm_filter_condition_t>> condition_lists;

  const char* Keep(std::string s) {
    strings.push_back(std::move(s));
    return strings.back().c_str();
  }
};

void ReadFilterColumn(CheckedSpecReader& reader, const Napi::Object& spec, AutoFilterInput& in, fm_filter_column_t& c) {
  c = fm_filter_column_t{};
  c.column_offset = reader.U32(spec, "columnOffset", 0U);
  c.hidden_button = reader.Bool(spec, "hiddenButton", false) ? 1 : 0;
  c.show_button = reader.Bool(spec, "showButton", true) ? 1 : 0;
  c.kind = reader.I32(spec, "kind", 0);
  c.filter_blank = reader.Bool(spec, "filterBlank", false) ? 1 : 0;
  in.value_ptrs.emplace_back();
  std::vector<const char*>& ptrs = in.value_ptrs.back();
  Napi::Array values;
  if (reader.Array(spec, "values", &values)) {
    for (uint32_t i = 0; i < values.Length(); ++i) {
      Napi::Value value;
      if (!reader.ArrayElement(values, i, &value)) {
        continue;
      }
      std::string text;
      if (!reader.String(value, "values[]", &text)) {
        break;
      }
      ptrs.push_back(in.Keep(std::move(text)));
    }
  }
  c.values = ptrs.empty() ? nullptr : ptrs.data();
  c.value_count = static_cast<uint32_t>(ptrs.size());
  in.groups.emplace_back();
  std::vector<fm_date_group_item_t>& groups = in.groups.back();
  Napi::Array date_groups;
  if (reader.Array(spec, "dateGroups", &date_groups)) {
    const Napi::Array& arr = date_groups;
    for (uint32_t i = 0; i < arr.Length(); ++i) {
      Napi::Value group_value;
      if (!reader.ArrayElement(arr, i, &group_value)) {
        continue;
      }
      Napi::Object g;
      if (!reader.Object(group_value, "dateGroups[]", &g)) {
        break;
      }
      fm_date_group_item_t item{};
      item.year = reader.U16(g, "year", 0U);
      item.month = reader.U8(g, "month", 0U);
      item.day = reader.U8(g, "day", 0U);
      item.hour = reader.U8(g, "hour", 0U);
      item.minute = reader.U8(g, "minute", 0U);
      item.second = reader.U8(g, "second", 0U);
      item.grouping = reader.U8(g, "grouping", 0U);
      if (!reader.ok()) {
        break;
      }
      groups.push_back(item);
    }
  }
  c.date_groups = groups.empty() ? nullptr : groups.data();
  c.date_group_count = static_cast<uint32_t>(groups.size());
  c.custom_and = reader.Bool(spec, "customAnd", false) ? 1 : 0;
  in.condition_lists.emplace_back();
  std::vector<fm_filter_condition_t>& conditions = in.condition_lists.back();
  Napi::Array condition_array;
  if (reader.Array(spec, "conditions", &condition_array)) {
    for (uint32_t i = 0; i < condition_array.Length(); ++i) {
      Napi::Value condition_value;
      if (!reader.ArrayElement(condition_array, i, &condition_value)) {
        continue;
      }
      Napi::Object condition;
      if (!reader.Object(condition_value, "conditions[]", &condition)) {
        break;
      }
      fm_filter_condition_t item{};
      item.op = reader.I32(condition, "operator", 0);
      std::string text;
      reader.String(condition, "value", &text);
      item.value = in.Keep(std::move(text));
      if (!reader.ok()) {
        break;
      }
      conditions.push_back(item);
    }
  }
  c.conditions = conditions.empty() ? nullptr : conditions.data();
  c.condition_count = static_cast<uint32_t>(conditions.size());
  c.top = reader.Bool(spec, "top", true) ? 1 : 0;
  c.percent = reader.Bool(spec, "percent", false) ? 1 : 0;
  // A nullable member is absent when omitted or `null`.
  bool present = false;
  c.top_val = reader.Double(spec, "topValue", 0.0);
  c.filter_val = reader.Double(spec, "filterValue", 0.0, &present);
  c.has_filter_val = present ? 1 : 0;
  c.dynamic_type = reader.I32(spec, "dynamicType", 0);
  c.dyn_val = reader.Double(spec, "dynamicValue", 0.0, &present);
  c.has_dyn_val = present ? 1 : 0;
  c.dyn_max_val = reader.Double(spec, "dynamicMaxValue", 0.0, &present);
  c.has_dyn_max_val = present ? 1 : 0;
  std::string val_iso;
  std::string max_val_iso;
  reader.String(spec, "valueIso", &val_iso);
  reader.String(spec, "maxValueIso", &max_val_iso);
  c.val_iso = in.Keep(std::move(val_iso));
  c.max_val_iso = in.Keep(std::move(max_val_iso));
  c.dxf_index = reader.U32(spec, "dxfIndex", kAbsentDxfId);
  c.cell_color = reader.Bool(spec, "cellColor", false) ? 1 : 0;
  c.icon_set = reader.I32(spec, "iconSet", 0);
  c.icon_index = reader.I32(spec, "iconIndex", 0, &present);
  c.has_icon_index = present ? 1 : 0;
}

void ReadSortCondition(CheckedSpecReader& reader, const Napi::Object& spec, AutoFilterInput& in,
                       fm_sort_condition_t& c) {
  c = fm_sort_condition_t{};
  PullRange(reader, spec, "ref", &c.ref);
  c.descending = reader.Bool(spec, "descending", false) ? 1 : 0;
  c.sort_by = reader.I32(spec, "sortBy", 0);
  std::string custom_list;
  reader.String(spec, "customList", &custom_list);
  c.custom_list = in.Keep(std::move(custom_list));
  // A nullable member is absent when omitted or `null`.
  bool present = false;
  c.dxf_index = reader.U32(spec, "dxfIndex", 0U, &present);
  c.has_dxf_index = present ? 1 : 0;
  c.icon_set = reader.I32(spec, "iconSet", 0);
  c.icon_index = reader.I32(spec, "iconIndex", 0, &present);
  c.has_icon_index = present ? 1 : 0;
}

// Reads the `AutoFilter` record at `info[idx]`; false when it is not an object.
bool ReadAutoFilter(CheckedSpecReader& reader, const Napi::CallbackInfo& info, size_t idx, AutoFilterInput& in) {
  Napi::Object spec;
  if (!ObjectArg(info, idx, &spec)) {
    return false;
  }
  fm_auto_filter_t& f = in.filter;
  PullRange(reader, spec, "range", &f.range);
  Napi::Array columns;
  if (reader.Array(spec, "columns", &columns)) {
    for (uint32_t i = 0; i < columns.Length() && reader.ok(); ++i) {
      Napi::Value column_value;
      if (!reader.ArrayElement(columns, i, &column_value)) {
        continue;
      }
      Napi::Object column;
      if (!reader.Object(column_value, "columns[]", &column)) {
        break;
      }
      in.columns.emplace_back();
      ReadFilterColumn(reader, column, in, in.columns.back());
    }
  }
  Napi::Object sort;
  const bool has_sort = reader.Object(spec, "sort", &sort);
  if (!reader.ok()) {
    return false;
  }
  Napi::Array conditions;
  if (has_sort && reader.Array(sort, "conditions", &conditions)) {
    for (uint32_t i = 0; i < conditions.Length() && reader.ok(); ++i) {
      Napi::Value condition_value;
      if (!reader.ArrayElement(conditions, i, &condition_value)) {
        continue;
      }
      Napi::Object condition;
      if (!reader.Object(condition_value, "conditions[]", &condition)) {
        break;
      }
      in.conditions.emplace_back();
      ReadSortCondition(reader, condition, in, in.conditions.back());
    }
  }
  if (!reader.ok()) {
    return false;
  }
  f.columns = in.columns.empty() ? nullptr : in.columns.data();
  f.column_count = static_cast<uint32_t>(in.columns.size());
  f.has_sort = has_sort ? 1 : 0;
  if (has_sort) {
    PullRange(reader, sort, "ref", &f.sort_ref);
    f.column_sort = reader.Bool(sort, "columnSort", false) ? 1 : 0;
    f.case_sensitive = reader.Bool(sort, "caseSensitive", false) ? 1 : 0;
    f.sort_method = reader.I32(sort, "sortMethod", 0);
  }
  f.conditions = in.conditions.empty() ? nullptr : in.conditions.data();
  f.condition_count = static_cast<uint32_t>(in.conditions.size());
  return reader.ok();
}

// Shared by the sheet and table entry points: `table` selects which C
// function family a call goes to. A sheet call is (sheet, ...); a table call
// is (sheet, table, ...).
Napi::Value GetAutoFilterImpl(const Napi::CallbackInfo& info, const fm_workbook_t* wb, bool table) {
  Napi::Env env = info.Env();
  if (wb == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const size_t sheet = Workbook::ArgU32(info, 0);
  const size_t table_index = table ? Workbook::ArgU32(info, 1) : 0;
  fm_auto_filter_t af = FM_STRUCT_INIT(fm_auto_filter_t);
  int32_t present = 0;
  const fm_status_t rc = table ? fm_sheet_get_table_auto_filter(wb, sheet, table_index, &af, &present)
                               : fm_sheet_get_auto_filter(wb, sheet, &af, &present);
  const bool has = rc == 0 && present != 0;
  return OptionalOrNull(env, rc, present, has ? static_cast<Napi::Value>(AutoFilterToJs(env, af)) : env.Null());
}

Napi::Value EvaluateAutoFilterImpl(const Napi::CallbackInfo& info, const fm_workbook_t* wb, bool table) {
  Napi::Env env = info.Env();
  if (wb == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const size_t sheet = Workbook::ArgU32(info, 0);
  const size_t table_index = table ? Workbook::ArgU32(info, 1) : 0;
  auto call = [&](uint8_t* buf, size_t cap, size_t* count, uint32_t* first) {
    return table ? fm_sheet_evaluate_table_auto_filter(wb, sheet, table_index, buf, cap, count, first)
                 : fm_sheet_evaluate_auto_filter(wb, sheet, buf, cap, count, first);
  };
  size_t len = 0;
  uint32_t first = 0;
  fm_status_t rc = call(nullptr, 0, &len, &first);
  std::vector<uint8_t> flags(len);
  if (rc == 0 && len > 0) {
    rc = call(flags.data(), flags.size(), &len, &first);
  }
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array match = Napi::Array::New(env, flags.size());
  for (size_t i = 0; i < flags.size(); ++i) {
    match.Set(static_cast<uint32_t>(i), Napi::Boolean::New(env, flags[i] != 0));
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("firstRow", JsNumber(env, first));
  out.Set("match", match);
  return out;
}

// `filter` is the record at `info[filter_idx]`, or `null` to remove the filter.
// `handle` aliases the wrapper's member, so the C call sees the handle as it
// is after the spec getters ran, not a copy taken before them.
Napi::Value SetAutoFilterImpl(const Napi::CallbackInfo& info, fm_workbook_t* const& handle, size_t filter_idx,
                              bool table, const char* usage) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const size_t sheet = Workbook::ArgU32(info, 0);
  const size_t table_index = table ? Workbook::ArgU32(info, 1) : 0;
  if (info[filter_idx].IsNull()) {
    return ValueOrThrow(env,
                        table ? fm_sheet_set_table_auto_filter(handle, sheet, table_index, nullptr)
                              : fm_sheet_set_auto_filter(handle, sheet, nullptr),
                        env.Undefined());
  }
  CheckedSpecReader reader(env);
  AutoFilterInput in;
  if (!ReadAutoFilter(reader, info, filter_idx, in)) {
    if (!reader.ok()) {
      return env.Undefined();
    }
    Napi::TypeError::New(env, usage).ThrowAsJavaScriptException();
    return env.Undefined();
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env,
                      table ? fm_sheet_set_table_auto_filter(handle, sheet, table_index, &in.filter)
                            : fm_sheet_set_auto_filter(handle, sheet, &in.filter),
                      env.Undefined());
}

}  // namespace

Napi::Value Workbook::GetAutoFilter(const Napi::CallbackInfo& info) {
  return GetAutoFilterImpl(info, handle_, false);
}

Napi::Value Workbook::GetTableAutoFilter(const Napi::CallbackInfo& info) {
  return GetAutoFilterImpl(info, handle_, true);
}

Napi::Value Workbook::SetAutoFilter(const Napi::CallbackInfo& info) {
  return SetAutoFilterImpl(info, handle_, 1, false, "setAutoFilter expects (sheet:number, autoFilter:object|null)");
}

Napi::Value Workbook::SetTableAutoFilter(const Napi::CallbackInfo& info) {
  return SetAutoFilterImpl(info, handle_, 2, true,
                           "setTableAutoFilter expects (sheet:number, tableIndex:number, autoFilter:object|null)");
}

Napi::Value Workbook::ApplyAutoFilter(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) { return fm_sheet_apply_auto_filter(handle, ArgU32(info, 0)); });
}

Napi::Value Workbook::ApplyTableAutoFilter(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_sheet_apply_table_auto_filter(handle, ArgU32(info, 0), ArgU32(info, 1));
  });
}

Napi::Value Workbook::ClearAutoFilter(const Napi::CallbackInfo& info) {
  return StatusCall(
      info, [&](fm_workbook_t* handle) { return fm_sheet_clear_auto_filter_criteria(handle, ArgU32(info, 0)); });
}

Napi::Value Workbook::ClearTableAutoFilter(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_sheet_clear_table_auto_filter_criteria(handle, ArgU32(info, 0), ArgU32(info, 1));
  });
}

Napi::Value Workbook::EvaluateAutoFilter(const Napi::CallbackInfo& info) {
  return EvaluateAutoFilterImpl(info, handle_, false);
}

Napi::Value Workbook::EvaluateTableAutoFilter(const Napi::CallbackInfo& info) {
  return EvaluateAutoFilterImpl(info, handle_, true);
}

}  // namespace formulon_node
