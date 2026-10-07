// Typed AutoFilter bindings for sheets and tables: read, evaluate, set,
// remove, apply and clear.

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Typed AutoFilter -----------------------------------------------

namespace {

constexpr uint32_t kAbsentDxfId = UINT32_MAX;

bool PullRange(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, fm_merge_range* out) {
  *out = fm_merge_range{};
  Napi::Object object;
  if (!reader.Object(spec, key, &object)) {
    return reader.ok();
  }
  return ReadMergeRange(reader, object, out);
}

Napi::Object FilterColumnToJs(Napi::Env env, const fm_filter_column& c) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("colId", JsNumber(env, c.col_id));
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
    const fm_date_group_item& g = c.date_groups[i];
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
  o.Set("customCount", JsNumber(env, c.custom_count));
  o.Set("op1", JsNumber(env, c.op1));
  o.Set("val1", JsString(env, c.val1));
  o.Set("op2", JsNumber(env, c.op2));
  o.Set("val2", JsString(env, c.val2));
  o.Set("top", JsBool(env, c.top));
  o.Set("percent", JsBool(env, c.percent));
  o.Set("hasFilterVal", JsBool(env, c.has_filter_val));
  o.Set("topVal", JsNumber(env, c.top_val));
  o.Set("filterVal", JsNumber(env, c.filter_val));
  o.Set("dynamicType", JsNumber(env, c.dynamic_type));
  o.Set("hasDynVal", JsBool(env, c.has_dyn_val));
  o.Set("hasDynMaxVal", JsBool(env, c.has_dyn_max_val));
  o.Set("dynVal", JsNumber(env, c.dyn_val));
  o.Set("dynMaxVal", JsNumber(env, c.dyn_max_val));
  o.Set("valIso", JsString(env, c.val_iso));
  o.Set("maxValIso", JsString(env, c.max_val_iso));
  o.Set("dxfId", JsNumber(env, c.dxf_id));
  o.Set("cellColor", JsBool(env, c.cell_color));
  o.Set("iconSet", JsNumber(env, c.icon_set));
  o.Set("iconId", JsNumber(env, c.icon_id));
  o.Set("hasIconId", JsBool(env, c.has_icon_id));
  return o;
}

Napi::Object AutoFilterToJs(Napi::Env env, const fm_auto_filter& f) {
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
    const fm_sort_condition& c = f.conditions[i];
    Napi::Object item = Napi::Object::New(env);
    item.Set("ref", RangeToJs(env, c.ref));
    item.Set("descending", JsBool(env, c.descending));
    item.Set("sortBy", JsNumber(env, c.sort_by));
    item.Set("customList", JsString(env, c.custom_list));
    item.Set("dxfId", JsNumber(env, c.dxf_id));
    item.Set("hasDxfId", JsBool(env, c.has_dxf_id));
    item.Set("iconSet", JsNumber(env, c.icon_set));
    item.Set("iconId", JsNumber(env, c.icon_id));
    item.Set("hasIconId", JsBool(env, c.has_icon_id));
    conditions.Set(i, item);
  }
  sort.Set("conditions", conditions);
  o.Set("sort", sort);
  return o;
}

// Owns the strings and arrays an `fm_auto_filter` borrows from.
struct AutoFilterInput {
  fm_auto_filter filter{};
  std::vector<fm_filter_column> columns;
  std::vector<fm_sort_condition> conditions;
  // A deque never relocates its elements, so the `c_str()` views stay valid.
  std::deque<std::string> strings;
  std::vector<std::vector<const char*>> value_ptrs;
  std::vector<std::vector<fm_date_group_item>> groups;

  const char* Keep(std::string s) {
    strings.push_back(std::move(s));
    return strings.back().c_str();
  }
};

void ReadFilterColumn(CheckedSpecReader& reader, const Napi::Object& spec, AutoFilterInput& in, fm_filter_column& c) {
  c = fm_filter_column{};
  c.col_id = reader.U32(spec, "colId", 0U);
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
  std::vector<fm_date_group_item>& groups = in.groups.back();
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
      fm_date_group_item item{};
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
  c.custom_count = reader.I32(spec, "customCount", 0);
  c.op1 = reader.I32(spec, "op1", 0);
  std::string val1;
  reader.String(spec, "val1", &val1);
  c.val1 = in.Keep(std::move(val1));
  c.op2 = reader.I32(spec, "op2", 0);
  std::string val2;
  reader.String(spec, "val2", &val2);
  c.val2 = in.Keep(std::move(val2));
  c.top = reader.Bool(spec, "top", false) ? 1 : 0;
  c.percent = reader.Bool(spec, "percent", false) ? 1 : 0;
  c.has_filter_val = reader.Bool(spec, "hasFilterVal", false) ? 1 : 0;
  c.top_val = reader.Double(spec, "topVal", 0.0);
  c.filter_val = reader.Double(spec, "filterVal", 0.0);
  c.dynamic_type = reader.I32(spec, "dynamicType", 0);
  c.has_dyn_val = reader.Bool(spec, "hasDynVal", false) ? 1 : 0;
  c.has_dyn_max_val = reader.Bool(spec, "hasDynMaxVal", false) ? 1 : 0;
  c.dyn_val = reader.Double(spec, "dynVal", 0.0);
  c.dyn_max_val = reader.Double(spec, "dynMaxVal", 0.0);
  std::string val_iso;
  std::string max_val_iso;
  reader.String(spec, "valIso", &val_iso);
  reader.String(spec, "maxValIso", &max_val_iso);
  c.val_iso = in.Keep(std::move(val_iso));
  c.max_val_iso = in.Keep(std::move(max_val_iso));
  c.dxf_id = reader.U32(spec, "dxfId", kAbsentDxfId);
  c.cell_color = reader.Bool(spec, "cellColor", false) ? 1 : 0;
  c.icon_set = reader.I32(spec, "iconSet", 0);
  c.icon_id = reader.I32(spec, "iconId", 0);
  c.has_icon_id = reader.Bool(spec, "hasIconId", false) ? 1 : 0;
}

void ReadSortCondition(CheckedSpecReader& reader, const Napi::Object& spec, AutoFilterInput& in, fm_sort_condition& c) {
  c = fm_sort_condition{};
  PullRange(reader, spec, "ref", &c.ref);
  c.descending = reader.Bool(spec, "descending", false) ? 1 : 0;
  c.sort_by = reader.I32(spec, "sortBy", 0);
  std::string custom_list;
  reader.String(spec, "customList", &custom_list);
  c.custom_list = in.Keep(std::move(custom_list));
  c.dxf_id = reader.U32(spec, "dxfId", kAbsentDxfId);
  c.has_dxf_id = reader.Bool(spec, "hasDxfId", false) ? 1 : 0;
  c.icon_set = reader.I32(spec, "iconSet", 0);
  c.icon_id = reader.I32(spec, "iconId", 0);
  c.has_icon_id = reader.Bool(spec, "hasIconId", false) ? 1 : 0;
}

// Reads the `AutoFilter` record at `info[idx]`; false when it is not an object.
bool ReadAutoFilter(CheckedSpecReader& reader, const Napi::CallbackInfo& info, size_t idx, AutoFilterInput& in) {
  if (info.Length() <= idx || !info[idx].IsObject()) {
    return false;
  }
  const Napi::Object spec = info[idx].As<Napi::Object>();
  fm_auto_filter& f = in.filter;
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
// function family a call goes to, `idx` being the sheet or table index.
Napi::Value GetAutoFilterImpl(Napi::Env env, const fm_workbook_t* wb, bool table, size_t idx) {
  fm_auto_filter af{};
  int32_t present = 0;
  const fm_status_t rc =
      table ? fm_table_get_auto_filter(wb, idx, &af, &present) : fm_sheet_get_auto_filter(wb, idx, &af, &present);
  const bool has = rc == 0 && present != 0;
  return MakeFieldResult(env, MakeStatus(env, rc), "autoFilter",
                         has ? static_cast<Napi::Value>(AutoFilterToJs(env, af)) : env.Null());
}

Napi::Value EvaluateAutoFilterImpl(Napi::Env env, const fm_workbook_t* wb, bool table, size_t idx) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("firstRow", JsNumber(env, 0));
  out.Set("match", Napi::Array::New(env));
  auto call = [&](uint8_t* buf, size_t cap, size_t* len, uint32_t* first) {
    return table ? fm_table_evaluate_auto_filter(wb, idx, buf, cap, len, first)
                 : fm_sheet_evaluate_auto_filter(wb, idx, buf, cap, len, first);
  };
  size_t len = 0;
  uint32_t first = 0;
  fm_status_t rc = call(nullptr, 0, &len, &first);
  std::vector<uint8_t> flags(len);
  if (rc == 0 && len > 0) {
    rc = call(flags.data(), flags.size(), &len, &first);
  }
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  Napi::Array match = Napi::Array::New(env, flags.size());
  for (size_t i = 0; i < flags.size(); ++i) {
    match.Set(static_cast<uint32_t>(i), Napi::Boolean::New(env, flags[i] != 0));
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("firstRow", JsNumber(env, first));
  out.Set("match", match);
  return out;
}

// Builds the `{ status, firstRow: 0, match: [] }` result for a dead handle.
Napi::Object EmptyAutoFilterEvaluation(Napi::Env env, Napi::Object status) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("firstRow", JsNumber(env, 0));
  out.Set("match", Napi::Array::New(env));
  return out;
}

// Shared bodies of the sheet / table AutoFilter mutators.
using AutoFilterSetFn = fm_status_t (*)(fm_workbook_t*, size_t, const fm_auto_filter*);
using AutoFilterOpFn = fm_status_t (*)(fm_workbook_t*, size_t);

Napi::Value SetAutoFilterImpl(const Napi::CallbackInfo& info, fm_workbook_t* handle, AutoFilterSetFn fn,
                              const char* usage) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  CheckedSpecReader reader(env);
  AutoFilterInput in;
  if (!ReadAutoFilter(reader, info, 1, in)) {
    if (!reader.ok()) {
      return env.Undefined();
    }
    return MakeBindingArgumentError(env, usage);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fn(handle, Workbook::ArgU32(info, 0), &in.filter));
}

Napi::Value InvokeAutoFilterOp(const Napi::CallbackInfo& info, fm_workbook_t* handle, AutoFilterOpFn fn) {
  Napi::Env env = info.Env();
  return handle == nullptr ? MakeErrorStatus(env, kBindingInvalidHandle)
                           : MakeStatus(env, fn(handle, Workbook::ArgU32(info, 0)));
}

}  // namespace

Napi::Value Workbook::GetAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "autoFilter", env.Null());
  }
  return GetAutoFilterImpl(env, handle_, false, ArgU32(info, 0));
}

Napi::Value Workbook::GetTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "autoFilter", env.Null());
  }
  return GetAutoFilterImpl(env, handle_, true, ArgU32(info, 0));
}

Napi::Value Workbook::SetAutoFilter(const Napi::CallbackInfo& info) {
  return SetAutoFilterImpl(info, handle_, &fm_sheet_set_auto_filter,
                           "setAutoFilter expects (sheet:number, autoFilter:object)");
}

Napi::Value Workbook::SetTableAutoFilter(const Napi::CallbackInfo& info) {
  return SetAutoFilterImpl(info, handle_, &fm_table_set_auto_filter,
                           "setTableAutoFilter expects (tableIdx:number, autoFilter:object)");
}

Napi::Value Workbook::RemoveAutoFilter(const Napi::CallbackInfo& info) {
  return InvokeAutoFilterOp(info, handle_, &fm_sheet_remove_auto_filter);
}

Napi::Value Workbook::RemoveTableAutoFilter(const Napi::CallbackInfo& info) {
  return InvokeAutoFilterOp(info, handle_, &fm_table_remove_auto_filter);
}

Napi::Value Workbook::ApplyAutoFilter(const Napi::CallbackInfo& info) {
  return InvokeAutoFilterOp(info, handle_, &fm_sheet_apply_auto_filter);
}

Napi::Value Workbook::ApplyTableAutoFilter(const Napi::CallbackInfo& info) {
  return InvokeAutoFilterOp(info, handle_, &fm_table_apply_auto_filter);
}

Napi::Value Workbook::ClearAutoFilter(const Napi::CallbackInfo& info) {
  return InvokeAutoFilterOp(info, handle_, &fm_sheet_clear_auto_filter);
}

Napi::Value Workbook::ClearTableAutoFilter(const Napi::CallbackInfo& info) {
  return InvokeAutoFilterOp(info, handle_, &fm_table_clear_auto_filter);
}

Napi::Value Workbook::EvaluateAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return EmptyAutoFilterEvaluation(env, NullHandleError(env));
  }
  return EvaluateAutoFilterImpl(env, handle_, false, ArgU32(info, 0));
}

Napi::Value Workbook::EvaluateTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return EmptyAutoFilterEvaluation(env, NullHandleError(env));
  }
  return EvaluateAutoFilterImpl(env, handle_, true, ArgU32(info, 0));
}

}  // namespace formulon_node
