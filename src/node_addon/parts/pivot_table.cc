// PivotTable mutation bindings: pivot create/remove, anchor / name /
// totals, per-field configuration (axis, sort, items, subtotals,
// date grouping, number format), data fields, filters,
// and the row/column field-order setters.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct PivotFieldArgs {
  std::size_t sheet;
  std::size_t pivot_idx;
  std::size_t field_idx;
};

PivotFieldArgs ReadPivotField(const Napi::CallbackInfo& info) {
  return {static_cast<std::size_t>(Workbook::ArgU32(info, 0)), static_cast<std::size_t>(Workbook::ArgU32(info, 1)),
          static_cast<std::size_t>(Workbook::ArgU32(info, 2))};
}

// Shared body for the entries that take `(sheet, pivot)` and return a count.
using PivotCountFn = fm_status_t (*)(const fm_workbook_t*, size_t, size_t, size_t*);

Napi::Value PivotCountResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle, PivotCountFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(Workbook::ArgU32(info, 1));
  std::size_t count = 0;
  const fm_status_t rc = fn(handle, sheet, pivot_idx, &count);
  return MakeNumberResult(env, rc, static_cast<double>(count));
}

}  // namespace

Napi::Value Workbook::PivotCreate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::string name = ArgString(info, 1);
  const uint32_t cache_id = ArgU32(info, 2);
  const uint32_t anchor_row = ArgU32(info, 3);
  const uint32_t anchor_col = ArgU32(info, 4);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_create(handle_, sheet, name.c_str(), cache_id, anchor_row, anchor_col, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotRemove(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_remove);
}

Napi::Value Workbook::PivotSetName(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
    const std::string name = ArgString(info, 2);
    return fm_workbook_pivot_set_name(handle, sheet, pivot_idx, name.c_str());
  });
}

Napi::Value Workbook::PivotSetAnchor(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_set_anchor);
}

Napi::Value Workbook::PivotSetGrandTotals(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
    const int32_t rows_enabled = ArgFlag(info, 2);
    const int32_t cols_enabled = ArgFlag(info, 3);
    return fm_workbook_pivot_set_grand_totals(handle, sheet, pivot_idx, rows_enabled, cols_enabled);
  });
}

Napi::Value Workbook::PivotGetLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, kBindingInvalidHandle, "layout", 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_layout_t layout = FM_PIVOT_LAYOUT_COMPACT;
  fm_status_t rc = fm_workbook_pivot_get_layout(handle_, sheet, pivot_idx, &layout);
  return MakeNumberFieldResult(env, rc, "layout", static_cast<uint32_t>(layout));
}

Napi::Value Workbook::PivotSetLayout(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_set_layout);
}

Napi::Value Workbook::PivotCacheId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_pivot_get_cache_id(handle_, sheet, pivot_idx, &out);
  return MakeIndexResult(env, rc, out);
}

Napi::Value Workbook::PivotName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringResult(env, kBindingInvalidHandle, nullptr);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const char* name = nullptr;
  const fm_status_t rc = fm_workbook_pivot_get_name(handle_, sheet, pivot_idx, &name);
  return MakeStringResult(env, rc, rc == 0 ? name : nullptr);
}

Napi::Value Workbook::PivotGetGrandTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  int32_t rows = 0;
  int32_t cols = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_pivot_get_grand_totals(handle_, sheet, pivot_idx, &rows, &cols)
                             : kBindingInvalidHandle;
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("rows", Napi::Boolean::New(env, rc == 0 && rows != 0));
  out.Set("cols", Napi::Boolean::New(env, rc == 0 && cols != 0));
  return out;
}

namespace {

using PivotFieldOrderGetter = fm_status_t (*)(const fm_workbook_t*, size_t, size_t, const uint32_t**, size_t*);

Napi::Value GetPivotFieldOrder(const Napi::CallbackInfo& info, const fm_workbook_t* handle, PivotFieldOrderGetter fn) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(Workbook::ArgU32(info, 1));
  const uint32_t* data = nullptr;
  std::size_t count = 0;
  const fm_status_t rc = handle != nullptr ? fn(handle, sheet, pivot_idx, &data, &count) : kBindingInvalidHandle;
  Napi::Array indices = Napi::Array::New(env);
  for (std::size_t i = 0; rc == 0 && i < count; ++i) {
    indices.Set(static_cast<uint32_t>(i), Napi::Number::New(env, data[i]));
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("indices", indices);
  return out;
}

}  // namespace

Napi::Value Workbook::PivotRowFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_workbook_pivot_get_row_field_order);
}

Napi::Value Workbook::PivotColFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_workbook_pivot_get_col_field_order);
}

Napi::Value Workbook::PivotFieldCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_workbook_pivot_field_count);
}

Napi::Value Workbook::PivotFieldAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const PivotFieldArgs args = ReadPivotField(info);
  fm_pivot_field_info_t field{};
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_pivot_field_at(handle_, args.sheet, args.pivot_idx, args.field_idx, &field)
                             : kBindingInvalidHandle;
  if (rc != 0) {
    field = fm_pivot_field_info_t{};
    field.axis = -1;
    field.cache_field_index = -1;
    field.date_start_year = -1;
    field.date_end_year = -1;
    field.date_start_serial = -1.0;
    field.date_end_serial = -1.0;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("sourceName", JsString(env, field.source_name));
  out.Set("customName", JsString(env, field.custom_name));
  out.Set("axis", Napi::Number::New(env, field.axis));
  out.Set("subtotalTop", Napi::Boolean::New(env, field.subtotal_top != 0));
  out.Set("numberFormat", JsString(env, field.number_format));
  out.Set("cacheFieldIndex", Napi::Number::New(env, field.cache_field_index));
  out.Set("sortOrder", Napi::Number::New(env, field.sort_order));
  out.Set("sortByField", JsString(env, field.sort_by_field));
  out.Set("hasDateGroup", Napi::Boolean::New(env, field.has_date_group != 0));
  out.Set("dateGranularity", Napi::Number::New(env, field.date_granularity));
  out.Set("dateCalendar", Napi::Number::New(env, field.date_calendar));
  out.Set("dateStartYear", Napi::Number::New(env, field.date_start_year));
  out.Set("dateEndYear", Napi::Number::New(env, field.date_end_year));
  out.Set("dateIntervalDays", Napi::Number::New(env, field.date_interval_days));
  out.Set("dateStartSerial", Napi::Number::New(env, field.date_start_serial));
  out.Set("dateEndSerial", Napi::Number::New(env, field.date_end_serial));
  return out;
}

Napi::Value Workbook::PivotFieldItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  const PivotFieldArgs args = ReadPivotField(info);
  std::size_t count = 0;
  const fm_status_t rc =
      fm_workbook_pivot_field_item_count(handle_, args.sheet, args.pivot_idx, args.field_idx, &count);
  return MakeNumberResult(env, rc, static_cast<double>(count));
}

Napi::Value Workbook::PivotFieldItemAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const PivotFieldArgs args = ReadPivotField(info);
  const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
  fm_pivot_item_info_t item{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_pivot_field_item_at(handle_, args.sheet, args.pivot_idx,
                                                                              args.field_idx, item_idx, &item)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    item = fm_pivot_item_info_t{};
    item.cache_index = -1;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("name", JsString(env, item.name));
  out.Set("cacheIndex", Napi::Number::New(env, item.cache_index));
  out.Set("visible", Napi::Boolean::New(env, item.visible != 0));
  return out;
}

Napi::Value Workbook::PivotFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeIndexResult(env, MakeBindingArgumentError(env, "pivotFieldAdd: `spec` must be an object"), 0);
  }
  Napi::Object spec = info[2].As<Napi::Object>();

  CheckedSpecReader reader(env);
  std::string source_name;
  std::string custom_name;
  std::string number_format;
  const bool has_source = reader.String(spec, "sourceName", &source_name);
  const bool has_custom = reader.String(spec, "customName", &custom_name);
  const bool has_nfmt = reader.String(spec, "numberFormat", &number_format);

  fm_pivot_field_spec_t c_spec{};
  // `sourceName` is required; an omitted key passes NULL through rather
  // than the coerced literal string "undefined" `.ToString()` would
  // otherwise produce, so the C ABI's own kBindingNullPointer check is
  // what rejects the call.
  c_spec.source_name = has_source ? source_name.c_str() : nullptr;
  c_spec.custom_name = has_custom ? custom_name.c_str() : nullptr;
  c_spec.axis = static_cast<fm_pivot_axis_t>(reader.U32(spec, "axis", 0U));
  c_spec.subtotal_top = reader.Bool(spec, "subtotalTop", true) ? 1 : 0;
  c_spec.number_format = has_nfmt ? number_format.c_str() : nullptr;

  if (!reader.ok()) {
    // A malformed nested value left a pending JS exception: stop before the
    // C ABI call commits a default value for it.
    return env.Undefined();
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_field_add(handle_, sheet, pivot_idx, &c_spec, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotFieldClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_field_clear);
}

Napi::Value Workbook::PivotFieldSetAxis(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_field_set_axis);
}

Napi::Value Workbook::PivotFieldSetSort(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::int32_t order = ArgI32(info, 3);
    const std::string by_field = ArgString(info, 4);
    const char* by = by_field.empty() ? nullptr : by_field.c_str();
    return fm_workbook_pivot_field_set_sort(handle, field.sheet, field.pivot_idx, field.field_idx, order, by);
  });
}

Napi::Value Workbook::PivotFieldSetSubtotalTop(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    return fm_workbook_pivot_field_set_subtotal_top(handle, field.sheet, field.pivot_idx, field.field_idx,
                                                    ArgFlag(info, 3));
  });
}

Napi::Value Workbook::PivotFieldAddItem(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::string name = ArgString(info, 3);
    const int32_t visible = ArgFlag(info, 4);
    return fm_workbook_pivot_field_add_item(handle, field.sheet, field.pivot_idx, field.field_idx, name.c_str(),
                                            visible);
  });
}

Napi::Value Workbook::PivotFieldAddItemAt(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const uint32_t cache_index = ArgU32(info, 3);
    const int32_t visible = ArgFlag(info, 4);
    return fm_workbook_pivot_field_add_item_at(handle, field.sheet, field.pivot_idx, field.field_idx, cache_index,
                                               visible);
  });
}

Napi::Value Workbook::PivotFieldClearItems(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_field_clear_items);
}

Napi::Value Workbook::PivotFieldSetItemVisible(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
    const int32_t visible = ArgFlag(info, 4);
    return fm_workbook_pivot_field_set_item_visible(handle, field.sheet, field.pivot_idx, field.field_idx, item_idx,
                                                    visible);
  });
}

Napi::Value Workbook::PivotFieldAddSubtotalFn(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_field_add_subtotal_fn);
}

Napi::Value Workbook::PivotFieldClearSubtotalFns(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_field_clear_subtotal_fns);
}

Napi::Value Workbook::PivotFieldSetDateGroup(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::int32_t granularity = ArgI32(info, 3);
    const std::int32_t calendar = ArgI32(info, 4);
    const int32_t start_year = info.Length() > 5 ? info[5].As<Napi::Number>().Int32Value() : -1;
    const int32_t end_year = info.Length() > 6 ? info[6].As<Napi::Number>().Int32Value() : -1;
    const std::uint32_t interval_days = info.Length() > 7 ? info[7].As<Napi::Number>().Uint32Value() : 1;
    const double start_serial = info.Length() > 8 ? info[8].As<Napi::Number>().DoubleValue() : -1.0;
    const double end_serial = info.Length() > 9 ? info[9].As<Napi::Number>().DoubleValue() : -1.0;
    return fm_workbook_pivot_field_set_date_group(handle, field.sheet, field.pivot_idx, field.field_idx, granularity,
                                                  calendar, start_year, end_year, interval_days, start_serial,
                                                  end_serial);
  });
}

Napi::Value Workbook::PivotFieldClearDateGroup(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_field_clear_date_group);
}

Napi::Value Workbook::PivotFieldSetNumberFormat(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::string utf8 = ArgString(info, 3);
    return fm_workbook_pivot_field_set_number_format(handle, field.sheet, field.pivot_idx, field.field_idx,
                                                     utf8.c_str());
  });
}

Napi::Value Workbook::InvokePivotFieldOrder(const Napi::CallbackInfo& info, PivotFieldOrderFn fn) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  CheckedSpecReader reader(env);
  std::vector<uint32_t> indices;
  if (info.Length() > 2 && !info[2].IsUndefined() && !info[2].IsNull()) {
    Napi::Array array;
    if (!reader.Array(info[2], "indices", &array)) {
      return env.Undefined();
    }
    const uint32_t length = array.Length();
    indices.reserve(length);
    for (uint32_t i = 0; i < length; ++i) {
      Napi::Value value;
      (void)reader.ArrayElement(array, i, &value);
      indices.push_back(reader.U32(value, "indices element", 0U));
    }
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_status_t rc = fn(handle_, sheet, pivot_idx, indices.empty() ? nullptr : indices.data(), indices.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetRowFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_workbook_pivot_set_row_field_order);
}

Napi::Value Workbook::PivotSetColFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_workbook_pivot_set_col_field_order);
}

Napi::Value Workbook::PivotDataFieldCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_workbook_pivot_data_field_count);
}

Napi::Value Workbook::PivotDataFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeIndexResult(env, MakeBindingArgumentError(env, "pivotDataFieldAdd: `spec` must be an object"), 0);
  }
  Napi::Object spec = info[2].As<Napi::Object>();
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  CheckedSpecReader reader(env);
  BuildDataFieldSpec(reader, spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_data_field_add(handle_, sheet, pivot_idx, &c_spec, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotDataFieldAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const PivotFieldArgs args = ReadPivotField(info);
  fm_pivot_data_field_spec_t spec{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_pivot_data_field_at(handle_, args.sheet, args.pivot_idx, args.field_idx, &spec)
                         : kBindingInvalidHandle;
  if (rc != 0) {
    spec = fm_pivot_data_field_spec_t{};
    spec.show_as_base_field = -1;
    spec.show_as_base_item = -1;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("name", JsString(env, spec.name));
  out.Set("fieldIndex", Napi::Number::New(env, spec.field_index));
  out.Set("aggregation", Napi::Number::New(env, static_cast<int32_t>(spec.aggregation)));
  out.Set("numberFormat", JsString(env, spec.number_format));
  out.Set("showAs", Napi::Number::New(env, static_cast<int32_t>(spec.show_as)));
  out.Set("showAsBaseField", Napi::Number::New(env, spec.show_as_base_field));
  out.Set("showAsBaseItem", Napi::Number::New(env, spec.show_as_base_item));
  return out;
}

Napi::Value Workbook::PivotDataFieldClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_data_field_clear);
}

Napi::Value Workbook::PivotDataFieldSet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t data_field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  if (info.Length() < 4 || !info[3].IsObject()) {
    return MakeBindingArgumentError(env, "pivotDataFieldSet: `spec` must be an object");
  }
  Napi::Object spec = info[3].As<Napi::Object>();
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  CheckedSpecReader reader(env);
  BuildDataFieldSpec(reader, spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  if (!reader.ok()) {
    // See the matching guard in PivotDataFieldAdd.
    return env.Undefined();
  }
  fm_status_t rc = fm_workbook_pivot_data_field_set(handle_, sheet, pivot_idx, data_field_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_workbook_pivot_filter_count);
}

Napi::Value Workbook::PivotFilterAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  // The payload is emitted on every exit path, matching the WASM binding;
  // only `status.ok` and the values differ.
  Napi::Object out = Napi::Object::New(env);
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t filter_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_pivot_filter_spec_t spec{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_pivot_filter_at(handle_, sheet, pivot_idx, filter_idx, &spec)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    spec = fm_pivot_filter_spec_t{};
  }
  out.Set("status", MakeStatus(env, rc));
  out.Set("axis", Napi::Number::New(env, static_cast<int32_t>(spec.axis)));
  out.Set("fieldName", JsString(env, spec.field_name));
  out.Set("type", Napi::Number::New(env, static_cast<int32_t>(spec.type)));
  out.Set("dataFieldIndex", Napi::Number::New(env, spec.data_field_index));
  out.Set("valueKind", Napi::Number::New(env, static_cast<int32_t>(spec.value_kind)));
  out.Set("valueInt", Napi::Number::New(env, spec.value_int));
  out.Set("valueDouble", Napi::Number::New(env, spec.value_double));
  out.Set("valueText", JsString(env, spec.value_text));
  out.Set("valueHighKind", Napi::Number::New(env, static_cast<int32_t>(spec.value_high_kind)));
  out.Set("valueHighInt", Napi::Number::New(env, spec.value_high_int));
  out.Set("valueHighDouble", Napi::Number::New(env, spec.value_high_double));
  return out;
}

Napi::Value Workbook::PivotFilterAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeBindingArgumentError(env, "pivotFilterAdd: `spec` must be an object");
  }
  Napi::Object spec = info[2].As<Napi::Object>();

  CheckedSpecReader reader(env);
  std::string field_name;
  std::string value_text;
  const bool has_field = reader.String(spec, "fieldName", &field_name);
  const bool has_text = reader.String(spec, "valueText", &value_text);

  fm_pivot_filter_spec_t c_spec{};
  c_spec.axis = static_cast<fm_pivot_axis_t>(reader.U32(spec, "axis", 0U));
  // `fieldName` is required; see the `sourceName` comment in
  // `PivotFieldAdd` above -- an omitted key must reach the C ABI as NULL,
  // not the coerced literal string "undefined".
  c_spec.field_name = has_field ? field_name.c_str() : nullptr;
  c_spec.type = static_cast<fm_pivot_filter_type_t>(reader.U32(spec, "type", 0U));
  c_spec.value_kind = static_cast<fm_pivot_filter_value_kind_t>(reader.I32(spec, "valueKind", -1));
  c_spec.value_int = reader.I32(spec, "valueInt", 0);
  c_spec.value_double = reader.Double(spec, "valueDouble", 0.0);
  c_spec.value_text = has_text ? value_text.c_str() : nullptr;
  c_spec.value_high_kind = static_cast<fm_pivot_filter_value_kind_t>(reader.I32(spec, "valueHighKind", -1));
  c_spec.value_high_int = reader.I32(spec, "valueHighInt", 0);
  c_spec.value_high_double = reader.Double(spec, "valueHighDouble", 0.0);
  c_spec.data_field_index = reader.U32(spec, "dataFieldIndex", 0U);

  if (!reader.ok()) {
    // A malformed numeric field left a pending JS exception: stop before
    // the C ABI call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_workbook_pivot_filter_add(handle_, sheet, pivot_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_filter_clear);
}

Napi::Value Workbook::PivotFilterRemoveAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_pivot_filter_remove_at);
}

}  // namespace formulon_node
