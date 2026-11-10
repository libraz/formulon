// PivotTable mutation bindings: pivot create/remove, anchor / name /
// totals, per-field configuration (axis, sort, items, subtotals,
// date grouping, number format), data fields, filters,
// and the row/column field-order setters.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct PivotFieldArgs {
  std::size_t sheet;
  std::size_t pivot_index;
  std::size_t field_index;
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
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(Workbook::ArgU32(info, 1));
  std::size_t count = 0;
  const fm_status_t rc = fn(handle, sheet, pivot_index, &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

// A date-group bound: absent, `null` or NaN takes the field's own first / last value.
double ReadOptionalBound(CheckedSpecReader& reader, const Napi::Value& value, const char* key, bool* present) {
  const double bound = reader.Double(value, key, std::numeric_limits<double>::quiet_NaN());
  *present = !std::isnan(bound);
  return *present ? bound : 0.0;
}

Napi::Value OptionsFlag(Napi::Env env, int32_t v) {
  return Napi::Boolean::New(env, v != 0);
}

}  // namespace

Napi::Value Workbook::PivotCreate(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::string name = ArgString(info, 1);
    const uint32_t cache_id = ArgU32(info, 2);
    const uint32_t anchor_row = ArgU32(info, 3);
    const uint32_t anchor_col = ArgU32(info, 4);
    fm_pivot_spec_t spec = FM_STRUCT_INIT(fm_pivot_spec_t);
    spec.name = name.c_str();
    spec.cache_id = cache_id;
    spec.anchor_row = anchor_row;
    spec.anchor_col = anchor_col;
    return fm_sheet_add_pivot(handle, sheet, &spec, out);
  });
}

Napi::Value Workbook::PivotRemove(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_pivot);
}

Napi::Value Workbook::PivotRename(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
    const std::string name = ArgString(info, 2);
    return fm_sheet_rename_pivot(handle, sheet, pivot_index, name.c_str());
  });
}

Napi::Value Workbook::PivotSetAnchor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // The anchor is one cell, `{ row, col }`.
  CheckedSpecReader reader(env);
  Napi::Object anchor;
  if (!reader.Object(info[2], "cell", &anchor)) {
    return env.Undefined();
  }
  const uint32_t row = reader.RequiredU32(anchor, "row");
  const uint32_t col = reader.RequiredU32(anchor, "col");
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_pivot_anchor(handle_, ArgU32(info, 0), ArgU32(info, 1), row, col),
                      env.Undefined());
}

Napi::Value Workbook::PivotUpdateGrandTotals(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
    fm_pivot_grand_totals_patch_t patch = FM_STRUCT_INIT(fm_pivot_grand_totals_patch_t);
    patch.rows_op = FM_PATCH_OP_SET;
    patch.rows = ArgFlag(info, 2);
    patch.columns_op = FM_PATCH_OP_SET;
    patch.columns = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_update_pivot_grand_totals(handle, sheet, pivot_index, &patch);
  });
}

Napi::Value Workbook::PivotUpdateStyleOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object style = ArgObject(info, 2);
  // An omitted key keeps the stored value; only `styleName` accepts `null`, which removes the style name.
  CheckedSpecReader reader(env);
  fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  std::string style_name;
  patch.style_name_op = PatchString(reader, style, "styleName", true, &style_name, &patch.style_name);
  patch.show_row_headers_op = PatchBool(reader, style, "showRowHeaders", &patch.show_row_headers);
  patch.show_column_headers_op = PatchBool(reader, style, "showColumnHeaders", &patch.show_column_headers);
  patch.show_row_stripes_op = PatchBool(reader, style, "showRowStripes", &patch.show_row_stripes);
  patch.show_column_stripes_op = PatchBool(reader, style, "showColumnStripes", &patch.show_column_stripes);
  patch.show_last_column_op = PatchBool(reader, style, "showLastColumn", &patch.show_last_column);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_pivot_style_options(handle_, sheet, pivot_index, &patch), env.Undefined());
}

Napi::Value Workbook::PivotGetLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  std::int32_t layout = FM_PIVOT_REPORT_LAYOUT_COMPACT;
  const fm_status_t rc = fm_sheet_get_pivot_report_layout(handle_, sheet, pivot_index, &layout);
  return ValueOrThrow(env, rc, Napi::Number::New(env, layout));
}

Napi::Value Workbook::PivotSetLayout(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_pivot_report_layout);
}

Napi::Value Workbook::PivotCacheId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  uint32_t out = 0;
  const fm_status_t rc = fm_sheet_get_pivot_cache_id(handle_, sheet, pivot_index, &out);
  return ValueOrThrow(env, rc, JsNumber(env, out));
}

Napi::Value Workbook::PivotName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  const char* name = nullptr;
  const fm_status_t rc = fm_sheet_get_pivot_name(handle_, sheet, pivot_index, &name);
  return ValueOrThrow(env, rc, JsString(env, name));
}

Napi::Value Workbook::PivotRuntimeId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  uint32_t out = 0;
  const fm_status_t rc = fm_sheet_get_pivot_runtime_id(handle_, sheet, pivot_index, &out);
  return ValueOrThrow(env, rc, JsNumber(env, out));
}

Napi::Value Workbook::ResolvePivotRuntimeId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t out = 0;
  const fm_status_t rc = fm_sheet_resolve_pivot_runtime_id(handle_, sheet, ArgU32(info, 1), &out);
  return ValueOrThrow(env, rc, JsNumber(env, static_cast<double>(out)));
}

Napi::Value Workbook::PivotGetGrandTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_grand_totals_t totals = FM_STRUCT_INIT(fm_pivot_grand_totals_t);
  const fm_status_t rc = fm_sheet_get_pivot_grand_totals(handle_, sheet, pivot_index, &totals);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("rows", Napi::Boolean::New(env, totals.rows != 0));
  out.Set("columns", Napi::Boolean::New(env, totals.columns != 0));
  return out;
}

namespace {

using PivotFieldOrderGetter = fm_status_t (*)(const fm_workbook_t*, size_t, size_t, uint32_t*, size_t, size_t*);

Napi::Value GetPivotFieldOrder(const Napi::CallbackInfo& info, const fm_workbook_t* handle, PivotFieldOrderGetter fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(Workbook::ArgU32(info, 1));
  std::vector<uint32_t> data;
  const fm_status_t rc = FillCallerBuffer(data, [&](uint32_t* out, std::size_t capacity, std::size_t* count) {
    return fn(handle, sheet, pivot_index, out, capacity, count);
  });
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array indices = Napi::Array::New(env);
  for (std::size_t i = 0; i < data.size(); ++i) {
    const double index = data[i] == FM_PIVOT_VALUES_FIELD ? -2.0 : static_cast<double>(data[i]);
    indices.Set(static_cast<uint32_t>(i), Napi::Number::New(env, index));
  }
  return indices;
}

}  // namespace

Napi::Value Workbook::PivotRowFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_sheet_get_pivot_row_field_order);
}

Napi::Value Workbook::PivotGetStyleOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_style_options_t style = FM_STRUCT_INIT(fm_pivot_style_options_t);
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_pivot_style_options(handle_, sheet, pivot_index, &present, &style);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, 0, env.Null());
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("styleName", JsStringOrNull(env, style.style_name));
  out.Set("showRowHeaders", Napi::Boolean::New(env, style.show_row_headers != 0));
  out.Set("showColumnHeaders", Napi::Boolean::New(env, style.show_column_headers != 0));
  out.Set("showRowStripes", Napi::Boolean::New(env, style.show_row_stripes != 0));
  out.Set("showColumnStripes", Napi::Boolean::New(env, style.show_column_stripes != 0));
  out.Set("showLastColumn", Napi::Boolean::New(env, style.show_last_column != 0));
  return out;
}

Napi::Value Workbook::PivotColFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_sheet_get_pivot_column_field_order);
}

Napi::Value Workbook::PivotPageFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_sheet_get_pivot_page_field_order);
}

Napi::Value Workbook::PivotFieldCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_sheet_get_pivot_field_count);
}

Napi::Value Workbook::PivotFieldAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs args = ReadPivotField(info);
  fm_pivot_field_info_t field = FM_STRUCT_INIT(fm_pivot_field_info_t);
  const fm_status_t rc = fm_sheet_get_pivot_field_at(handle_, args.sheet, args.pivot_index, args.field_index, &field);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("sourceName", JsString(env, field.source_name));
  out.Set("customName", JsStringOrNull(env, field.custom_name));
  out.Set("axis", field.axis == FM_PIVOT_AXIS_ABSENT ? env.Null() : Napi::Value(Napi::Number::New(env, field.axis)));
  out.Set("subtotalTop", Napi::Boolean::New(env, field.subtotal_top == FM_TRISTATE_TRUE));
  out.Set("numberFormat", JsStringOrNull(env, field.number_format));
  out.Set("cacheFieldIndex", JsNumberOrNull(env, field.has_cache_field_index, field.cache_field_index));
  out.Set("sortOrder", Napi::Number::New(env, field.sort_order));
  out.Set("sortByField", JsStringOrNull(env, field.sort_by_field));
  out.Set("datePeriods", Napi::Number::New(env, field.date_periods));
  out.Set("dateCalendar", Napi::Number::New(env, field.date_calendar));
  out.Set("groupBaseIndex", JsNumberOrNull(env, field.has_group_base_index, field.group_base_index));
  out.Set("defaultSubtotal", Napi::Boolean::New(env, field.default_subtotal != 0));
  out.Set("insertBlankRow", Napi::Boolean::New(env, field.insert_blank_row != 0));
  out.Set("fillDownLabels", Napi::Boolean::New(env, field.fill_down_labels != 0));
  out.Set("showAll", Napi::Boolean::New(env, field.show_all != 0));
  out.Set("showDetail", Napi::Boolean::New(env, field.show_detail != 0));
  Napi::Array fns = Napi::Array::New(env);
  for (std::uint32_t agg = 0; agg < 32U; ++agg) {
    if ((field.subtotal_functions & (1U << agg)) != 0U) {
      fns.Set(fns.Length(), Napi::Number::New(env, agg));
    }
  }
  out.Set("subtotalFunctions", fns);
  return out;
}

Napi::Value Workbook::PivotFieldItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs args = ReadPivotField(info);
  std::size_t count = 0;
  const fm_status_t rc =
      fm_sheet_get_pivot_field_item_count(handle_, args.sheet, args.pivot_index, args.field_index, &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::PivotFieldItemAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs args = ReadPivotField(info);
  const std::size_t item_index = static_cast<std::size_t>(ArgU32(info, 3));
  fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
  const fm_status_t rc =
      fm_sheet_get_pivot_field_item_at(handle_, args.sheet, args.pivot_index, args.field_index, item_index, &item);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsStringOrNull(env, item.name));
  out.Set("cacheIndex", JsNumberOrNull(env, item.has_cache_index, item.cache_index));
  out.Set("visible", Napi::Boolean::New(env, item.visible != 0));
  out.Set("caption", item.caption != nullptr ? Napi::Value(Napi::String::New(env, item.caption)) : env.Null());
  out.Set("showDetail", Napi::Boolean::New(env, item.show_detail != 0));
  return out;
}

Napi::Value Workbook::PivotFieldClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_fields);
}

Napi::Value Workbook::PivotFieldSetAxis(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_pivot_field_axis);
}

// `item` is a new item's name, or `{ cacheIndex }` to bring in the cache's shared item at that position.
Napi::Value Workbook::PivotFieldAddItem(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs field = ReadPivotField(info);
  const int32_t visible = ArgFlag(info, 4);
  if (env.IsExceptionPending()) {
    return env.Undefined();
  }
  std::size_t item_index = 0;
  fm_status_t rc = 0;
  if (info[3].IsString()) {
    const std::string name = ArgString(info, 3);
    rc = fm_sheet_add_pivot_field_item(handle_, field.sheet, field.pivot_index, field.field_index, name.c_str(),
                                       visible, &item_index);
  } else {
    CheckedSpecReader reader(env);
    Napi::Object item;
    if (!reader.Object(info[3], "item", &item)) {
      if (reader.ok()) {
        ThrowArgumentTypeError(env, "item", "must be a string or an object");
      }
      return env.Undefined();
    }
    const uint32_t cache_index = reader.RequiredU32(item, "cacheIndex");
    if (!reader.ok()) {
      return env.Undefined();
    }
    rc = fm_sheet_add_pivot_field_item_by_cache_index(handle_, field.sheet, field.pivot_index, field.field_index,
                                                      cache_index, visible, &item_index);
  }
  return ValueOrThrow(env, rc, JsNumber(env, static_cast<double>(item_index)));
}

// Every member is optional: an omitted key keeps the stored value, and `null` clears `customName`,
// `numberFormat` and `sortByField`.
Napi::Value Workbook::PivotUpdateField(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs field = ReadPivotField(info);
  const Napi::Object spec = ArgObject(info, 3);
  CheckedSpecReader reader(env);
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  std::string custom_name;
  std::string number_format;
  std::string sort_by_field;
  patch.custom_name_op = PatchString(reader, spec, "customName", true, &custom_name, &patch.custom_name);
  patch.number_format_op = PatchString(reader, spec, "numberFormat", true, &number_format, &patch.number_format);
  patch.default_subtotal_op = PatchBool(reader, spec, "defaultSubtotal", &patch.default_subtotal);
  patch.insert_blank_row_op = PatchBool(reader, spec, "insertBlankRow", &patch.insert_blank_row);
  patch.fill_down_labels_op = PatchBool(reader, spec, "fillDownLabels", &patch.fill_down_labels);
  patch.show_all_op = PatchBool(reader, spec, "showAll", &patch.show_all);
  patch.subtotal_top_op = PatchBool(reader, spec, "subtotalTop", &patch.subtotal_top);
  patch.sort_order_op = PatchI32(reader, spec, "sortOrder", &patch.sort_order);
  patch.sort_by_field_op = PatchString(reader, spec, "sortByField", true, &sort_by_field, &patch.sort_by_field);
  patch.show_detail_op = PatchBool(reader, spec, "showDetail", &patch.show_detail);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env,
                      fm_sheet_update_pivot_field(handle_, field.sheet, field.pivot_index, field.field_index, &patch),
                      env.Undefined());
}

// Every member is optional: an omitted key keeps the stored value, and `null` clears `caption`.
Napi::Value Workbook::PivotUpdateFieldItem(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs field = ReadPivotField(info);
  const std::size_t item_index = static_cast<std::size_t>(ArgU32(info, 3));
  const Napi::Object spec = ArgObject(info, 4);
  CheckedSpecReader reader(env);
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  std::string caption;
  patch.visible_op = PatchBool(reader, spec, "visible", &patch.visible);
  patch.show_detail_op = PatchBool(reader, spec, "showDetail", &patch.show_detail);
  patch.caption_op = PatchString(reader, spec, "caption", true, &caption, &patch.caption);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(
      env,
      fm_sheet_update_pivot_field_item(handle_, field.sheet, field.pivot_index, field.field_index, item_index, &patch),
      env.Undefined());
}

Napi::Value Workbook::PivotFieldClearItems(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_field_items);
}

Napi::Value Workbook::PivotGetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_options_t o = FM_STRUCT_INIT(fm_pivot_options_t);
  const fm_status_t rc = fm_sheet_get_pivot_options(handle_, sheet, pivot_index, &o);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("showError", OptionsFlag(env, o.show_error));
  out.Set("errorCaption", JsStringOrNull(env, o.error_caption));
  out.Set("showMissing", OptionsFlag(env, o.show_missing));
  out.Set("missingCaption", JsStringOrNull(env, o.missing_caption));
  out.Set("showHeaders", OptionsFlag(env, o.show_headers));
  out.Set("showDrill", OptionsFlag(env, o.show_drill));
  out.Set("showDataTips", OptionsFlag(env, o.show_data_tips));
  out.Set("gridDropZones", OptionsFlag(env, o.grid_drop_zones));
  out.Set("showEmptyRow", OptionsFlag(env, o.show_empty_row));
  out.Set("showEmptyColumn", OptionsFlag(env, o.show_empty_column));
  out.Set("showDataDropDown", OptionsFlag(env, o.show_data_drop_down));
  out.Set("printDrill", OptionsFlag(env, o.print_drill));
  out.Set("fieldPrintTitles", OptionsFlag(env, o.field_print_titles));
  out.Set("itemPrintTitles", OptionsFlag(env, o.item_print_titles));
  out.Set("mergeItem", OptionsFlag(env, o.merge_item));
  out.Set("indent", Napi::Number::New(env, o.indent));
  out.Set("pageOverThenDown", OptionsFlag(env, o.page_over_then_down));
  out.Set("pageWrap", Napi::Number::New(env, o.page_wrap));
  out.Set("multipleFieldFilters", OptionsFlag(env, o.multiple_field_filters));
  out.Set("preserveFormatting", OptionsFlag(env, o.preserve_formatting));
  out.Set("customListSort", OptionsFlag(env, o.custom_list_sort));
  out.Set("altText", JsStringOrNull(env, o.alt_text));
  out.Set("altTextSummary", JsStringOrNull(env, o.alt_text_summary));
  out.Set("grandTotalCaption", JsStringOrNull(env, o.grand_total_caption));
  out.Set("autoFormatting", OptionsFlag(env, o.auto_formatting));
  out.Set("enableDrill", OptionsFlag(env, o.enable_drill));
  return out;
}

// Every member is optional: an omitted key keeps the pivot's current value, and `null` clears a caption.
Napi::Value Workbook::PivotUpdateOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object spec = ArgObject(info, 2);
  fm_pivot_options_patch_t o = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  CheckedSpecReader reader(env);
  std::deque<std::string> strings;
  const auto flag = [&](const char* key, fm_patch_op_t& op, int32_t& field) {
    op = PatchBool(reader, spec, key, &field);
  };
  const auto text = [&](const char* key, fm_patch_op_t& op, const char*& field) {
    strings.emplace_back();
    op = PatchString(reader, spec, key, true, &strings.back(), &field);
  };
  const auto count = [&](const char* key, fm_patch_op_t& op, uint32_t& field) {
    op = PatchU32(reader, spec, key, &field);
  };
  flag("showError", o.show_error_op, o.show_error);
  text("errorCaption", o.error_caption_op, o.error_caption);
  flag("showMissing", o.show_missing_op, o.show_missing);
  text("missingCaption", o.missing_caption_op, o.missing_caption);
  flag("showHeaders", o.show_headers_op, o.show_headers);
  flag("showDrill", o.show_drill_op, o.show_drill);
  flag("showDataTips", o.show_data_tips_op, o.show_data_tips);
  flag("gridDropZones", o.grid_drop_zones_op, o.grid_drop_zones);
  flag("showEmptyRow", o.show_empty_row_op, o.show_empty_row);
  flag("showEmptyColumn", o.show_empty_column_op, o.show_empty_column);
  flag("showDataDropDown", o.show_data_drop_down_op, o.show_data_drop_down);
  flag("printDrill", o.print_drill_op, o.print_drill);
  flag("fieldPrintTitles", o.field_print_titles_op, o.field_print_titles);
  flag("itemPrintTitles", o.item_print_titles_op, o.item_print_titles);
  flag("mergeItem", o.merge_item_op, o.merge_item);
  count("indent", o.indent_op, o.indent);
  flag("pageOverThenDown", o.page_over_then_down_op, o.page_over_then_down);
  count("pageWrap", o.page_wrap_op, o.page_wrap);
  flag("multipleFieldFilters", o.multiple_field_filters_op, o.multiple_field_filters);
  flag("preserveFormatting", o.preserve_formatting_op, o.preserve_formatting);
  flag("customListSort", o.custom_list_sort_op, o.custom_list_sort);
  text("altText", o.alt_text_op, o.alt_text);
  text("altTextSummary", o.alt_text_summary_op, o.alt_text_summary);
  text("grandTotalCaption", o.grand_total_caption_op, o.grand_total_caption);
  flag("autoFormatting", o.auto_formatting_op, o.auto_formatting);
  flag("enableDrill", o.enable_drill_op, o.enable_drill);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_pivot_options(handle_, sheet, pivot_index, &o), env.Undefined());
}

// `excludeDataField` is optional: omitted counts every data field's name as taken.
Napi::Value Workbook::PivotDataFieldDefaultName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 2));
  const int32_t aggregation = static_cast<int32_t>(ArgU32(info, 3));
  const std::size_t exclude = info.Length() > 4 && info[4].IsNumber() ? static_cast<std::size_t>(ArgU32(info, 4))
                                                                      : std::numeric_limits<std::size_t>::max();
  const char* name = nullptr;
  const fm_status_t rc =
      fm_sheet_get_pivot_data_field_default_name(handle_, sheet, pivot_index, field_index, aggregation, exclude, &name);
  return ValueOrThrow(env, rc, JsString(env, name));
}

Napi::Value Workbook::PivotProject(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_project_pivot);
}

Napi::Value Workbook::PivotProjectionPending(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t pending = 0;
  const fm_status_t rc = fm_sheet_is_pivot_projection_pending(handle_, static_cast<std::size_t>(ArgU32(info, 0)),
                                                              static_cast<std::size_t>(ArgU32(info, 1)), &pending);
  return ValueOrThrow(env, rc, Napi::Boolean::New(env, pending != 0));
}

Napi::Value Workbook::PivotShowDetails(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
    return fm_sheet_show_pivot_details(handle, sheet, pivot_index, ArgU32(info, 2), ArgU32(info, 3), out);
  });
}

Napi::Value Workbook::PivotFieldSetSubtotalFunctions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs field = ReadPivotField(info);
  CheckedSpecReader reader(env);
  Napi::Array list;
  std::vector<int32_t> aggregations;
  if (reader.Array(info[3], "aggregations", &list)) {
    aggregations.reserve(list.Length());
    for (uint32_t i = 0; i < list.Length() && reader.ok(); ++i) {
      Napi::Value value;
      if (reader.ArrayElement(list, i, &value)) {
        aggregations.push_back(reader.I32(value, "aggregations element", 0));
      }
    }
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc = fm_sheet_set_pivot_field_subtotal_functions(
      handle_, field.sheet, field.pivot_index, field.field_index, aggregations.empty() ? nullptr : aggregations.data(),
      aggregations.size());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::PivotFieldSetDateGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs field = ReadPivotField(info);
  // A null grouping removes the field's date grouping.
  if (info[3].IsNull() || info[3].IsUndefined()) {
    return ValueOrThrow(
        env, fm_sheet_set_pivot_field_date_group(handle_, field.sheet, field.pivot_index, field.field_index, nullptr),
        env.Undefined());
  }
  // The trailing three are optional on both surfaces: one-day intervals and the source data's own bounds.
  CheckedSpecReader reader(env);
  fm_pivot_date_group_spec_t group = FM_STRUCT_INIT(fm_pivot_date_group_spec_t);
  group.periods = ArgU32(info, 3);
  group.calendar = ArgI32(info, 4);
  group.interval_days = reader.U32(info[5], "intervalDays", 1U);
  bool has_start = false;
  bool has_end = false;
  group.start = ReadOptionalBound(reader, info[6], "start", &has_start);
  group.end = ReadOptionalBound(reader, info[7], "end", &has_end);
  group.has_start = has_start ? 1 : 0;
  group.has_end = has_end ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(
      env, fm_sheet_set_pivot_field_date_group(handle_, field.sheet, field.pivot_index, field.field_index, &group),
      env.Undefined());
}

Napi::Value Workbook::InvokePivotFieldOrder(const Napi::CallbackInfo& info, PivotFieldOrderFn fn) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  Napi::Array array;
  if (!ArrayArg(info, 2, &array)) {
    ThrowArgumentTypeError(env, "order", "must be an array");
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  std::vector<uint32_t> indices;
  const uint32_t length = array.Length();
  indices.reserve(length);
  for (uint32_t i = 0; i < length; ++i) {
    Napi::Value value;
    (void)reader.ArrayElement(array, i, &value);
    // -2 is the Values pseudo-field; any other negative reaches the C ABI as
    // an out-of-range index and is rejected there.
    indices.push_back(static_cast<uint32_t>(reader.I32(value, "indices element", 0)));
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc = fn(handle_, sheet, pivot_index, indices.empty() ? nullptr : indices.data(), indices.size());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::PivotSetRowFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_sheet_set_pivot_row_field_order);
}

Napi::Value Workbook::PivotSetColFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_sheet_set_pivot_column_field_order);
}

Napi::Value Workbook::PivotSetPageFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_sheet_set_pivot_page_field_order);
}

Napi::Value Workbook::PivotDataFieldCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_sheet_get_pivot_data_field_count);
}

Napi::Value Workbook::PivotDataFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  Napi::Object spec = ArgObject(info, 2);
  fm_pivot_data_field_t c_spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  std::string name_buf;
  std::string nfmt_buf;
  CheckedSpecReader reader(env);
  BuildDataFieldSpec(reader, spec, c_spec, name_buf, nfmt_buf);
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  std::size_t out = 0;
  const fm_status_t rc = fm_sheet_add_pivot_data_field(handle_, sheet, pivot_index, &c_spec, &out);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(out)));
}

Napi::Value Workbook::PivotDataFieldAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const PivotFieldArgs args = ReadPivotField(info);
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  const fm_status_t rc =
      fm_sheet_get_pivot_data_field_at(handle_, args.sheet, args.pivot_index, args.field_index, &spec);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsString(env, spec.name));
  out.Set("fieldIndex", Napi::Number::New(env, spec.field_index));
  out.Set("aggregation", Napi::Number::New(env, static_cast<int32_t>(spec.aggregation)));
  out.Set("numberFormat", JsStringOrNull(env, spec.number_format));
  out.Set("showAs", Napi::Number::New(env, static_cast<int32_t>(spec.show_as)));
  out.Set("showAsBaseField", JsNumberOrNull(env, spec.has_show_as_base_field, spec.show_as_base_field));
  out.Set("showAsBaseItem", JsNumberOrNull(env, spec.has_show_as_base_item, spec.show_as_base_item));
  return out;
}

Napi::Value Workbook::PivotDataFieldClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_data_fields);
}

Napi::Value Workbook::PivotDataFieldSet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t data_field_index = static_cast<std::size_t>(ArgU32(info, 2));
  Napi::Object spec = ArgObject(info, 3);
  fm_pivot_data_field_t c_spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  std::string name_buf;
  std::string nfmt_buf;
  CheckedSpecReader reader(env);
  BuildDataFieldSpec(reader, spec, c_spec, name_buf, nfmt_buf);
  if (!reader.ok()) {
    // See the matching guard in PivotDataFieldAdd.
    return env.Undefined();
  }
  const fm_status_t rc = fm_sheet_set_pivot_data_field(handle_, sheet, pivot_index, data_field_index, &c_spec);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::PivotFilterCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_sheet_get_pivot_filter_count);
}

namespace {

/// One pivot filter operand slot of a C filter spec as `number | string | null`.
Napi::Value FilterOperandToJs(Napi::Env env, int32_t kind, int32_t as_int, double as_double, const char* text) {
  switch (kind) {
    case FM_PIVOT_FILTER_OPERAND_KIND_INT:
      return Napi::Number::New(env, as_int);
    case FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE:
      return Napi::Number::New(env, as_double);
    case FM_PIVOT_FILTER_OPERAND_KIND_TEXT:
      return JsString(env, text != nullptr ? text : "");
    default:
      return env.Null();
  }
}

/// One pivot filter operand read from a JS `number | string | null` slot.
struct FilterOperand {
  int32_t kind = FM_PIVOT_FILTER_OPERAND_KIND_ABSENT;
  int32_t as_int = 0;
  double as_double = 0.0;
  std::string text;
};

/// Reads `owner[key]`: absent or null is no operand, a Number is an INT when `integral` and whole, else a
/// DOUBLE, a String is TEXT; anything else is a TypeError.
FilterOperand ReadFilterOperand(CheckedSpecReader& reader, const Napi::Object& owner, const char* key, bool integral) {
  FilterOperand out;
  const Napi::Value value = reader.Get(owner, key);
  if (!reader.ok() || value.IsUndefined() || value.IsNull()) {
    return out;
  }
  if (value.IsNumber()) {
    const double number = reader.Double(owner, key, 0.0);
    if (!reader.ok()) {
      return out;
    }
    if (integral && number >= -2147483648.0 && number <= 2147483647.0 && std::floor(number) == number) {
      out.kind = FM_PIVOT_FILTER_OPERAND_KIND_INT;
      out.as_int = static_cast<int32_t>(number);
    } else {
      out.kind = FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE;
      out.as_double = number;
    }
  } else if (value.IsString()) {
    reader.String(owner, key, &out.text);
    out.kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
  } else {
    reader.FailType(key, "a number or a string");
  }
  return out;
}

}  // namespace

Napi::Value Workbook::PivotFilterAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t filter_index = static_cast<std::size_t>(ArgU32(info, 2));
  fm_pivot_filter_t spec = FM_STRUCT_INIT(fm_pivot_filter_t);
  const fm_status_t rc = fm_sheet_get_pivot_filter_at(handle_, sheet, pivot_index, filter_index, &spec);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("fieldIndex", Napi::Number::New(env, spec.field_index));
  out.Set("kind", Napi::Number::New(env, spec.kind));
  out.Set("dataFieldIndex", Napi::Number::New(env, spec.data_field_index));
  out.Set("value", FilterOperandToJs(env, spec.value_kind, spec.value_int, spec.value_double, spec.value_text));
  out.Set("valueHigh", FilterOperandToJs(env, spec.value_high_kind, spec.value_high_int, spec.value_high_double,
                                         spec.value_high_text));
  return out;
}

Napi::Value Workbook::PivotFilterAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  Napi::Object spec = ArgObject(info, 2);

  CheckedSpecReader reader(env);
  fm_pivot_filter_t c_spec = FM_STRUCT_INIT(fm_pivot_filter_t);
  c_spec.field_index = reader.RequiredU32(spec, "fieldIndex");
  c_spec.kind = static_cast<int32_t>(reader.RequiredU32(spec, "kind"));
  // A count filter takes a whole number; every other numeric operand is a double.
  const bool count_kind =
      c_spec.kind == FM_PIVOT_FILTER_KIND_VALUE_TOP_COUNT || c_spec.kind == FM_PIVOT_FILTER_KIND_VALUE_BOTTOM_COUNT;
  const FilterOperand value = ReadFilterOperand(reader, spec, "value", count_kind);
  const FilterOperand value_high = ReadFilterOperand(reader, spec, "valueHigh", false);
  c_spec.value_kind = static_cast<fm_pivot_filter_operand_kind_t>(value.kind);
  c_spec.value_int = value.as_int;
  c_spec.value_double = value.as_double;
  c_spec.value_text = value.kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT ? value.text.c_str() : nullptr;
  c_spec.value_high_kind = static_cast<fm_pivot_filter_operand_kind_t>(value_high.kind);
  c_spec.value_high_int = value_high.as_int;
  c_spec.value_high_double = value_high.as_double;
  c_spec.value_high_text = value_high.kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT ? value_high.text.c_str() : nullptr;
  c_spec.data_field_index = reader.U32(spec, "dataFieldIndex", 0U);

  if (!reader.ok()) {
    // A malformed numeric field left a pending JS exception: stop before
    // the C ABI call commits a default value for it.
    return env.Undefined();
  }
  std::size_t out = 0;
  const fm_status_t rc = fm_sheet_add_pivot_filter(handle_, sheet, pivot_index, &c_spec, &out);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(out)));
}

Napi::Value Workbook::PivotFilterClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_filters);
}

Napi::Value Workbook::PivotFilterRemoveAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_pivot_filter);
}

}  // namespace formulon_node
