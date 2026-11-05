// PivotTable mutation bindings: pivot create/remove, anchor / name /
// totals, per-field configuration (axis, sort, items, subtotals,
// date grouping, number format), data fields, filters,
// and the row/column field-order setters.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <utility>
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

// The pivot options with their strings copied out of the model, so the
// struct stays valid across the `set_options` call that rewrites them.
struct OptionsSnapshot {
  fm_pivot_options_t c{};
  std::string error_caption;
  std::string missing_caption;
  std::string alt_text;
  std::string alt_text_summary;
  std::string grand_total_caption;

  void Pin() {
    const auto pin = [](const char*& field, std::string& store) {
      store = field != nullptr ? field : "";
      field = store.c_str();
    };
    pin(c.error_caption, error_caption);
    pin(c.missing_caption, missing_caption);
    pin(c.alt_text, alt_text);
    pin(c.alt_text_summary, alt_text_summary);
    pin(c.grand_total_caption, grand_total_caption);
  }
};

fm_status_t ReadOptions(const fm_workbook_t* handle, std::size_t sheet, std::size_t pivot_idx, OptionsSnapshot* out) {
  if (handle == nullptr) {
    return kBindingInvalidHandle;
  }
  const fm_status_t rc = fm_sheet_get_pivot_options(handle, sheet, pivot_idx, &out->c);
  if (rc == 0) {
    out->Pin();
  }
  return rc;
}

// A string member: an omitted key keeps the current value, '' clears it.
void MergeString(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, const char*& field,
                 std::string& store) {
  std::string value;
  if (reader.String(spec, key, &value)) {
    store = std::move(value);
    field = store.c_str();
  }
}

Napi::Value OptionsFlag(Napi::Env env, int32_t v) {
  return Napi::Boolean::New(env, v != 0);
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
  fm_status_t rc = fm_sheet_add_pivot(handle_, sheet, name.c_str(), cache_id, anchor_row, anchor_col, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotRemove(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_pivot);
}

Napi::Value Workbook::PivotSetName(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
    const std::string name = ArgString(info, 2);
    return fm_sheet_set_pivot_name(handle, sheet, pivot_idx, name.c_str());
  });
}

Napi::Value Workbook::PivotSetAnchor(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_pivot_anchor);
}

Napi::Value Workbook::PivotSetGrandTotals(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
    const int32_t rows_enabled = ArgFlag(info, 2);
    const int32_t cols_enabled = ArgFlag(info, 3);
    return fm_sheet_set_pivot_grand_totals(handle, sheet, pivot_idx, rows_enabled, cols_enabled);
  });
}

Napi::Value Workbook::PivotSetStyleInfo(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object style = ArgObject(info, 2);
  CheckedSpecReader reader(env);
  std::string name;
  const bool has_name = reader.String(style, "name", &name);
  const bool row_headers = reader.Bool(style, "showRowHeaders", false);
  const bool col_headers = reader.Bool(style, "showColHeaders", false);
  const bool row_stripes = reader.Bool(style, "showRowStripes", false);
  const bool col_stripes = reader.Bool(style, "showColStripes", false);
  const bool last_column = reader.Bool(style, "showLastColumn", false);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_pivot_style_info(handle_, sheet, pivot_idx, has_name ? name.c_str() : nullptr,
                                                          row_headers ? 1 : 0, col_headers ? 1 : 0, row_stripes ? 1 : 0,
                                                          col_stripes ? 1 : 0, last_column ? 1 : 0));
}

Napi::Value Workbook::PivotGetLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, kBindingInvalidHandle, "layout", 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::int32_t layout = FM_PIVOT_REPORT_LAYOUT_COMPACT;
  fm_status_t rc = fm_sheet_get_pivot_report_layout(handle_, sheet, pivot_idx, &layout);
  return MakeNumberFieldResult(env, rc, "layout", layout);
}

Napi::Value Workbook::PivotSetLayout(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_pivot_report_layout);
}

Napi::Value Workbook::PivotCacheId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  uint32_t out = 0;
  const fm_status_t rc = fm_sheet_get_pivot_cache_id(handle_, sheet, pivot_idx, &out);
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
  const fm_status_t rc = fm_sheet_get_pivot_name(handle_, sheet, pivot_idx, &name);
  return MakeStringResult(env, rc, rc == 0 ? name : nullptr);
}

Napi::Value Workbook::PivotGetGrandTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  int32_t rows = 0;
  int32_t cols = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_grand_totals(handle_, sheet, pivot_idx, &rows, &cols)
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
    const double index = data[i] == FM_PIVOT_VALUES_FIELD ? -2.0 : static_cast<double>(data[i]);
    indices.Set(static_cast<uint32_t>(i), Napi::Number::New(env, index));
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("indices", indices);
  return out;
}

}  // namespace

Napi::Value Workbook::PivotRowFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_sheet_get_pivot_row_field_order);
}

Napi::Value Workbook::PivotGetStyleInfo(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_style_info_t style{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_style_info(handle_, sheet, pivot_idx, &style) : kBindingInvalidHandle;
  if (rc != 0) {
    style = fm_pivot_style_info_t{};
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("present", Napi::Boolean::New(env, style.present != 0));
  out.Set("name", JsString(env, style.name));
  out.Set("showRowHeaders", Napi::Boolean::New(env, style.show_row_headers != 0));
  out.Set("showColHeaders", Napi::Boolean::New(env, style.show_col_headers != 0));
  out.Set("showRowStripes", Napi::Boolean::New(env, style.show_row_stripes != 0));
  out.Set("showColStripes", Napi::Boolean::New(env, style.show_col_stripes != 0));
  out.Set("showLastColumn", Napi::Boolean::New(env, style.show_last_column != 0));
  return out;
}

Napi::Value Workbook::PivotColFieldOrder(const Napi::CallbackInfo& info) {
  return GetPivotFieldOrder(info, handle_, &fm_sheet_get_pivot_col_field_order);
}

Napi::Value Workbook::PivotFieldCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_sheet_get_pivot_field_count);
}

Napi::Value Workbook::PivotFieldAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const PivotFieldArgs args = ReadPivotField(info);
  fm_pivot_field_info_t field{};
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_field(handle_, args.sheet, args.pivot_idx, args.field_idx, &field)
                             : kBindingInvalidHandle;
  if (rc != 0) {
    field = fm_pivot_field_info_t{};
    field.axis = -1;
    field.cache_field_index = -1;
    field.group_base_index = -1;
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
  out.Set("datePeriods", Napi::Number::New(env, field.date_periods));
  out.Set("dateCalendar", Napi::Number::New(env, field.date_calendar));
  out.Set("groupBaseIndex", Napi::Number::New(env, field.group_base_index));
  out.Set("defaultSubtotal", Napi::Boolean::New(env, field.default_subtotal != 0));
  out.Set("insertBlankRow", Napi::Boolean::New(env, field.insert_blank_row != 0));
  out.Set("fillDownLabels", Napi::Boolean::New(env, field.fill_down_labels != 0));
  out.Set("showAll", Napi::Boolean::New(env, field.show_all != 0));
  Napi::Array fns = Napi::Array::New(env);
  for (std::uint32_t agg = 0; agg < 32U; ++agg) {
    if ((field.subtotal_fns & (1U << agg)) != 0U) {
      fns.Set(fns.Length(), Napi::Number::New(env, agg));
    }
  }
  out.Set("subtotalFns", fns);
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
      fm_sheet_get_pivot_field_item_count(handle_, args.sheet, args.pivot_idx, args.field_idx, &count);
  return MakeNumberResult(env, rc, static_cast<double>(count));
}

Napi::Value Workbook::PivotFieldItemAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const PivotFieldArgs args = ReadPivotField(info);
  const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
  fm_pivot_item_info_t item{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_field_item(handle_, args.sheet, args.pivot_idx,
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
  out.Set("caption", item.caption != nullptr ? Napi::Value(Napi::String::New(env, item.caption)) : env.Null());
  out.Set("showDetail", Napi::Boolean::New(env, item.show_detail != 0));
  return out;
}

Napi::Value Workbook::PivotFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  Napi::Object spec = ArgObject(info, 2);

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
  fm_status_t rc = fm_sheet_add_pivot_field(handle_, sheet, pivot_idx, &c_spec, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotFieldClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_fields);
}

Napi::Value Workbook::PivotFieldSetAxis(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_pivot_field_axis);
}

Napi::Value Workbook::PivotFieldSetSort(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::int32_t order = ArgI32(info, 3);
    const std::string by_field = ArgString(info, 4);
    const char* by = by_field.empty() ? nullptr : by_field.c_str();
    return fm_sheet_set_pivot_field_sort(handle, field.sheet, field.pivot_idx, field.field_idx, order, by);
  });
}

Napi::Value Workbook::PivotFieldSetSubtotalTop(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    return fm_sheet_set_pivot_field_subtotal_top(handle, field.sheet, field.pivot_idx, field.field_idx,
                                                    ArgFlag(info, 3));
  });
}

Napi::Value Workbook::PivotFieldAddItem(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::string name = ArgString(info, 3);
    const int32_t visible = ArgFlag(info, 4);
    return fm_sheet_add_pivot_field_item(handle, field.sheet, field.pivot_idx, field.field_idx, name.c_str(),
                                            visible);
  });
}

Napi::Value Workbook::PivotFieldAddItemAt(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const uint32_t cache_index = ArgU32(info, 3);
    const int32_t visible = ArgFlag(info, 4);
    return fm_sheet_add_pivot_field_shared_item(handle, field.sheet, field.pivot_idx, field.field_idx, cache_index,
                                               visible);
  });
}

Napi::Value Workbook::PivotFieldClearItems(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_field_items);
}

Napi::Value Workbook::PivotFieldSetItemVisible(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
    const int32_t visible = ArgFlag(info, 4);
    return fm_sheet_set_pivot_field_item_visible(handle, field.sheet, field.pivot_idx, field.field_idx, item_idx,
                                                    visible);
  });
}

Napi::Value Workbook::PivotGetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_options_t o{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_options(handle_, sheet, pivot_idx, &o) : kBindingInvalidHandle;
  if (rc != 0) {
    o = fm_pivot_options_t{};
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("showError", OptionsFlag(env, o.show_error));
  out.Set("errorCaption", JsString(env, o.error_caption));
  out.Set("showMissing", OptionsFlag(env, o.show_missing));
  out.Set("missingCaption", JsString(env, o.missing_caption));
  out.Set("showHeaders", OptionsFlag(env, o.show_headers));
  out.Set("showDrill", OptionsFlag(env, o.show_drill));
  out.Set("showDataTips", OptionsFlag(env, o.show_data_tips));
  out.Set("gridDropZones", OptionsFlag(env, o.grid_drop_zones));
  out.Set("showEmptyRow", OptionsFlag(env, o.show_empty_row));
  out.Set("showEmptyCol", OptionsFlag(env, o.show_empty_col));
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
  out.Set("altText", JsString(env, o.alt_text));
  out.Set("altTextSummary", JsString(env, o.alt_text_summary));
  out.Set("grandTotalCaption", JsString(env, o.grand_total_caption));
  out.Set("autoFormatting", OptionsFlag(env, o.auto_formatting));
  out.Set("enableDrill", OptionsFlag(env, o.enable_drill));
  return out;
}

// Every member is optional: an omitted key keeps the pivot's current value,
// and a string member set to '' clears it.
Napi::Value Workbook::PivotSetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object spec = ArgObject(info, 2);
  OptionsSnapshot cur;
  const fm_status_t read_rc = ReadOptions(handle_, sheet, pivot_idx, &cur);
  if (read_rc != 0) {
    return MakeStatus(env, read_rc);
  }
  fm_pivot_options_t& o = cur.c;
  CheckedSpecReader reader(env);
  const auto flag = [&](const char* key, int32_t& field) { field = reader.Bool(spec, key, field != 0) ? 1 : 0; };
  flag("showError", o.show_error);
  MergeString(reader, spec, "errorCaption", o.error_caption, cur.error_caption);
  flag("showMissing", o.show_missing);
  MergeString(reader, spec, "missingCaption", o.missing_caption, cur.missing_caption);
  flag("showHeaders", o.show_headers);
  flag("showDrill", o.show_drill);
  flag("showDataTips", o.show_data_tips);
  flag("gridDropZones", o.grid_drop_zones);
  flag("showEmptyRow", o.show_empty_row);
  flag("showEmptyCol", o.show_empty_col);
  flag("showDataDropDown", o.show_data_drop_down);
  flag("printDrill", o.print_drill);
  flag("fieldPrintTitles", o.field_print_titles);
  flag("itemPrintTitles", o.item_print_titles);
  flag("mergeItem", o.merge_item);
  o.indent = reader.U32(spec, "indent", o.indent);
  flag("pageOverThenDown", o.page_over_then_down);
  o.page_wrap = reader.U32(spec, "pageWrap", o.page_wrap);
  flag("multipleFieldFilters", o.multiple_field_filters);
  flag("preserveFormatting", o.preserve_formatting);
  flag("customListSort", o.custom_list_sort);
  MergeString(reader, spec, "altText", o.alt_text, cur.alt_text);
  MergeString(reader, spec, "altTextSummary", o.alt_text_summary, cur.alt_text_summary);
  MergeString(reader, spec, "grandTotalCaption", o.grand_total_caption, cur.grand_total_caption);
  flag("autoFormatting", o.auto_formatting);
  flag("enableDrill", o.enable_drill);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_pivot_options(handle_, sheet, pivot_idx, &o));
}

Napi::Value Workbook::PivotCacheGetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_pivot_cache_options_t o{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_pivot_cache_options(handle_, ArgU32(info, 0), &o) : kBindingInvalidHandle;
  if (rc != 0) {
    o = fm_pivot_cache_options_t{};
    o.missing_items_limit = -1;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("saveData", OptionsFlag(env, o.save_data));
  out.Set("missingItemsLimit", Napi::Number::New(env, o.missing_items_limit));
  out.Set("refreshOnLoad", OptionsFlag(env, o.refresh_on_load));
  return out;
}

// Every member is optional; an omitted key keeps the cache's current value.
Napi::Value Workbook::PivotCacheSetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const Napi::Object spec = ArgObject(info, 1);
  fm_pivot_cache_options_t o{};
  const fm_status_t read_rc = fm_workbook_get_pivot_cache_options(handle_, cache_id, &o);
  if (read_rc != 0) {
    return MakeStatus(env, read_rc);
  }
  CheckedSpecReader reader(env);
  o.save_data = reader.Bool(spec, "saveData", o.save_data != 0) ? 1 : 0;
  o.missing_items_limit = reader.I32(spec, "missingItemsLimit", o.missing_items_limit);
  o.refresh_on_load = reader.Bool(spec, "refreshOnLoad", o.refresh_on_load != 0) ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_workbook_set_pivot_cache_options(handle_, cache_id, &o));
}

Napi::Value Workbook::PivotCacheCalculatedFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::string name = ArgString(info, 1);
  const std::string formula = ArgString(info, 2);
  std::size_t out = 0;
  const fm_status_t rc =
      fm_workbook_add_pivot_cache_calculated_field(handle_, ArgU32(info, 0), name.c_str(), formula.c_str(), &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotCacheCalculatedFieldSetFormula(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string formula = ArgString(info, 2);
    return fm_workbook_set_pivot_cache_calculated_field_formula(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), formula.c_str());
  });
}

Napi::Value Workbook::PivotCacheCalculatedFieldRemove(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_workbook_remove_pivot_cache_calculated_field(handle, ArgU32(info, 0),
                                                           static_cast<std::size_t>(ArgU32(info, 1)));
  });
}

Napi::Value Workbook::PivotCacheFieldFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* formula = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_field_formula(
                                   handle_, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), &formula)
                             : kBindingInvalidHandle;
  return MakeStringResult(env, rc, rc == 0 ? formula : nullptr);
}

Napi::Value Workbook::PivotCacheCalculatedItemAdd(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 2);
    const std::string formula = ArgString(info, 3);
    return fm_workbook_add_pivot_cache_calculated_item(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), name.c_str(), formula.c_str());
  });
}

Napi::Value Workbook::PivotCacheCalculatedItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_get_pivot_cache_calculated_item_count(handle_, ArgU32(info, 0), &count);
  return MakeNumberResult(env, rc, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheCalculatedItemAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_pivot_calculated_item_t item{};
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_calculated_item(
                                   handle_, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), &item)
                             : kBindingInvalidHandle;
  if (rc != 0) {
    item = fm_pivot_calculated_item_t{};
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("fieldIndex", Napi::Number::New(env, item.field_index));
  out.Set("name", JsString(env, item.name));
  out.Set("formula", JsString(env, item.formula));
  out.Set("scoped", Napi::Boolean::New(env, item.scoped != 0));
  return out;
}

Napi::Value Workbook::PivotCacheCalculatedItemSetFormula(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string formula = ArgString(info, 2);
    return fm_workbook_set_pivot_cache_calculated_item_formula(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), formula.c_str());
  });
}

Napi::Value Workbook::PivotCacheCalculatedItemMove(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_workbook_move_pivot_cache_calculated_item(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), static_cast<std::size_t>(ArgU32(info, 2)));
  });
}

Napi::Value Workbook::PivotCacheCalculatedItemRemove(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_workbook_remove_pivot_cache_calculated_item(handle, ArgU32(info, 0),
                                                          static_cast<std::size_t>(ArgU32(info, 1)));
  });
}

Napi::Value Workbook::PivotFieldSetCustomName(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::string name = ArgString(info, 3);
    return fm_sheet_set_pivot_field_custom_name(handle, field.sheet, field.pivot_idx, field.field_idx, name.c_str());
  });
}

// `excludeDataField` is optional: omitted counts every data field's name as taken.
Napi::Value Workbook::PivotDataFieldDefaultName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 2));
  const int32_t aggregation = static_cast<int32_t>(ArgU32(info, 3));
  const std::size_t exclude = info.Length() > 4 && info[4].IsNumber() ? static_cast<std::size_t>(ArgU32(info, 4))
                                                                      : std::numeric_limits<std::size_t>::max();
  const char* name = nullptr;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_data_field_default_name(
                                                  handle_, sheet, pivot_idx, field_index, aggregation, exclude, &name)
                                            : kBindingInvalidHandle;
  return MakeStringResult(env, rc, rc == 0 ? name : nullptr);
}

Napi::Value Workbook::PivotProject(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_project_pivot);
}

Napi::Value Workbook::PivotProjectionPending(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t pending = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_is_pivot_projection_pending(handle_, static_cast<std::size_t>(ArgU32(info, 0)),
                                                                    static_cast<std::size_t>(ArgU32(info, 1)), &pending)
                             : kBindingInvalidHandle;
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("value", Napi::Boolean::New(env, rc == 0 && pending != 0));
  return out;
}

Napi::Value Workbook::PivotShowDetails(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t out = 0;
  const fm_status_t rc =
      fm_sheet_show_pivot_details(handle_, sheet, pivot_idx, ArgU32(info, 2), ArgU32(info, 3), &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

namespace {

using PivotFieldFlagFn = fm_status_t (*)(fm_workbook_t*, size_t, size_t, size_t, int32_t);

Napi::Value InvokePivotFieldFlag(const Napi::CallbackInfo& info, fm_workbook_t* handle, PivotFieldFlagFn fn) {
  if (handle == nullptr) {
    return NullHandleError(info.Env());
  }
  const PivotFieldArgs field = ReadPivotField(info);
  return MakeStatus(info.Env(), fn(handle, field.sheet, field.pivot_idx, field.field_idx, Workbook::ArgFlag(info, 3)));
}

}  // namespace

Napi::Value Workbook::PivotFieldSetDefaultSubtotal(const Napi::CallbackInfo& info) {
  return InvokePivotFieldFlag(info, handle_, &fm_sheet_set_pivot_field_default_subtotal);
}

Napi::Value Workbook::PivotFieldSetInsertBlankRow(const Napi::CallbackInfo& info) {
  return InvokePivotFieldFlag(info, handle_, &fm_sheet_set_pivot_field_insert_blank_row);
}

Napi::Value Workbook::PivotFieldSetFillDownLabels(const Napi::CallbackInfo& info) {
  return InvokePivotFieldFlag(info, handle_, &fm_sheet_set_pivot_field_fill_down_labels);
}

Napi::Value Workbook::PivotFieldSetShowAll(const Napi::CallbackInfo& info) {
  return InvokePivotFieldFlag(info, handle_, &fm_sheet_set_pivot_field_show_all);
}

Napi::Value Workbook::PivotFieldSetShowDetail(const Napi::CallbackInfo& info) {
  return InvokePivotFieldFlag(info, handle_, &fm_sheet_set_pivot_field_show_detail);
}

Napi::Value Workbook::PivotFieldSetItemShowDetail(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
    return fm_sheet_set_pivot_field_item_show_detail(handle, field.sheet, field.pivot_idx, field.field_idx, item_idx,
                                                        ArgFlag(info, 4));
  });
}

Napi::Value Workbook::PivotFieldSetItemCaption(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
    const std::string caption = ArgString(info, 4);
    return fm_sheet_set_pivot_field_item_caption(handle, field.sheet, field.pivot_idx, field.field_idx, item_idx,
                                                    caption.c_str());
  });
}

Napi::Value Workbook::PivotFieldAddSubtotalFn(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_add_pivot_field_subtotal_fn);
}

Napi::Value Workbook::PivotFieldClearSubtotalFns(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_field_subtotal_fns);
}

Napi::Value Workbook::PivotFieldSetDateGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // The trailing three are optional on both surfaces: one-day intervals and
  // the source data's own bounds (NaN).
  constexpr double kAuto = std::numeric_limits<double>::quiet_NaN();
  CheckedSpecReader reader(env);
  const std::uint32_t interval_days = reader.U32(info[5], "intervalDays", 1U);
  const double start = reader.Double(info[6], "start", kAuto);
  const double end = reader.Double(info[7], "end", kAuto);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const PivotFieldArgs field = ReadPivotField(info);
  return MakeStatus(
      env, fm_sheet_set_pivot_field_date_group(handle_, field.sheet, field.pivot_idx, field.field_idx,
                                                  ArgU32(info, 3), ArgI32(info, 4), interval_days, start, end));
}

Napi::Value Workbook::PivotFieldClearDateGroup(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_field_date_group);
}

Napi::Value Workbook::PivotFieldSetNumberFormat(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const PivotFieldArgs field = ReadPivotField(info);
    const std::string utf8 = ArgString(info, 3);
    return fm_sheet_set_pivot_field_number_format(handle, field.sheet, field.pivot_idx, field.field_idx,
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
  Napi::Array array;
  if (!ArrayArg(info, 2, &array)) {
    return MakeBindingArgumentError(env, "pivot field order must be an array");
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
  fm_status_t rc = fn(handle_, sheet, pivot_idx, indices.empty() ? nullptr : indices.data(), indices.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetRowFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_sheet_set_pivot_row_field_order);
}

Napi::Value Workbook::PivotSetColFieldOrder(const Napi::CallbackInfo& info) {
  return InvokePivotFieldOrder(info, &fm_sheet_set_pivot_col_field_order);
}

Napi::Value Workbook::PivotDataFieldCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_sheet_get_pivot_data_field_count);
}

Napi::Value Workbook::PivotDataFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  Napi::Object spec = ArgObject(info, 2);
  fm_pivot_data_field_spec_t c_spec{};
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
  fm_status_t rc = fm_sheet_add_pivot_data_field(handle_, sheet, pivot_idx, &c_spec, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotDataFieldAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const PivotFieldArgs args = ReadPivotField(info);
  fm_pivot_data_field_spec_t spec{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_data_field(handle_, args.sheet, args.pivot_idx, args.field_idx, &spec)
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
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_data_fields);
}

Napi::Value Workbook::PivotDataFieldSet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t data_field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  Napi::Object spec = ArgObject(info, 3);
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  CheckedSpecReader reader(env);
  BuildDataFieldSpec(reader, spec, c_spec, name_buf, nfmt_buf);
  if (!reader.ok()) {
    // See the matching guard in PivotDataFieldAdd.
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_pivot_data_field(handle_, sheet, pivot_idx, data_field_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterCount(const Napi::CallbackInfo& info) {
  return PivotCountResult(info, handle_, &fm_sheet_get_pivot_filter_count);
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
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_filter(handle_, sheet, pivot_idx, filter_idx, &spec)
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
  out.Set("valueHighText", JsString(env, spec.value_high_text));
  return out;
}

Napi::Value Workbook::PivotFilterAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  Napi::Object spec = ArgObject(info, 2);

  CheckedSpecReader reader(env);
  std::string field_name;
  std::string value_text;
  const bool has_field = reader.String(spec, "fieldName", &field_name);
  const bool has_text = reader.String(spec, "valueText", &value_text);
  std::string value_high_text;
  const bool has_high_text = reader.String(spec, "valueHighText", &value_high_text);

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
  c_spec.value_high_text = has_high_text ? value_high_text.c_str() : nullptr;
  c_spec.data_field_index = reader.U32(spec, "dataFieldIndex", 0U);

  if (!reader.ok()) {
    // A malformed numeric field left a pending JS exception: stop before
    // the C ABI call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_add_pivot_filter(handle_, sheet, pivot_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_pivot_filters);
}

Napi::Value Workbook::PivotFilterRemoveAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_pivot_filter);
}

}  // namespace formulon_node
