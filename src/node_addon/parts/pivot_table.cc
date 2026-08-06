// PivotTable mutation bindings: pivot create/remove, anchor / name /
// totals, per-field configuration (axis, sort, items, subtotals,
// aggregations, date grouping, number format), data fields, filters,
// and the row/column field-order setters.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

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
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotRemove(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_remove(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::string name = ArgString(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_name(handle_, sheet, pivot_idx, name.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetAnchor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const uint32_t anchor_row = ArgU32(info, 2);
  const uint32_t anchor_col = ArgU32(info, 3);
  const uint32_t span_rows = ArgU32(info, 4);
  const uint32_t span_cols = ArgU32(info, 5);
  fm_status_t rc =
      fm_workbook_pivot_set_anchor(handle_, sheet, pivot_idx, anchor_row, anchor_col, span_rows, span_cols);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetGrandTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const bool rows_enabled = ArgBool(info, 2);
  const bool cols_enabled = ArgBool(info, 3);
  fm_status_t rc =
      fm_workbook_pivot_set_grand_totals(handle_, sheet, pivot_idx, rows_enabled ? 1 : 0, cols_enabled ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotGetLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("layout", Napi::Number::New(env, 0));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_layout_t layout = FM_PIVOT_LAYOUT_COMPACT;
  fm_status_t rc = fm_workbook_pivot_get_layout(handle_, sheet, pivot_idx, &layout);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("layout", Napi::Number::New(env, 0));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("layout", Napi::Number::New(env, static_cast<uint32_t>(layout)));
  return out;
}

Napi::Value Workbook::PivotSetLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const uint32_t layout = ArgU32(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_layout(handle_, sheet, pivot_idx, static_cast<fm_pivot_layout_t>(layout));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_field_count(handle_, sheet, pivot_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeIndexResult(env, MakeErrorStatus(env, kBindingInvalidHandle), 0);
  }
  Napi::Object spec = info[2].As<Napi::Object>();

  const bool has_custom = SpecHas(spec, "customName");
  const std::string source_name = spec.Get("sourceName").ToString().Utf8Value();
  const std::string custom_name = has_custom ? spec.Get("customName").ToString().Utf8Value() : std::string();
  const bool has_nfmt = SpecHas(spec, "numberFormat");
  const std::string number_format = has_nfmt ? spec.Get("numberFormat").ToString().Utf8Value() : std::string();

  fm_pivot_field_spec_t c_spec{};
  c_spec.source_name = source_name.c_str();
  c_spec.custom_name = has_custom ? custom_name.c_str() : nullptr;
  c_spec.axis = static_cast<fm_pivot_axis_t>(SpecPullU32(spec, "axis", 0U));
  c_spec.subtotal_top = SpecPullBool(spec, "subtotalTop", false) ? 1 : 0;
  c_spec.number_format = has_nfmt ? number_format.c_str() : nullptr;

  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_field_add(handle_, sheet, pivot_idx, &c_spec, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotFieldClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_field_clear(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetAxis(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t axis = ArgU32(info, 3);
  fm_status_t rc =
      fm_workbook_pivot_field_set_axis(handle_, sheet, pivot_idx, field_idx, static_cast<fm_pivot_axis_t>(axis));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetSort(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const bool ascending = ArgBool(info, 3);
  const std::string by_field = ArgString(info, 4);
  const char* by = by_field.empty() ? nullptr : by_field.c_str();
  fm_status_t rc = fm_workbook_pivot_field_set_sort(handle_, sheet, pivot_idx, field_idx, ascending ? 1 : 0, by);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetSubtotalTop(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const bool top = ArgBool(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_set_subtotal_top(handle_, sheet, pivot_idx, field_idx, top ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldAddAggregation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t agg = ArgU32(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_add_aggregation(handle_, sheet, pivot_idx, field_idx,
                                                           static_cast<fm_pivot_aggregation_t>(agg));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearAggregations(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_aggregations(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldAddItem(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string name = ArgString(info, 3);
  const bool visible = ArgBool(info, 4);
  fm_status_t rc =
      fm_workbook_pivot_field_add_item(handle_, sheet, pivot_idx, field_idx, name.c_str(), visible ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearItems(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_items(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetItemVisible(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::size_t item_idx = static_cast<std::size_t>(ArgU32(info, 3));
  const bool visible = ArgBool(info, 4);
  fm_status_t rc =
      fm_workbook_pivot_field_set_item_visible(handle_, sheet, pivot_idx, field_idx, item_idx, visible ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldAddSubtotalFn(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t agg = ArgU32(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_add_subtotal_fn(handle_, sheet, pivot_idx, field_idx,
                                                           static_cast<fm_pivot_aggregation_t>(agg));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearSubtotalFns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_subtotal_fns(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetDateGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const uint32_t granularity = ArgU32(info, 3);
  const uint32_t calendar = ArgU32(info, 4);
  const int32_t start_year = info.Length() > 5 ? info[5].As<Napi::Number>().Int32Value() : -1;
  const int32_t end_year = info.Length() > 6 ? info[6].As<Napi::Number>().Int32Value() : -1;
  fm_status_t rc = fm_workbook_pivot_field_set_date_group(
      handle_, sheet, pivot_idx, field_idx, static_cast<fm_pivot_date_grouping_t>(granularity),
      static_cast<fm_pivot_calendar_t>(calendar), start_year, end_year);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldClearDateGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_field_clear_date_group(handle_, sheet, pivot_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFieldSetNumberFormat(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string utf8 = ArgString(info, 3);
  fm_status_t rc = fm_workbook_pivot_field_set_number_format(handle_, sheet, pivot_idx, field_idx, utf8.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetRowFieldOrder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::vector<uint32_t> indices = ReadU32Array(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_row_field_order(handle_, sheet, pivot_idx,
                                                         indices.empty() ? nullptr : indices.data(), indices.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotSetColFieldOrder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::vector<uint32_t> indices = ReadU32Array(info, 2);
  fm_status_t rc = fm_workbook_pivot_set_col_field_order(handle_, sheet, pivot_idx,
                                                         indices.empty() ? nullptr : indices.data(), indices.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotDataFieldCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_data_field_count(handle_, sheet, pivot_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotDataFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeIndexResult(env, MakeErrorStatus(env, kBindingInvalidHandle), 0);
  }
  Napi::Object spec = info[2].As<Napi::Object>();
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  BuildDataFieldSpec(spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_data_field_add(handle_, sheet, pivot_idx, &c_spec, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotDataFieldClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_data_field_clear(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
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
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  Napi::Object spec = info[3].As<Napi::Object>();
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  BuildDataFieldSpec(spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  fm_status_t rc = fm_workbook_pivot_data_field_set(handle_, sheet, pivot_idx, data_field_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_filter_count(handle_, sheet, pivot_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotFilterAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  Napi::Object spec = info[2].As<Napi::Object>();

  const std::string field_name = spec.Get("fieldName").ToString().Utf8Value();
  const bool has_text = SpecHas(spec, "valueText");
  const std::string value_text = has_text ? spec.Get("valueText").ToString().Utf8Value() : std::string();

  fm_pivot_filter_spec_t c_spec{};
  c_spec.axis = static_cast<fm_pivot_axis_t>(SpecPullU32(spec, "axis", 0U));
  c_spec.field_name = field_name.c_str();
  c_spec.type = static_cast<fm_pivot_filter_type_t>(SpecPullU32(spec, "type", 0U));
  c_spec.value_kind = static_cast<fm_pivot_filter_value_kind_t>(SpecPullInt32(spec, "valueKind", -1));
  c_spec.value_int = SpecPullInt32(spec, "valueInt", 0);
  c_spec.value_double = SpecPullDouble(spec, "valueDouble", 0.0);
  c_spec.value_text = has_text ? value_text.c_str() : nullptr;
  c_spec.value_high_kind = static_cast<fm_pivot_filter_value_kind_t>(SpecPullInt32(spec, "valueHighKind", -1));
  c_spec.value_high_int = SpecPullInt32(spec, "valueHighInt", 0);
  c_spec.value_high_double = SpecPullDouble(spec, "valueHighDouble", 0.0);

  fm_status_t rc = fm_workbook_pivot_filter_add(handle_, sheet, pivot_idx, &c_spec);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_filter_clear(handle_, sheet, pivot_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotFilterRemoveAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t filter_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_filter_remove_at(handle_, sheet, pivot_idx, filter_idx);
  return MakeStatus(env, rc);
}

}  // namespace formulon_node
