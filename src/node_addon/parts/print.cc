// Print-pagination binding for the native Node addon.

#include <cstddef>
#include <cstdint>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

Napi::Value Workbook::Paginate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object result = Napi::Object::New(env);
  Napi::Array print_area = Napi::Array::New(env);
  Napi::Array horizontal_breaks = Napi::Array::New(env);
  Napi::Array vertical_breaks = Napi::Array::New(env);
  if (handle_ == nullptr) {
    result.Set("status", NullHandleError(env));
    result.Set("printArea", print_area);
    result.Set("horizontalBreaks", horizontal_breaks);
    result.Set("verticalBreaks", vertical_breaks);
    result.Set("pageCount", Napi::Number::New(env, 0));
    return result;
  }
  if (info.Length() < 1 || !info[0].IsNumber()) {
    Napi::TypeError::New(env, "paginate expects (sheet:number)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  fm_pagination_t* pagination = nullptr;
  const fm_status_t rc = fm_workbook_paginate(handle_, ArgU32(info, 0), &pagination);
  if (rc != 0) {
    result.Set("status", MakeErrorStatus(env, rc));
    result.Set("printArea", print_area);
    result.Set("horizontalBreaks", horizontal_breaks);
    result.Set("verticalBreaks", vertical_breaks);
    result.Set("pageCount", Napi::Number::New(env, 0));
    return result;
  }
  for (std::size_t i = 0; i < fm_pagination_print_area_count(pagination); ++i) {
    fm_print_range_t range{};
    if (fm_pagination_print_area_at(pagination, i, &range) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("firstRow", Napi::Number::New(env, range.first_row));
    item.Set("firstCol", Napi::Number::New(env, range.first_col));
    item.Set("lastRow", Napi::Number::New(env, range.last_row));
    item.Set("lastCol", Napi::Number::New(env, range.last_col));
    print_area.Set(i, item);
  }
  for (std::size_t i = 0; i < fm_pagination_horizontal_break_count(pagination); ++i) {
    uint32_t row = 0;
    if (fm_pagination_horizontal_break_at(pagination, i, &row) == 0) {
      horizontal_breaks.Set(i, Napi::Number::New(env, row));
    }
  }
  for (std::size_t i = 0; i < fm_pagination_vertical_break_count(pagination); ++i) {
    uint32_t col = 0;
    if (fm_pagination_vertical_break_at(pagination, i, &col) == 0) {
      vertical_breaks.Set(i, Napi::Number::New(env, col));
    }
  }
  result.Set("status", MakeOkStatus(env));
  result.Set("printArea", print_area);
  result.Set("horizontalBreaks", horizontal_breaks);
  result.Set("verticalBreaks", vertical_breaks);
  result.Set("pageCount", Napi::Number::New(env, fm_pagination_page_count(pagination)));
  fm_pagination_destroy(pagination);
  return result;
}

}  // namespace formulon_node
