// Table-slicer bindings: list, add, update and remove the slicers of a
// worksheet.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct SlicerText {
  std::string name;
  std::string caption;
  std::string style;
};

// Reads the fields `spec` carries; the `fields` mask names exactly those.
fm_slicer PullSlicer(CheckedSpecReader& reader, const Napi::Object& spec, SlicerText& store) {
  fm_slicer s{};
  bool present = false;
  if (reader.String(spec, "name", &store.name)) {
    s.fields |= FM_SLICER_FIELD_NAME;
    s.name = store.name.c_str();
  }
  if (reader.String(spec, "caption", &store.caption)) {
    s.fields |= FM_SLICER_FIELD_CAPTION;
    s.caption = store.caption.c_str();
  }
  if (reader.String(spec, "style", &store.style)) {
    s.fields |= FM_SLICER_FIELD_STYLE;
    s.style = store.style.c_str();
  }
  s.column_count = reader.U32(spec, "columnCount", 1U, &present);
  s.fields |= present ? FM_SLICER_FIELD_COLUMN_COUNT : 0U;
  s.row_height = reader.I64(spec, "rowHeight", 0, &present);
  s.fields |= present ? FM_SLICER_FIELD_ROW_HEIGHT : 0U;
  s.from_row = reader.U32(spec, "fromRow", 0U, &present);
  if (present) {
    s.fields |= FM_SLICER_FIELD_ANCHOR;
    s.from_col = reader.U32(spec, "fromCol", 0U);
    s.from_row_off = reader.I64(spec, "fromRowOff", 0);
    s.from_col_off = reader.I64(spec, "fromColOff", 0);
    s.to_row = reader.U32(spec, "toRow", 0U);
    s.to_col = reader.U32(spec, "toCol", 0U);
    s.to_row_off = reader.I64(spec, "toRowOff", 0);
    s.to_col_off = reader.I64(spec, "toColOff", 0);
  }
  s.sort_order = reader.I32(spec, "sortOrder", 0, &present);
  s.fields |= present ? FM_SLICER_FIELD_SORT_ORDER : 0U;
  s.cross_filter = reader.I32(spec, "crossFilter", 0, &present);
  s.fields |= present ? FM_SLICER_FIELD_CROSS_FILTER : 0U;
  return s;
}

}  // namespace

Napi::Value Workbook::ListSlicers(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  size_t count = 0;
  fm_status_t rc = fm_sheet_slicer_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_slicer s{};
    rc = fm_sheet_slicer_at(handle_, sheet, i, &s);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("objectId", JsNumber(env, s.object_id));
    item.Set("columnCount", JsNumber(env, s.column_count));
    item.Set("sortOrder", JsNumber(env, s.sort_order));
    item.Set("crossFilter", JsNumber(env, s.cross_filter));
    item.Set("anchorKind", JsNumber(env, s.anchor_kind));
    item.Set("fromRow", JsNumber(env, s.from_row));
    item.Set("fromCol", JsNumber(env, s.from_col));
    item.Set("fromRowOff", JsNumber(env, static_cast<double>(s.from_row_off)));
    item.Set("fromColOff", JsNumber(env, static_cast<double>(s.from_col_off)));
    item.Set("toRow", JsNumber(env, s.to_row));
    item.Set("toCol", JsNumber(env, s.to_col));
    item.Set("toRowOff", JsNumber(env, static_cast<double>(s.to_row_off)));
    item.Set("toColOff", JsNumber(env, static_cast<double>(s.to_col_off)));
    item.Set("rowHeight", JsNumber(env, static_cast<double>(s.row_height)));
    item.Set("name", JsString(env, s.name));
    item.Set("caption", JsString(env, s.caption));
    item.Set("cacheName", JsString(env, s.cache_name));
    item.Set("tableName", JsString(env, s.table_name));
    item.Set("columnName", JsString(env, s.column_name));
    item.Set("style", JsString(env, s.style));
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::AddSlicer(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  if (info.Length() < 4 || !info[3].IsObject()) {
    out.Set("status", MakeBindingArgumentError(
                          env, "addSlicer expects (sheet:number, table:number, column:number, spec:object)"));
    return out;
  }
  CheckedSpecReader reader(env);
  SlicerText store;
  const fm_slicer slicer = PullSlicer(reader, info[3].As<Napi::Object>(), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc =
      fm_sheet_add_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &slicer, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::UpdateSlicer(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeBindingArgumentError(env, "updateSlicer expects (sheet:number, index:number, spec:object)");
  }
  CheckedSpecReader reader(env);
  SlicerText store;
  const fm_slicer slicer = PullSlicer(reader, info[2].As<Napi::Object>(), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_update_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), &slicer));
}

Napi::Value Workbook::RemoveSlicer(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_slicer);
}

}  // namespace formulon_node
