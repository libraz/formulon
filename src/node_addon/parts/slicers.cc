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
fm_slicer_t PullSlicer(CheckedSpecReader& reader, const Napi::Object& spec, SlicerText& store) {
  fm_slicer_t s = FM_STRUCT_INIT(fm_slicer_t);
  bool present = false;
  if (reader.String(spec, "name", &store.name)) {
    s.fields |= FM_SLICER_FIELD_NAME;
    s.name = store.name.c_str();
  }
  if (reader.String(spec, "caption", &store.caption)) {
    s.fields |= FM_SLICER_FIELD_CAPTION;
    s.caption = store.caption.c_str();
  }
  // `null` removes the style, as the empty string does.
  const Napi::Value style = reader.Get(spec, "style");
  if (reader.ok() && !style.IsUndefined()) {
    s.fields |= FM_SLICER_FIELD_STYLE;
    s.style = style.IsNull() ? "" : (reader.String(style, "style", &store.style) ? store.style.c_str() : "");
  }
  s.column_count = reader.U32(spec, "columnCount", 1U, &present);
  s.fields |= present ? FM_SLICER_FIELD_COLUMN_COUNT : 0U;
  s.row_height_emu = reader.I64(spec, "rowHeightEmu", 0, &present);
  s.fields |= present ? FM_SLICER_FIELD_ROW_HEIGHT : 0U;
  s.from_row = reader.U32(spec, "fromRow", 0U, &present);
  if (present) {
    s.fields |= FM_SLICER_FIELD_ANCHOR;
    s.from_col = reader.U32(spec, "fromCol", 0U);
    s.from_row_offset_emu = reader.I64(spec, "fromRowOffsetEmu", 0);
    s.from_col_offset_emu = reader.I64(spec, "fromColOffsetEmu", 0);
    s.to_row = reader.U32(spec, "toRow", 0U);
    s.to_col = reader.U32(spec, "toCol", 0U);
    s.to_row_offset_emu = reader.I64(spec, "toRowOffsetEmu", 0);
    s.to_col_offset_emu = reader.I64(spec, "toColOffsetEmu", 0);
  }
  s.sort_order = reader.I32(spec, "sortOrder", 0, &present);
  s.fields |= present ? FM_SLICER_FIELD_SORT_ORDER : 0U;
  s.cross_filter = reader.I32(spec, "crossFilter", 0, &present);
  s.fields |= present ? FM_SLICER_FIELD_CROSS_FILTER : 0U;
  return s;
}

// The JS `Slicer` record of `s`.
Napi::Object SlicerToJs(Napi::Env env, const fm_slicer_t& s) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("objectId", JsNumber(env, s.object_id));
  item.Set("columnCount", JsNumber(env, s.column_count));
  item.Set("sortOrder", JsNumber(env, s.sort_order));
  item.Set("crossFilter", JsNumber(env, s.cross_filter));
  item.Set("anchorKind", JsNumber(env, s.anchor_kind));
  item.Set("fromRow", JsNumber(env, s.from_row));
  item.Set("fromCol", JsNumber(env, s.from_col));
  item.Set("fromRowOffsetEmu", JsNumber(env, static_cast<double>(s.from_row_offset_emu)));
  item.Set("fromColOffsetEmu", JsNumber(env, static_cast<double>(s.from_col_offset_emu)));
  item.Set("toRow", JsNumber(env, s.to_row));
  item.Set("toCol", JsNumber(env, s.to_col));
  item.Set("toRowOffsetEmu", JsNumber(env, static_cast<double>(s.to_row_offset_emu)));
  item.Set("toColOffsetEmu", JsNumber(env, static_cast<double>(s.to_col_offset_emu)));
  item.Set("rowHeightEmu", JsNumber(env, static_cast<double>(s.row_height_emu)));
  item.Set("name", JsString(env, s.name));
  item.Set("caption", JsString(env, s.caption));
  item.Set("cacheName", JsString(env, s.cache_name));
  item.Set("tableName", JsString(env, s.table_name));
  item.Set("columnName", JsString(env, s.column_name));
  item.Set("style", JsStringOrNull(env, s.style));
  return item;
}

}  // namespace

Napi::Value Workbook::ListSlicers(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr = Napi::Array::New(env);
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_slicer_t>(
      arr, [&](size_t* n) { return fm_sheet_get_slicer_count(handle_, sheet, n); },
      [&](size_t i, fm_slicer_t* s) { return fm_sheet_get_slicer_at(handle_, sheet, i, s); },
      [&](const fm_slicer_t& s) { return SlicerToJs(env, s); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetSlicer(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_slicer_t s = FM_STRUCT_INIT(fm_slicer_t);
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &s);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(SlicerToJs(env, s)) : env.Null());
}

Napi::Value Workbook::AddSlicer(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  SlicerText store;
  const fm_slicer_t slicer = PullSlicer(reader, ArgObject(info, 4), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_add_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                             ArgU32(info, 3), &slicer, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

Napi::Value Workbook::UpdateSlicer(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  SlicerText store;
  const fm_slicer_t slicer = PullSlicer(reader, ArgObject(info, 2), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), &slicer), env.Undefined());
}

Napi::Value Workbook::RemoveSlicer(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_slicer);
}

}  // namespace formulon_node
