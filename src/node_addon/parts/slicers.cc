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

// Reads the fields `spec` carries into a patch: an omitted key keeps, and `null` clears `style`. The C ABI
// requires the four anchor cells together.
fm_slicer_patch_t PullSlicer(CheckedSpecReader& reader, const Napi::Object& spec, SlicerText& store) {
  fm_slicer_patch_t s = FM_STRUCT_INIT(fm_slicer_patch_t);
  s.name_op = PatchString(reader, spec, "name", false, &store.name, &s.name);
  s.caption_op = PatchString(reader, spec, "caption", false, &store.caption, &s.caption);
  s.style_op = PatchString(reader, spec, "style", true, &store.style, &s.style);
  s.column_count_op = PatchU32(reader, spec, "columnCount", &s.column_count);
  s.row_height_emu_op = PatchI64(reader, spec, "rowHeightEmu", &s.row_height_emu);
  s.from_row_op = PatchU32(reader, spec, "fromRow", &s.from_row);
  s.from_col_op = PatchU32(reader, spec, "fromCol", &s.from_col);
  s.from_row_offset_emu_op = PatchI64(reader, spec, "fromRowOffsetEmu", &s.from_row_offset_emu);
  s.from_col_offset_emu_op = PatchI64(reader, spec, "fromColOffsetEmu", &s.from_col_offset_emu);
  s.to_row_op = PatchU32(reader, spec, "toRow", &s.to_row);
  s.to_col_op = PatchU32(reader, spec, "toCol", &s.to_col);
  s.to_row_offset_emu_op = PatchI64(reader, spec, "toRowOffsetEmu", &s.to_row_offset_emu);
  s.to_col_offset_emu_op = PatchI64(reader, spec, "toColOffsetEmu", &s.to_col_offset_emu);
  s.sort_order_op = PatchI32(reader, spec, "sortOrder", &s.sort_order);
  s.cross_filter_op = PatchI32(reader, spec, "crossFilter", &s.cross_filter);
  return s;
}

// The JS `Slicer` record of `s`.
Napi::Object SlicerToJs(Napi::Env env, const fm_slicer_info_t& s) {
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
  const fm_status_t rc = CollectList<fm_slicer_info_t>(
      arr, [&](size_t* n) { return fm_sheet_get_slicer_count(handle_, sheet, n); },
      [&](size_t i, fm_slicer_info_t* s) { return fm_sheet_get_slicer_at(handle_, sheet, i, s); },
      [&](const fm_slicer_info_t& s) { return SlicerToJs(env, s); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetSlicer(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_slicer_info_t s = FM_STRUCT_INIT(fm_slicer_info_t);
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
  const fm_slicer_patch_t slicer = PullSlicer(reader, ArgObject(info, 4), store);
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
  const fm_slicer_patch_t slicer = PullSlicer(reader, ArgObject(info, 2), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), &slicer), env.Undefined());
}

Napi::Value Workbook::RemoveSlicer(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_slicer);
}

}  // namespace formulon_node
