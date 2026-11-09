// Table-slicer bindings: read, add and update the slicers of a worksheet.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct SlicerText {
  std::string name;
  std::string description;
  std::string caption;
  std::string style;
};

// Reads a slicer to add: the common placement keys plus the slicer keys; an omitted key takes the default.
fm_slicer_spec_t PullSlicerSpec(CheckedSpecReader& reader, const Napi::Object& spec, SlicerText& store) {
  fm_slicer_spec_t s = FM_STRUCT_INIT(fm_slicer_spec_t);
  s.placement = PullPlacement(reader, spec, store.name, store.description);
  s.column_count = reader.U32(spec, "columnCount", 0U);
  s.sort_order = reader.I32(spec, "sortOrder", FM_SLICER_SORT_ORDER_ASCENDING);
  s.cross_filter = reader.I32(spec, "crossFilter", FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_DATA_AT_TOP);
  s.row_height_emu = reader.I64(spec, "rowHeightEmu", 0);
  s.caption = reader.String(spec, "caption", &store.caption) ? store.caption.c_str() : nullptr;
  s.style = reader.String(spec, "style", &store.style) ? store.style.c_str() : nullptr;
  return s;
}

// Reads the slicer fields `spec` carries into a patch: an omitted key keeps, and `null` clears `style`.
fm_slicer_patch_t PullSlicerPatch(CheckedSpecReader& reader, const Napi::Object& spec, SlicerText& store) {
  fm_slicer_patch_t s = FM_STRUCT_INIT(fm_slicer_patch_t);
  s.column_count_op = PatchU32(reader, spec, "columnCount", &s.column_count);
  s.sort_order_op = PatchI32(reader, spec, "sortOrder", &s.sort_order);
  s.cross_filter_op = PatchI32(reader, spec, "crossFilter", &s.cross_filter);
  s.row_height_emu_op = PatchI64(reader, spec, "rowHeightEmu", &s.row_height_emu);
  s.caption_op = PatchString(reader, spec, "caption", false, &store.caption, &s.caption);
  s.style_op = PatchString(reader, spec, "style", true, &store.style, &s.style);
  return s;
}

// The JS `Slicer` record of `s`: the slicer-specific fields only.
Napi::Object SlicerToJs(Napi::Env env, const fm_slicer_info_t& s) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("objectId", JsNumber(env, s.object_id));
  item.Set("columnCount", JsNumber(env, s.column_count));
  item.Set("sortOrder", JsNumber(env, s.sort_order));
  item.Set("crossFilter", JsNumber(env, s.cross_filter));
  item.Set("rowHeightEmu", JsNumber(env, static_cast<double>(s.row_height_emu)));
  item.Set("caption", JsString(env, s.caption));
  item.Set("cacheName", JsString(env, s.cache_name));
  item.Set("tableName", JsString(env, s.table_name));
  item.Set("columnName", JsString(env, s.column_name));
  item.Set("style", JsStringOrNull(env, s.style));
  return item;
}

}  // namespace

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
  const fm_slicer_spec_t slicer = PullSlicerSpec(reader, ArgObject(info, 4), store);
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
  const fm_slicer_patch_t slicer = PullSlicerPatch(reader, ArgObject(info, 2), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_slicer(handle_, ArgU32(info, 0), ArgU32(info, 1), &slicer), env.Undefined());
}

}  // namespace formulon_node
