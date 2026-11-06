// Drawing bindings: probe, enumerate, read and insert pictures, and the
// remove, re-anchor, re-order, snapshot and restore operations that apply to
// any picture or chart.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Drawing objects -------------------------------------------------

namespace {

// Returns the bytes of the Uint8Array at `info[idx]`; false when it is not one.
bool ReadImageBytes(const Napi::CallbackInfo& info, size_t idx, const uint8_t*& data, size_t& len) {
  Napi::Uint8Array u8;
  if (!BytesArg(info, idx, &u8)) {
    return false;
  }
  data = u8.Data();
  len = u8.ElementLength();
  return true;
}

struct BytesWithOpts {
  const uint8_t* data = nullptr;
  size_t len = 0;
  Napi::Object spec;
};

// Reads the `(sheet, bytes, opts?)` tail shared by `insertImage` and `restoreDrawing`;
// a rejected argument list throws a TypeError and returns false.
bool ReadBytesWithOpts(const Napi::CallbackInfo& info, const char* usage, BytesWithOpts& args) {
  Napi::Env env = info.Env();
  const bool has_opts = ObjectArg(info, 2, &args.spec);
  if (!ReadImageBytes(info, 1, args.data, args.len) || (!info[2].IsUndefined() && !has_opts)) {
    Napi::TypeError::New(env, usage).ThrowAsJavaScriptException();
    return false;
  }
  if (!has_opts) {
    args.spec = Napi::Object::New(env);
  }
  return true;
}

// The JS `DrawingObject` record of `o`.
Napi::Object DrawingObjectToJs(Napi::Env env, const fm_drawing_object_t& o) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("objectId", JsNumber(env, o.object_id));
  item.Set("kind", JsNumber(env, o.kind));
  item.Set("anchorKind", JsNumber(env, o.anchor_kind));
  item.Set("editAs", JsNumber(env, o.edit_as));
  item.Set("fromRow", JsNumber(env, o.from_row));
  item.Set("fromCol", JsNumber(env, o.from_col));
  item.Set("fromRowOffsetEmu", JsNumber(env, static_cast<double>(o.from_row_offset_emu)));
  item.Set("fromColOffsetEmu", JsNumber(env, static_cast<double>(o.from_col_offset_emu)));
  item.Set("toRow", JsNumber(env, o.to_row));
  item.Set("toCol", JsNumber(env, o.to_col));
  item.Set("toRowOffsetEmu", JsNumber(env, static_cast<double>(o.to_row_offset_emu)));
  item.Set("toColOffsetEmu", JsNumber(env, static_cast<double>(o.to_col_offset_emu)));
  item.Set("widthEmu", JsNumber(env, static_cast<double>(o.width_emu)));
  item.Set("heightEmu", JsNumber(env, static_cast<double>(o.height_emu)));
  item.Set("imageFormat", JsNumber(env, o.image_format));
  item.Set("name", JsString(env, o.name));
  item.Set("description", JsString(env, o.description));
  item.Set("mediaPath", JsStringOrNull(env, o.media_path));
  return item;
}

// The snapshot bytes, or the snapshot's failure; the buffer is released either way.
Napi::Value SnapshotBytesOrThrow(Napi::Env env, fm_status_t rc, uint8_t* bytes, std::size_t len) {
  if (rc != 0) {
    fm_buffer_destroy(bytes);
    return ThrowStatus(env, rc);
  }
  return TakeSavedBytes(env, bytes, len);
}

}  // namespace

Napi::Value Workbook::ProbeImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  if (!ReadImageBytes(info, 0, data, len)) {
    Napi::TypeError::New(env, "getImageInfo expects (bytes:Uint8Array)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  fm_image_info_t img{};
  const fm_status_t rc = fm_workbook_get_image_info(handle_, data, len, &img);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", JsNumber(env, img.format));
  out.Set("pxWidth", JsNumber(env, img.px_width));
  out.Set("pxHeight", JsNumber(env, img.px_height));
  return out;
}

Napi::Value Workbook::ListDrawingObjects(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr = Napi::Array::New(env);
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_drawing_object_t>(
      arr, [&](size_t* n) { return fm_sheet_get_drawing_object_count(handle_, sheet, n); },
      [&](size_t i, fm_drawing_object_t* o) { return fm_sheet_get_drawing_object_at(handle_, sheet, i, o); },
      [&](const fm_drawing_object_t& o) { return DrawingObjectToJs(env, o); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetDrawingObject(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_drawing_object_t o{};
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_drawing_object(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &o);
  return OptionalOrNull(env, rc, present,
                        rc == 0 && present != 0 ? Napi::Value(DrawingObjectToJs(env, o)) : env.Null());
}

Napi::Value Workbook::GetImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  const uint8_t* bytes = nullptr;
  size_t len = 0;
  fm_image_info_t img{};
  const fm_status_t rc = fm_sheet_get_image(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &bytes, &len, &img);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object image = Napi::Object::New(env);
  image.Set("format", JsNumber(env, img.format));
  image.Set("pxWidth", JsNumber(env, img.px_width));
  image.Set("pxHeight", JsNumber(env, img.px_height));
  // The C pointer dies on the next mutation, so the bytes are copied out.
  image.Set("bytes", CopyBytes(env, bytes, len));
  return OptionalOrNull(env, rc, present, image);
}

Napi::Value Workbook::InsertImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  BytesWithOpts args;
  if (!ReadBytesWithOpts(info, "addImage expects (sheet:number, bytes:Uint8Array, input?:object)", args)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  std::string name;
  std::string description;
  const fm_image_insert_t opts = PullImageInsert(reader, args.spec, FM_ANCHOR_KIND_ONE_CELL, name, description);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_add_image(handle_, ArgU32(info, 0), args.data, args.len, &opts, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

Napi::Value Workbook::RemoveDrawing(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_drawing);
}

Napi::Value Workbook::SetDrawingAnchor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 2);
  const fm_drawing_anchor_t anchor = PullDrawingAnchor(reader, spec, FM_ANCHOR_KIND_ONE_CELL);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_drawing_anchor(handle_, ArgU32(info, 0), ArgU32(info, 1), &anchor),
                      env.Undefined());
}

Napi::Value Workbook::SetDrawingName(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 2);
    return fm_sheet_set_drawing_name(handle, ArgU32(info, 0), ArgU32(info, 1), name.c_str());
  });
}

Napi::Value Workbook::SetDrawingZOrder(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_drawing_z_order);
}

Napi::Value Workbook::SnapshotDrawing(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  uint8_t* bytes = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_drawing(handle_, ArgU32(info, 0), ArgU32(info, 1), &bytes, &len);
  return SnapshotBytesOrThrow(env, rc, bytes, len);
}

Napi::Value Workbook::RestoreDrawing(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  BytesWithOpts args;
  if (!ReadBytesWithOpts(info, "restoreDrawing expects (sheet:number, bytes:Uint8Array, opts?:object)", args)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  const uint32_t flags = reader.Bool(args.spec, "newId", false) ? FM_DRAWING_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_restore_drawing(handle_, ArgU32(info, 0), args.data, args.len, flags, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

}  // namespace formulon_node
