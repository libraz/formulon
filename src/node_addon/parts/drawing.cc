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
// a rejected argument list sets `result.status` and returns false.
bool ReadBytesWithOpts(const Napi::CallbackInfo& info, const char* usage, Napi::Object& result, BytesWithOpts& args) {
  Napi::Env env = info.Env();
  const bool has_opts = ObjectArg(info, 2, &args.spec);
  if (!ReadImageBytes(info, 1, args.data, args.len) || (!info[2].IsUndefined() && !has_opts)) {
    result.Set("status", MakeBindingArgumentError(env, usage));
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
  item.Set("fromRowOffset", JsNumber(env, static_cast<double>(o.from_row_offset)));
  item.Set("fromColOffset", JsNumber(env, static_cast<double>(o.from_col_offset)));
  item.Set("toRow", JsNumber(env, o.to_row));
  item.Set("toCol", JsNumber(env, o.to_col));
  item.Set("toRowOffset", JsNumber(env, static_cast<double>(o.to_row_offset)));
  item.Set("toColOffset", JsNumber(env, static_cast<double>(o.to_col_offset)));
  item.Set("cx", JsNumber(env, static_cast<double>(o.cx)));
  item.Set("cy", JsNumber(env, static_cast<double>(o.cy)));
  item.Set("imageFormat", JsNumber(env, o.image_format));
  item.Set("name", JsString(env, o.name));
  item.Set("descr", JsString(env, o.descr));
  item.Set("mediaPath", JsString(env, o.media_path));
  return item;
}

}  // namespace

Napi::Value Workbook::ProbeImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", JsNumber(env, 0));
  out.Set("pxWidth", JsNumber(env, 0));
  out.Set("pxHeight", JsNumber(env, 0));
  const uint8_t* data = nullptr;
  size_t len = 0;
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  if (!ReadImageBytes(info, 0, data, len)) {
    out.Set("status", MakeBindingArgumentError(env, "getImageInfo expects (bytes:Uint8Array)"));
    return out;
  }
  fm_image_info_t img{};
  const fm_status_t rc = fm_workbook_get_image_info(handle_, data, len, &img);
  if (rc == 0) {
    out.Set("format", JsNumber(env, img.format));
    out.Set("pxWidth", JsNumber(env, img.px_width));
    out.Set("pxHeight", JsNumber(env, img.px_height));
  }
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::ListDrawingObjects(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_drawing_object_t>(
      arr, [&](size_t* n) { return fm_sheet_get_drawing_object_count(handle_, sheet, n); },
      [&](size_t i, fm_drawing_object_t* o) { return fm_sheet_get_drawing_object_at(handle_, sheet, i, o); },
      [&](const fm_drawing_object_t& o) { return DrawingObjectToJs(env, o); });
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::GetDrawingObject(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  fm_drawing_object_t o{};
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_drawing_object(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &o);
  return MakeOptionalResult(env, rc, present,
                            rc == 0 && present != 0 ? Napi::Value(DrawingObjectToJs(env, o)) : env.Null());
}

Napi::Value Workbook::GetImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  int32_t present = 0;
  const uint8_t* bytes = nullptr;
  size_t len = 0;
  fm_image_info_t img{};
  const fm_status_t rc = fm_sheet_get_image(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &bytes, &len, &img);
  if (rc != 0 || present == 0) {
    return MakeOptionalResult(env, rc, present, env.Null());
  }
  Napi::Object image = Napi::Object::New(env);
  image.Set("format", JsNumber(env, img.format));
  image.Set("pxWidth", JsNumber(env, img.px_width));
  image.Set("pxHeight", JsNumber(env, img.px_height));
  // The C pointer dies on the next mutation, so the bytes are copied out.
  image.Set("bytes", CopyBytes(env, bytes, len));
  return MakeOptionalResult(env, rc, present, image);
}

Napi::Value Workbook::InsertImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  BytesWithOpts args;
  if (!ReadBytesWithOpts(info, "addImage expects (sheet:number, bytes:Uint8Array, opts?:object)", out, args)) {
    return out;
  }
  CheckedSpecReader reader(env);
  std::string name;
  std::string descr;
  reader.String(args.spec, "name", &name);
  reader.String(args.spec, "descr", &descr);
  fm_image_insert_t opts{};
  opts.name = name.c_str();
  opts.descr = descr.c_str();
  PullPlacement(reader, args.spec, opts, FM_ANCHOR_KIND_ONE_CELL);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_add_image(handle_, ArgU32(info, 0), args.data, args.len, &opts, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::RemoveDrawing(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_drawing);
}

Napi::Value Workbook::SetDrawingAnchor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 2);
  fm_drawing_anchor_t anchor{};
  PullPlacement(reader, spec, anchor, FM_ANCHOR_KIND_ONE_CELL);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_drawing_anchor(handle_, ArgU32(info, 0), ArgU32(info, 1), &anchor));
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
    return SnapshotBytesResult(env, kBindingInvalidHandle, nullptr, 0);
  }
  uint8_t* bytes = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_drawing(handle_, ArgU32(info, 0), ArgU32(info, 1), &bytes, &len);
  return SnapshotBytesResult(env, rc, bytes, len);
}

Napi::Value Workbook::RestoreDrawing(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  BytesWithOpts args;
  if (!ReadBytesWithOpts(info, "restoreDrawing expects (sheet:number, bytes:Uint8Array, opts?:object)", out, args)) {
    return out;
  }
  CheckedSpecReader reader(env);
  const uint32_t flags = reader.Bool(args.spec, "newId", false) ? FM_DRAWING_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_restore_drawing(handle_, ArgU32(info, 0), args.data, args.len, flags, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

}  // namespace formulon_node
