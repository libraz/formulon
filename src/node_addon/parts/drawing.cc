// Drawing-image bindings: probe, enumerate, read, insert, remove,
// re-anchor, re-order, snapshot and restore embedded pictures.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Drawing images -------------------------------------------------

namespace {

// Returns the bytes of the Uint8Array at `info[idx]`; false when it is not one.
bool ReadImageBytes(const Napi::CallbackInfo& info, size_t idx, const uint8_t*& data, size_t& len) {
  if (info.Length() <= idx || !info[idx].IsTypedArray() ||
      info[idx].As<Napi::TypedArray>().TypedArrayType() != napi_uint8_array) {
    return false;
  }
  const Napi::Uint8Array u8 = info[idx].As<Napi::Uint8Array>();
  data = u8.Data();
  len = u8.ElementLength();
  return true;
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
    out.Set("status", MakeBindingArgumentError(env, "probeImage expects (bytes:Uint8Array)"));
    return out;
  }
  fm_image_info img{};
  const fm_status_t rc = fm_workbook_probe_image(handle_, data, len, &img);
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
  size_t count = 0;
  fm_status_t rc = fm_sheet_drawing_object_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_drawing_object o{};
    rc = fm_sheet_drawing_object_at(handle_, sheet, i, &o);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("objectId", JsNumber(env, o.object_id));
    item.Set("kind", JsNumber(env, o.kind));
    item.Set("anchorKind", JsNumber(env, o.anchor_kind));
    item.Set("editAs", JsNumber(env, o.edit_as));
    item.Set("fromRow", JsNumber(env, o.from_row));
    item.Set("fromCol", JsNumber(env, o.from_col));
    item.Set("fromRowOff", JsNumber(env, static_cast<double>(o.from_row_off)));
    item.Set("fromColOff", JsNumber(env, static_cast<double>(o.from_col_off)));
    item.Set("toRow", JsNumber(env, o.to_row));
    item.Set("toCol", JsNumber(env, o.to_col));
    item.Set("toRowOff", JsNumber(env, static_cast<double>(o.to_row_off)));
    item.Set("toColOff", JsNumber(env, static_cast<double>(o.to_col_off)));
    item.Set("cx", JsNumber(env, static_cast<double>(o.cx)));
    item.Set("cy", JsNumber(env, static_cast<double>(o.cy)));
    item.Set("imageFormat", JsNumber(env, o.image_format));
    item.Set("name", JsString(env, o.name));
    item.Set("descr", JsString(env, o.descr));
    item.Set("mediaPath", JsString(env, o.media_path));
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::GetImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", JsNumber(env, 0));
  out.Set("bytes", Napi::Uint8Array::New(env, 0));
  out.Set("pxWidth", JsNumber(env, 0));
  out.Set("pxHeight", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint8_t* bytes = nullptr;
  size_t len = 0;
  fm_image_info img{};
  const fm_status_t rc = fm_sheet_get_image(handle_, ArgU32(info, 0), ArgU32(info, 1), &bytes, &len, &img);
  if (rc == 0) {
    // The C pointer dies on the next mutation, so the bytes are copied out.
    Napi::Uint8Array copy = Napi::Uint8Array::New(env, len);
    if (len != 0) {
      std::memcpy(copy.Data(), bytes, len);
    }
    out.Set("format", JsNumber(env, img.format));
    out.Set("bytes", copy);
    out.Set("pxWidth", JsNumber(env, img.px_width));
    out.Set("pxHeight", JsNumber(env, img.px_height));
  }
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::InsertImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  const bool has_opts = info.Length() > 2 && info[2].IsObject();
  if (!ReadImageBytes(info, 1, data, len) || (info.Length() > 2 && !info[2].IsUndefined() && !has_opts)) {
    out.Set("status",
            MakeBindingArgumentError(env, "insertImage expects (sheet:number, bytes:Uint8Array, opts?:object)"));
    return out;
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = has_opts ? info[2].As<Napi::Object>() : Napi::Object::New(env);
  std::string name;
  std::string descr;
  reader.String(spec, "name", &name);
  reader.String(spec, "descr", &descr);
  fm_image_insert opts{};
  opts.name = name.c_str();
  opts.descr = descr.c_str();
  opts.anchor_kind = reader.I32(spec, "anchorKind", FM_ANCHOR_KIND_ONE_CELL);
  opts.edit_as = reader.I32(spec, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL);
  opts.row = reader.U32(spec, "row", 0U);
  opts.col = reader.U32(spec, "col", 0U);
  opts.row_off_emu = reader.I64(spec, "rowOffEmu", 0);
  opts.col_off_emu = reader.I64(spec, "colOffEmu", 0);
  opts.width_emu = reader.I64(spec, "widthEmu", 0);
  opts.height_emu = reader.I64(spec, "heightEmu", 0);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_insert_image(handle_, ArgU32(info, 0), data, len, &opts, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::RemoveImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_remove_image(handle_, ArgU32(info, 0), ArgU32(info, 1)));
}

Napi::Value Workbook::SetImageAnchor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeBindingArgumentError(env, "setImageAnchor expects (sheet:number, objectId:number, placement:object)");
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = info[2].As<Napi::Object>();
  fm_image_anchor anchor{};
  anchor.anchor_kind = reader.I32(spec, "anchorKind", FM_ANCHOR_KIND_ONE_CELL);
  anchor.edit_as = reader.I32(spec, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL);
  anchor.row = reader.U32(spec, "row", 0U);
  anchor.col = reader.U32(spec, "col", 0U);
  anchor.row_off_emu = reader.I64(spec, "rowOffEmu", 0);
  anchor.col_off_emu = reader.I64(spec, "colOffEmu", 0);
  anchor.width_emu = reader.I64(spec, "widthEmu", 0);
  anchor.height_emu = reader.I64(spec, "heightEmu", 0);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_image_anchor(handle_, ArgU32(info, 0), ArgU32(info, 1), &anchor));
}

Napi::Value Workbook::SetImageZOrder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_set_image_z_order(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2)));
}

Napi::Value Workbook::SnapshotImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("bytes", Napi::Uint8Array::New(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  uint8_t* bytes = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_image(handle_, ArgU32(info, 0), ArgU32(info, 1), &bytes, &len);
  if (rc == 0) {
    Napi::Uint8Array copy = Napi::Uint8Array::New(env, len);
    if (len != 0) {
      std::memcpy(copy.Data(), bytes, len);
    }
    out.Set("bytes", copy);
  }
  fm_buffer_free(bytes);
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::RestoreImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  const bool has_opts = info.Length() > 2 && info[2].IsObject();
  if (!ReadImageBytes(info, 1, data, len) || (info.Length() > 2 && !info[2].IsUndefined() && !has_opts)) {
    out.Set("status",
            MakeBindingArgumentError(env, "restoreImage expects (sheet:number, bytes:Uint8Array, opts?:object)"));
    return out;
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = has_opts ? info[2].As<Napi::Object>() : Napi::Object::New(env);
  const uint32_t flags = reader.Bool(spec, "newId", false) ? FM_IMAGE_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_restore_image(handle_, ArgU32(info, 0), data, len, flags, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

}  // namespace formulon_node
