// Drawing bindings: probe, enumerate, read and insert pictures, and the
// remove, update, re-order, snapshot and restore operations that apply to
// any drawing object.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/chart_internal.h"
#include "node_addon/parts/free_funcs.h"
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
bool ReadBytesWithOpts(const Napi::CallbackInfo& info, BytesWithOpts& args) {
  Napi::Env env = info.Env();
  const bool has_opts = ObjectArg(info, 2, &args.spec);
  if (!ReadImageBytes(info, 1, args.data, args.len)) {
    ThrowArgumentTypeError(env, "bytes", "must be a Uint8Array");
    return false;
  }
  if (!info[2].IsUndefined() && !has_opts) {
    ThrowArgumentTypeError(env, "", "must be an object");
    return false;
  }
  if (!has_opts) {
    args.spec = Napi::Object::New(env);
  }
  return true;
}

// The JS `DrawingAnchor` record of `a`.
Napi::Object DrawingAnchorToJs(Napi::Env env, const fm_drawing_anchor_t& a) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("kind", JsNumber(env, a.kind));
  item.Set("editAs", JsNumber(env, a.edit_as));
  item.Set("fromRow", JsNumber(env, a.from_row));
  item.Set("fromCol", JsNumber(env, a.from_col));
  item.Set("fromRowOffsetEmu", JsNumber(env, static_cast<double>(a.from_row_offset_emu)));
  item.Set("fromColOffsetEmu", JsNumber(env, static_cast<double>(a.from_col_offset_emu)));
  item.Set("toRow", JsNumber(env, a.to_row));
  item.Set("toCol", JsNumber(env, a.to_col));
  item.Set("toRowOffsetEmu", JsNumber(env, static_cast<double>(a.to_row_offset_emu)));
  item.Set("toColOffsetEmu", JsNumber(env, static_cast<double>(a.to_col_offset_emu)));
  item.Set("widthEmu", JsNumber(env, static_cast<double>(a.width_emu)));
  item.Set("heightEmu", JsNumber(env, static_cast<double>(a.height_emu)));
  return item;
}

// The JS `DrawingObject` record of `o`.
Napi::Object DrawingObjectToJs(Napi::Env env, const fm_drawing_object_info_t& o) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("objectId", JsNumber(env, o.object_id));
  item.Set("kind", JsNumber(env, o.kind));
  item.Set("hidden", JsBool(env, o.hidden));
  item.Set("anchor", DrawingAnchorToJs(env, o.anchor));
  item.Set("name", JsString(env, o.name));
  item.Set("description", JsStringOrNull(env, o.description));
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

namespace detail {

fm_drawing_patch_t PullDrawingPatch(CheckedSpecReader& reader, const Napi::Object& spec, std::string& name_store,
                                    std::string& description_store) {
  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.name_op = PatchString(reader, spec, "name", false, &name_store, &patch.name);
  patch.description_op = PatchString(reader, spec, "description", true, &description_store, &patch.description);
  patch.hidden_op = PatchBool(reader, spec, "hidden", &patch.hidden);
  Napi::Object anchor;
  if (ClassifyPatchKey(reader, reader.Get(spec, "anchor"), "anchor", false) == PatchKey::kSet &&
      reader.Object(spec, "anchor", &anchor)) {
    patch.anchor = PullDrawingAnchorPatch(reader, anchor);
  }
  return patch;
}

}  // namespace detail

Napi::Value ProbeImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint8_t* data = nullptr;
  size_t len = 0;
  if (!ReadImageBytes(info, 0, data, len)) {
    ThrowArgumentTypeError(env, "bytes", "must be a Uint8Array");
    return env.Undefined();
  }
  fm_image_info_t img = FM_STRUCT_INIT(fm_image_info_t);
  const fm_status_t rc = fm_get_image_info(data, len, &img);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", JsNumber(env, img.format));
  out.Set("widthPx", JsNumber(env, img.width_px));
  out.Set("heightPx", JsNumber(env, img.height_px));
  return out;
}

Napi::Value Workbook::ListDrawingObjects(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr = Napi::Array::New(env);
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_drawing_object_info_t>(
      arr, [&](size_t* n) { return fm_sheet_get_drawing_object_count(handle_, sheet, n); },
      [&](size_t i, fm_drawing_object_info_t* o) { return fm_sheet_get_drawing_object_at(handle_, sheet, i, o); },
      [&](const fm_drawing_object_info_t& o) { return DrawingObjectToJs(env, o); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetDrawingObject(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_drawing_object_info_t o = FM_STRUCT_INIT(fm_drawing_object_info_t);
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
  fm_image_info_t img = FM_STRUCT_INIT(fm_image_info_t);
  const fm_status_t rc = fm_sheet_get_picture(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &bytes, &len, &img);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object image = Napi::Object::New(env);
  image.Set("format", JsNumber(env, img.format));
  image.Set("widthPx", JsNumber(env, img.width_px));
  image.Set("heightPx", JsNumber(env, img.height_px));
  // The C pointer dies on the next mutation, so the bytes are copied out.
  image.Set("bytes", CopyBytes(env, bytes, len));
  return OptionalOrNull(env, rc, present, image);
}

Napi::Value Workbook::GetCellPicture(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  fm_cell_picture_t picture = FM_STRUCT_INIT(fm_cell_picture_t);
  const fm_status_t rc =
      fm_sheet_get_cell_picture(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &picture);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", JsNumber(env, picture.format));
  out.Set("widthPx", JsNumber(env, picture.width_px));
  out.Set("heightPx", JsNumber(env, picture.height_px));
  // The C pointers die on the next mutation, so the bytes are copied out.
  out.Set("bytes", CopyBytes(env, picture.bytes, static_cast<std::size_t>(picture.len)));
  out.Set("altText", JsStringOrNull(env, picture.alt_text));
  return OptionalOrNull(env, rc, present, out);
}

Napi::Value Workbook::SetCellPicture(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // `null` removes the picture and does nothing when there is none; an omitted `altText` is none.
  CheckedSpecReader reader(env);
  Napi::Object in;
  const bool present = reader.Object(info[3], "", &in);
  fm_cell_picture_t picture = FM_STRUCT_INIT(fm_cell_picture_t);
  std::string alt_text;
  if (present) {
    const Napi::Value bytes = reader.Get(in, "bytes");
    if (!reader.ok()) {
      return env.Undefined();
    }
    if (!bytes.IsTypedArray() || bytes.As<Napi::TypedArray>().TypedArrayType() != napi_uint8_array) {
      ThrowArgumentTypeError(env, "picture.bytes", "must be a Uint8Array");
      return env.Undefined();
    }
    const Napi::Uint8Array u8 = bytes.As<Napi::Uint8Array>();
    // An empty array may have no backing store; a non-NULL pointer lets the C ABI report the empty image.
    static const uint8_t kEmptyInput = 0;
    picture.len = u8.ElementLength();
    picture.bytes = picture.len == 0 ? &kEmptyInput : u8.Data();
    picture.alt_text = reader.String(in, "altText", &alt_text) ? alt_text.c_str() : nullptr;
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc = fm_sheet_set_cell_picture(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                                   present ? &picture : nullptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::InsertImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  BytesWithOpts args;
  if (!ReadBytesWithOpts(info, args)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  std::string name;
  std::string description;
  const fm_drawing_placement_spec_t opts = PullPlacement(reader, args.spec, name, description);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_add_picture(handle_, ArgU32(info, 0), args.data, args.len, &opts, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

Napi::Value Workbook::RemoveDrawing(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_drawing);
}

Napi::Value Workbook::UpdateDrawing(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  std::string name;
  std::string description;
  const fm_drawing_patch_t patch = PullDrawingPatch(reader, ArgObject(info, 2), name, description);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_drawing(handle_, ArgU32(info, 0), ArgU32(info, 1), &patch), env.Undefined());
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
  if (!ReadBytesWithOpts(info, args)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  const uint32_t flags = reader.Bool(args.spec, "assignNewId", false) ? FM_DRAWING_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_restore_drawing(handle_, ArgU32(info, 0), args.data, args.len, flags, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

}  // namespace formulon_node
