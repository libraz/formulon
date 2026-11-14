//
// JsWorkbook drawing surface: probing, listing, reading and inserting
// pictures, and updating, reordering, snapshotting, restoring and removing any
// drawing object.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- Drawing objects ---------------------------------------------------

namespace {

/// The header info of a probed image.
emscripten::val image_info_to_val(const fm_image_info_t& info) {
  emscripten::val o = emscripten::val::object();
  o.set("format", info.format);
  o.set("widthPx", info.width_px);
  o.set("heightPx", info.height_px);
  return o;
}

emscripten::val drawing_anchor_to_val(const fm_drawing_anchor_t& a) {
  emscripten::val o = emscripten::val::object();
  o.set("kind", a.kind);
  o.set("editAs", a.edit_as);
  o.set("fromRow", a.from_row);
  o.set("fromCol", a.from_col);
  o.set("fromRowOffsetEmu", static_cast<double>(a.from_row_offset_emu));
  o.set("fromColOffsetEmu", static_cast<double>(a.from_col_offset_emu));
  o.set("toRow", a.to_row);
  o.set("toCol", a.to_col);
  o.set("toRowOffsetEmu", static_cast<double>(a.to_row_offset_emu));
  o.set("toColOffsetEmu", static_cast<double>(a.to_col_offset_emu));
  o.set("widthEmu", static_cast<double>(a.width_emu));
  o.set("heightEmu", static_cast<double>(a.height_emu));
  return o;
}

emscripten::val drawing_object_to_val(const fm_drawing_object_info_t& d) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", d.object_id);
  o.set("kind", d.kind);
  o.set("hidden", d.hidden != 0);
  o.set("anchor", drawing_anchor_to_val(d.anchor));
  js_set_cstr_fields(o, {{"name", d.name}});
  js_set_cstr_or_null(o, "description", d.description);
  return o;
}

fm_status_t drawing_object_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_drawing_object_info_t d = FM_STRUCT_INIT(fm_drawing_object_info_t);
  const fm_status_t rc = fm_sheet_get_drawing_object_at(handle, sheet, index, &d);
  if (rc == 0) {
    *out = drawing_object_to_val(d);
  }
  return rc;
}

/// Reads a `DrawingAnchorInput` (`kind`, `editAs`, the `from*` / `to*` cells and EMU offsets, `widthEmu`,
/// `heightEmu`); an absent key is zero, so an absent `kind` is the operation's default. `path` names `anchor`.
fm_drawing_anchor_t pull_drawing_anchor(JsNarrowNumericReader& reader, const emscripten::val& anchor,
                                        const std::string& path) {
  const auto label = [&](const char* key) { return path + "." + key; };
  fm_drawing_anchor_t out{};
  out.kind = reader.i32(anchor, "kind", FM_ANCHOR_KIND_DEFAULT, label("kind").c_str());
  out.edit_as = reader.i32(anchor, "editAs", FM_ANCHOR_EDIT_AS_DEFAULT, label("editAs").c_str());
  out.from_row = reader.u32(anchor, "fromRow", 0U, label("fromRow").c_str());
  out.from_col = reader.u32(anchor, "fromCol", 0U, label("fromCol").c_str());
  out.from_row_offset_emu = reader.i64(anchor, "fromRowOffsetEmu", 0, label("fromRowOffsetEmu").c_str());
  out.from_col_offset_emu = reader.i64(anchor, "fromColOffsetEmu", 0, label("fromColOffsetEmu").c_str());
  out.to_row = reader.u32(anchor, "toRow", 0U, label("toRow").c_str());
  out.to_col = reader.u32(anchor, "toCol", 0U, label("toCol").c_str());
  out.to_row_offset_emu = reader.i64(anchor, "toRowOffsetEmu", 0, label("toRowOffsetEmu").c_str());
  out.to_col_offset_emu = reader.i64(anchor, "toColOffsetEmu", 0, label("toColOffsetEmu").c_str());
  out.width_emu = reader.i64(anchor, "widthEmu", 0, label("widthEmu").c_str());
  out.height_emu = reader.i64(anchor, "heightEmu", 0, label("heightEmu").c_str());
  return out;
}

}  // namespace

fm_drawing_placement_spec_t pull_drawing_placement(JsNarrowNumericReader& reader, const emscripten::val& input,
                                                   const char* param, std::string& name_store,
                                                   std::string& description_store) {
  const std::string path(param);
  const emscripten::val options = js_object_or_empty(input);
  fm_drawing_placement_spec_t placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  placement.name = reader.optional_string(options, "name", name_store, (path + ".name").c_str());
  placement.description =
      reader.optional_string(options, "description", description_store, (path + ".description").c_str());
  placement.hidden = js_tristate(reader, options, "hidden", (path + ".hidden").c_str());
  const emscripten::val anchor = reader.value(options, "anchor", (path + ".anchor").c_str());
  placement.anchor = pull_drawing_anchor(reader, js_object_or_empty(anchor), path + ".anchor");
  return placement;
}

fm_drawing_patch_t pull_drawing_patch(JsNarrowNumericReader& reader, const emscripten::val& patch, const char* param,
                                      std::string& name_store, std::string& description_store) {
  const std::string path(param);
  const auto label = [&](const char* key) { return path + "." + key; };
  fm_drawing_patch_t p = FM_STRUCT_INIT(fm_drawing_patch_t);
  js_patch_string(reader, patch, "name", label("name").c_str(), false, name_store, &p.name_op, &p.name);
  js_patch_string(reader, patch, "description", label("description").c_str(), true, description_store,
                  &p.description_op, &p.description);
  js_patch_flag(reader, patch, "hidden", label("hidden").c_str(), &p.hidden_op, &p.hidden);
  const emscripten::val anchor_value = reader.value(patch, "anchor", label("anchor").c_str());
  const emscripten::val anchor =
      js_object_or_empty(js_patch_key(reader, anchor_value, label("anchor").c_str(), false) == JsPatchKey::kSet
                             ? anchor_value
                             : emscripten::val::undefined());
  fm_drawing_anchor_patch_t& a = p.anchor;
  // No anchor field can be cleared: `null` is a type error, like any other flag.
  const auto field = [&](const char* key, fm_patch_op_t* op, auto* value, auto read) {
    const std::string name = label("anchor") + "." + key;
    const emscripten::val v = reader.value(anchor, key, name.c_str());
    if (js_patch_key(reader, v, name.c_str(), false) == JsPatchKey::kSet) {
      *op = FM_PATCH_OP_SET;
      *value = read(key, name.c_str());
    }
  };
  const auto u32 = [&](const char* key, const char* name) { return reader.u32(anchor, key, 0U, name); };
  const auto i32 = [&](const char* key, const char* name) { return reader.i32(anchor, key, 0, name); };
  const auto i64 = [&](const char* key, const char* name) { return reader.i64(anchor, key, 0, name); };
  field("kind", &a.kind_op, &a.kind, i32);
  field("editAs", &a.edit_as_op, &a.edit_as, i32);
  field("fromRow", &a.from_row_op, &a.from_row, u32);
  field("fromCol", &a.from_col_op, &a.from_col, u32);
  field("fromRowOffsetEmu", &a.from_row_offset_emu_op, &a.from_row_offset_emu, i64);
  field("fromColOffsetEmu", &a.from_col_offset_emu_op, &a.from_col_offset_emu, i64);
  field("toRow", &a.to_row_op, &a.to_row, u32);
  field("toCol", &a.to_col_op, &a.to_col, u32);
  field("toRowOffsetEmu", &a.to_row_offset_emu_op, &a.to_row_offset_emu, i64);
  field("toColOffsetEmu", &a.to_col_offset_emu_op, &a.to_col_offset_emu, i64);
  field("widthEmu", &a.width_emu_op, &a.width_emu, i64);
  field("heightEmu", &a.height_emu_op, &a.height_emu, i64);
  return p;
}

emscripten::val probe_image(emscripten::val bytes) {
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return emscripten::val::undefined();
  }
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  if (!succeeded(fm_get_image_info(data.data(), data.size(), &info))) {
    return emscripten::val::undefined();
  }
  return image_info_to_val(info);
}

emscripten::val JsWorkbook::listDrawingObjects(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, sheet, &fm_sheet_get_drawing_object_count, &drawing_object_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

emscripten::val JsWorkbook::getDrawingObject(uint32_t sheet, uint32_t objectId) const {
  fm_drawing_object_info_t d = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_drawing_object(handle_, sheet, objectId, &present, &d) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? drawing_object_to_val(d) : emscripten::val::null());
}

emscripten::val JsWorkbook::getImage(uint32_t sheet, uint32_t objectId) const {
  int32_t present = 0;
  const uint8_t* data = nullptr;
  size_t len = 0;
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_picture(handle_, sheet, objectId, &present, &data, &len, &info)
                             : kBindingInvalidHandle;
  if (rc != 0 || present == 0) {
    return optional_or_null(rc, present, emscripten::val::null());
  }
  emscripten::val image = emscripten::val::object();
  image.set("format", info.format);
  image.set("widthPx", info.width_px);
  image.set("heightPx", info.height_px);
  image.set("bytes", bytes_to_val(data, len));
  return optional_or_null(rc, present, image);
}

emscripten::val JsWorkbook::getCellPicture(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t present = 0;
  fm_cell_picture_t picture = FM_STRUCT_INIT(fm_cell_picture_t);
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_cell_picture(handle_, sheet, row, col, &present, &picture)
                                            : kBindingInvalidHandle;
  if (rc != 0 || present == 0) {
    return optional_or_null(rc, present, emscripten::val::null());
  }
  emscripten::val out = emscripten::val::object();
  out.set("format", picture.format);
  out.set("widthPx", picture.width_px);
  out.set("heightPx", picture.height_px);
  out.set("bytes", bytes_to_val(picture.bytes, static_cast<std::size_t>(picture.len)));
  js_set_cstr_or_null(out, "altText", picture.alt_text);
  return optional_or_null(rc, present, out);
}

void JsWorkbook::setCellPicture(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val picture) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (picture.isNull()) {
    succeeded(fm_sheet_set_cell_picture(handle_, sheet, row, col, nullptr));
    return;
  }
  JsNarrowNumericReader reader("Sheet.setCellPicture");
  if (!reader.require_object(picture, "picture")) {
    return;
  }
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(picture["bytes"], &data)) {
    return;
  }
  std::string alt_store;
  fm_cell_picture_t input = FM_STRUCT_INIT(fm_cell_picture_t);
  input.alt_text = reader.optional_string(picture, "altText", alt_store, "picture.altText");
  if (!reader.ok()) {
    return;
  }
  // An empty vector has no data pointer; the engine refuses empty data, not a NULL pointer.
  static const uint8_t kEmpty = 0;
  input.bytes = data.empty() ? &kEmpty : data.data();
  input.len = data.size();
  succeeded(fm_sheet_set_cell_picture(handle_, sheet, row, col, &input));
}

uint32_t JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  std::string name_store;
  std::string description_store;
  JsNarrowNumericReader reader("Sheet.addPicture");
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return 0U;
  }
  const fm_drawing_placement_spec_t ins = pull_drawing_placement(reader, input, "input", name_store, description_store);
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t id = 0;
  if (!succeeded(fm_sheet_add_picture(handle_, sheet, data.data(), data.size(), &ins, &id))) {
    return 0U;
  }
  return id;
}

void JsWorkbook::removeDrawing(uint32_t sheet, uint32_t objectId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_drawing(handle_, sheet, objectId) : kBindingInvalidHandle);
}

void JsWorkbook::updateDrawing(uint32_t sheet, uint32_t objectId, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Drawing.update");
  std::string name_store;
  std::string description_store;
  const fm_drawing_patch_t p =
      pull_drawing_patch(reader, js_object_or_empty(patch), "patch", name_store, description_store);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_drawing(handle_, sheet, objectId, &p));
}

void JsWorkbook::setDrawingZOrder(uint32_t sheet, uint32_t objectId, uint32_t zOrderIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_set_drawing_z_order(handle_, sheet, objectId, zOrderIndex)
                               : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::snapshotDrawing(uint32_t sheet, uint32_t objectId) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_drawing(handle_, sheet, objectId, &data, &len);
  return snapshot_bytes(rc, data, len);
}

uint32_t JsWorkbook::restoreDrawing(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.restoreDrawing");
  const emscripten::val options = js_object_or_empty(opts);
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return 0U;
  }
  const uint32_t flags = reader.boolean(options, "assignNewId", false, "assignNewId") ? FM_DRAWING_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t id = 0;
  if (!succeeded(fm_sheet_restore_drawing(handle_, sheet, data.data(), data.size(), flags, &id))) {
    return 0U;
  }
  return id;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
