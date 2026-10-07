//
// JsWorkbook drawing surface: probing, listing, reading, inserting,
// re-anchoring, reordering, snapshotting and restoring sheet pictures.

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

// ---- Drawing images ----------------------------------------------------

namespace {

/// Envelope for the image calls: `status` plus the header info, zeroed on failure.
emscripten::val image_info_to_val(fm_status_t rc, const fm_image_info& info) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("format", rc == 0 ? info.format : 0);
  o.set("pxWidth", rc == 0 ? info.px_width : 0U);
  o.set("pxHeight", rc == 0 ? info.px_height : 0U);
  return o;
}

emscripten::val image_info_binding_error(JsStatus status) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status);
  o.set("format", 0);
  o.set("pxWidth", 0U);
  o.set("pxHeight", 0U);
  return o;
}

/// Envelope for the calls that place a picture: `objectId`, zero unless the call succeeded.
emscripten::val image_id_result(JsStatus status, uint32_t object_id) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", object_id);
  o.set("status", status);
  return o;
}

emscripten::val drawing_object_to_val(const fm_drawing_object& d) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", d.object_id);
  o.set("kind", d.kind);
  o.set("anchorKind", d.anchor_kind);
  o.set("editAs", d.edit_as);
  o.set("fromRow", d.from_row);
  o.set("fromCol", d.from_col);
  o.set("fromRowOff", static_cast<double>(d.from_row_off));
  o.set("fromColOff", static_cast<double>(d.from_col_off));
  o.set("toRow", d.to_row);
  o.set("toCol", d.to_col);
  o.set("toRowOff", static_cast<double>(d.to_row_off));
  o.set("toColOff", static_cast<double>(d.to_col_off));
  o.set("cx", static_cast<double>(d.cx));
  o.set("cy", static_cast<double>(d.cy));
  o.set("imageFormat", d.image_format);
  js_set_cstr_fields(o, {{"name", d.name}, {"descr", d.descr}, {"mediaPath", d.media_path}});
  return o;
}

}  // namespace

emscripten::val JsWorkbook::probeImage(emscripten::val bytes) const {
  if (handle_ == nullptr) {
    return image_info_to_val(kBindingInvalidHandle, fm_image_info{});
  }
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_info_binding_error(binding_error_status(kInvalidArgument, bytes_result.message.c_str()));
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  fm_image_info info{};
  const fm_status_t rc = fm_workbook_probe_image(handle_, data.data(), data.size(), &info);
  return image_info_to_val(rc, info);
}

emscripten::val JsWorkbook::listDrawingObjects(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(kBindingInvalidHandle));
    return arr;
  }
  size_t count = 0;
  fm_status_t rc = fm_sheet_drawing_object_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_drawing_object d{};
    rc = fm_sheet_drawing_object_at(handle_, sheet, i, &d);
    if (rc == 0) {
      arr.call<void>("push", drawing_object_to_val(d));
    }
  }
  arr.set("status", status_from_rc(rc));
  return arr;
}

emscripten::val JsWorkbook::getImage(uint32_t sheet, uint32_t objectId) const {
  emscripten::val o;
  if (handle_ == nullptr) {
    o = image_info_to_val(kBindingInvalidHandle, fm_image_info{});
    o.set("bytes", bytes_to_val(nullptr, 0));
    return o;
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  fm_image_info info{};
  const fm_status_t rc = fm_sheet_get_image(handle_, sheet, objectId, &data, &len, &info);
  o = image_info_to_val(rc, info);
  o.set("bytes", bytes_to_val(rc == 0 ? data : nullptr, rc == 0 ? len : 0));
  return o;
}

emscripten::val JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    return image_id_result(error_status(kBindingInvalidHandle), 0U);
  }
  std::string name_store;
  std::string descr_store;
  JsNarrowNumericReader reader("insertImage");
  const emscripten::val options = opts.isUndefined() || opts.isNull() ? emscripten::val::object() : opts;
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_id_result(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), 0U);
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  fm_image_insert ins{};
  ins.name = reader.optional_string(options, "name", name_store, "image.name");
  ins.descr = reader.optional_string(options, "descr", descr_store, "image.descr");
  ins.anchor_kind = reader.i32(options, "anchorKind", FM_ANCHOR_KIND_ONE_CELL, "image.anchorKind");
  ins.edit_as = reader.i32(options, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "image.editAs");
  ins.row = reader.u32(options, "row", 0U, "image.row");
  ins.col = reader.u32(options, "col", 0U, "image.col");
  ins.row_off_emu = reader.i64(options, "rowOffEmu", 0, "image.rowOffEmu");
  ins.col_off_emu = reader.i64(options, "colOffEmu", 0, "image.colOffEmu");
  ins.width_emu = reader.i64(options, "widthEmu", 0, "image.widthEmu");
  ins.height_emu = reader.i64(options, "heightEmu", 0, "image.heightEmu");
  if (!reader.ok()) {
    return image_id_result(binding_error_status(kInvalidArgument, reader.message().c_str()), 0U);
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_insert_image(handle_, sheet, data.data(), data.size(), &ins, &id);
  return image_id_result(status_from_rc(rc), rc == 0 ? id : 0U);
}

JsStatus JsWorkbook::removeImage(uint32_t sheet, uint32_t objectId) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_remove_image(handle_, sheet, objectId));
}

JsStatus JsWorkbook::setImageAnchor(uint32_t sheet, uint32_t objectId, emscripten::val placement) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setImageAnchor");
  const emscripten::val options = placement.isUndefined() || placement.isNull() ? emscripten::val::object() : placement;
  fm_image_anchor anchor{};
  anchor.anchor_kind = reader.i32(options, "anchorKind", FM_ANCHOR_KIND_ONE_CELL, "placement.anchorKind");
  anchor.edit_as = reader.i32(options, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "placement.editAs");
  anchor.row = reader.u32(options, "row", 0U, "placement.row");
  anchor.col = reader.u32(options, "col", 0U, "placement.col");
  anchor.row_off_emu = reader.i64(options, "rowOffEmu", 0, "placement.rowOffEmu");
  anchor.col_off_emu = reader.i64(options, "colOffEmu", 0, "placement.colOffEmu");
  anchor.width_emu = reader.i64(options, "widthEmu", 0, "placement.widthEmu");
  anchor.height_emu = reader.i64(options, "heightEmu", 0, "placement.heightEmu");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_set_image_anchor(handle_, sheet, objectId, &anchor));
}

JsStatus JsWorkbook::setImageZOrder(uint32_t sheet, uint32_t objectId, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_set_image_z_order(handle_, sheet, objectId, index));
}

emscripten::val JsWorkbook::snapshotImage(uint32_t sheet, uint32_t objectId) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    o.set("bytes", bytes_to_val(nullptr, 0));
    return o;
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_image(handle_, sheet, objectId, &data, &len);
  o.set("status", status_from_rc(rc));
  o.set("bytes", bytes_to_val(rc == 0 ? data : nullptr, rc == 0 ? len : 0));
  fm_buffer_free(data);
  return o;
}

emscripten::val JsWorkbook::restoreImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    return image_id_result(error_status(kBindingInvalidHandle), 0U);
  }
  JsNarrowNumericReader reader("restoreImage");
  const emscripten::val options = opts.isUndefined() || opts.isNull() ? emscripten::val::object() : opts;
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_id_result(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), 0U);
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  const uint32_t flags = reader.boolean(options, "newId", false, "options.newId") ? FM_IMAGE_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return image_id_result(binding_error_status(kInvalidArgument, reader.message().c_str()), 0U);
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_restore_image(handle_, sheet, data.data(), data.size(), flags, &id);
  return image_id_result(status_from_rc(rc), rc == 0 ? id : 0U);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
