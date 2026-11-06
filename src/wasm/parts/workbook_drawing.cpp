//
// JsWorkbook drawing surface: probing, listing, reading and inserting
// pictures, and re-anchoring, reordering, snapshotting and restoring any
// drawing object (picture or chart).

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

/// Envelope for the image calls: `status` plus the header info, zeroed on failure.
emscripten::val image_info_to_val(JsStatus status, const fm_image_info_t& info) {
  const bool ok = status.ok;
  emscripten::val o = emscripten::val::object();
  o.set("status", status);
  o.set("format", ok ? info.format : 0);
  o.set("pxWidth", ok ? info.px_width : 0U);
  o.set("pxHeight", ok ? info.px_height : 0U);
  return o;
}

/// Envelope for the calls that place a picture: `objectId`, zero unless the call succeeded.
emscripten::val image_id_result(JsStatus status, uint32_t object_id) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", object_id);
  o.set("status", status);
  return o;
}

emscripten::val drawing_object_to_val(const fm_drawing_object_t& d) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", d.object_id);
  o.set("kind", d.kind);
  o.set("anchorKind", d.anchor_kind);
  o.set("editAs", d.edit_as);
  o.set("fromRow", d.from_row);
  o.set("fromCol", d.from_col);
  o.set("fromRowOffset", static_cast<double>(d.from_row_offset));
  o.set("fromColOffset", static_cast<double>(d.from_col_offset));
  o.set("toRow", d.to_row);
  o.set("toCol", d.to_col);
  o.set("toRowOffset", static_cast<double>(d.to_row_offset));
  o.set("toColOffset", static_cast<double>(d.to_col_offset));
  o.set("cx", static_cast<double>(d.cx));
  o.set("cy", static_cast<double>(d.cy));
  o.set("imageFormat", d.image_format);
  js_set_cstr_fields(o, {{"name", d.name}, {"descr", d.descr}, {"mediaPath", d.media_path}});
  return o;
}

fm_status_t drawing_object_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_drawing_object_t d{};
  const fm_status_t rc = fm_sheet_get_drawing_object_at(handle, sheet, index, &d);
  if (rc == 0) {
    *out = drawing_object_to_val(d);
  }
  return rc;
}

}  // namespace

fm_drawing_anchor_t pull_drawing_anchor(JsNarrowNumericReader& reader, const emscripten::val& placement,
                                        int32_t default_anchor_kind) {
  fm_drawing_anchor_t anchor{};
  anchor.anchor_kind = reader.i32(placement, "anchorKind", default_anchor_kind, "placement.anchorKind");
  anchor.edit_as = reader.i32(placement, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "placement.editAs");
  anchor.row = reader.u32(placement, "row", 0U, "placement.row");
  anchor.col = reader.u32(placement, "col", 0U, "placement.col");
  anchor.row_offset_emu = reader.i64(placement, "rowOffsetEmu", 0, "placement.rowOffsetEmu");
  anchor.col_offset_emu = reader.i64(placement, "colOffsetEmu", 0, "placement.colOffsetEmu");
  anchor.width_emu = reader.i64(placement, "widthEmu", 0, "placement.widthEmu");
  anchor.height_emu = reader.i64(placement, "heightEmu", 0, "placement.heightEmu");
  return anchor;
}

void set_insert_placement(fm_image_insert_t& ins, const fm_drawing_anchor_t& anchor) {
  ins.anchor_kind = anchor.anchor_kind;
  ins.edit_as = anchor.edit_as;
  ins.row = anchor.row;
  ins.col = anchor.col;
  ins.row_offset_emu = anchor.row_offset_emu;
  ins.col_offset_emu = anchor.col_offset_emu;
  ins.width_emu = anchor.width_emu;
  ins.height_emu = anchor.height_emu;
}

emscripten::val JsWorkbook::probeImage(emscripten::val bytes) const {
  if (handle_ == nullptr) {
    return image_info_to_val(invalid_handle_status(), fm_image_info_t{});
  }
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_info_to_val(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), fm_image_info_t{});
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  fm_image_info_t info{};
  const fm_status_t rc = fm_workbook_get_image_info(handle_, data.data(), data.size(), &info);
  return image_info_to_val(status_from_rc(rc), info);
}

emscripten::val JsWorkbook::listDrawingObjects(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  arr.set("status", status_from_rc(js_collect_list(handle_, sheet, &fm_sheet_get_drawing_object_count,
                                                   &drawing_object_item, arr)));
  return arr;
}

emscripten::val JsWorkbook::getDrawingObject(uint32_t sheet, uint32_t objectId) const {
  fm_drawing_object_t d{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_drawing_object(handle_, sheet, objectId, &present, &d) : kBindingInvalidHandle;
  return optional_result(rc, present, rc == 0 && present != 0 ? drawing_object_to_val(d) : emscripten::val::null());
}

emscripten::val JsWorkbook::getImage(uint32_t sheet, uint32_t objectId) const {
  int32_t present = 0;
  const uint8_t* data = nullptr;
  size_t len = 0;
  fm_image_info_t info{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_image(handle_, sheet, objectId, &present, &data, &len, &info)
                                            : kBindingInvalidHandle;
  if (rc != 0 || present == 0) {
    return optional_result(rc, present, emscripten::val::null());
  }
  emscripten::val image = emscripten::val::object();
  image.set("format", info.format);
  image.set("pxWidth", info.px_width);
  image.set("pxHeight", info.px_height);
  image.set("bytes", bytes_to_val(data, len));
  return optional_result(rc, present, image);
}

emscripten::val JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    return image_id_result(invalid_handle_status(), 0U);
  }
  std::string name_store;
  std::string descr_store;
  JsNarrowNumericReader reader("addImage");
  const emscripten::val options = js_object_or_empty(opts);
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_id_result(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), 0U);
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  fm_image_insert_t ins{};
  ins.name = reader.optional_string(options, "name", name_store, "image.name");
  ins.descr = reader.optional_string(options, "descr", descr_store, "image.descr");
  set_insert_placement(ins, pull_drawing_anchor(reader, options, FM_ANCHOR_KIND_ONE_CELL));
  if (!reader.ok()) {
    return image_id_result(reader.failure(), 0U);
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_add_image(handle_, sheet, data.data(), data.size(), &ins, &id);
  return image_id_result(status_from_rc(rc), rc == 0 ? id : 0U);
}

JsStatus JsWorkbook::removeDrawing(uint32_t sheet, uint32_t objectId) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_remove_drawing(handle_, sheet, objectId) : kBindingInvalidHandle);
}

JsStatus JsWorkbook::setDrawingName(uint32_t sheet, uint32_t objectId, const std::string& name) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_set_drawing_name(handle_, sheet, objectId, name.c_str())
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::setDrawingAnchor(uint32_t sheet, uint32_t objectId, emscripten::val placement) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setDrawingAnchor");
  const emscripten::val options = js_object_or_empty(placement);
  const fm_drawing_anchor_t anchor = pull_drawing_anchor(reader, options, FM_ANCHOR_KIND_ONE_CELL);
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_sheet_set_drawing_anchor(handle_, sheet, objectId, &anchor));
}

JsStatus JsWorkbook::setDrawingZOrder(uint32_t sheet, uint32_t objectId, uint32_t index) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_set_drawing_z_order(handle_, sheet, objectId, index)
                                           : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::snapshotDrawing(uint32_t sheet, uint32_t objectId) const {
  if (handle_ == nullptr) {
    return snapshot_bytes_result(kBindingInvalidHandle, nullptr, 0);
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_drawing(handle_, sheet, objectId, &data, &len);
  return snapshot_bytes_result(rc, data, len);
}

emscripten::val JsWorkbook::restoreDrawing(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    return image_id_result(invalid_handle_status(), 0U);
  }
  JsNarrowNumericReader reader("restoreDrawing");
  const emscripten::val options = js_object_or_empty(opts);
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_id_result(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), 0U);
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  const uint32_t flags = reader.boolean(options, "newId", false, "options.newId") ? FM_DRAWING_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return image_id_result(reader.failure(), 0U);
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_restore_drawing(handle_, sheet, data.data(), data.size(), flags, &id);
  return image_id_result(status_from_rc(rc), rc == 0 ? id : 0U);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
