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

/// The header info of a probed image.
emscripten::val image_info_to_val(const fm_image_info_t& info) {
  emscripten::val o = emscripten::val::object();
  o.set("format", info.format);
  o.set("pxWidth", info.px_width);
  o.set("pxHeight", info.px_height);
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
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return emscripten::val::undefined();
  }
  fm_image_info_t info{};
  if (!succeeded(fm_workbook_get_image_info(handle_, data.data(), data.size(), &info))) {
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
  fm_drawing_object_t d{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_drawing_object(handle_, sheet, objectId, &present, &d) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? drawing_object_to_val(d) : emscripten::val::null());
}

emscripten::val JsWorkbook::getImage(uint32_t sheet, uint32_t objectId) const {
  int32_t present = 0;
  const uint8_t* data = nullptr;
  size_t len = 0;
  fm_image_info_t info{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_image(handle_, sheet, objectId, &present, &data, &len, &info)
                                            : kBindingInvalidHandle;
  if (rc != 0 || present == 0) {
    return optional_or_null(rc, present, emscripten::val::null());
  }
  emscripten::val image = emscripten::val::object();
  image.set("format", info.format);
  image.set("pxWidth", info.px_width);
  image.set("pxHeight", info.px_height);
  image.set("bytes", bytes_to_val(data, len));
  return optional_or_null(rc, present, image);
}

uint32_t JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  std::string name_store;
  std::string descr_store;
  JsNarrowNumericReader reader("addImage");
  const emscripten::val options = js_object_or_empty(opts);
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return 0U;
  }
  fm_image_insert_t ins{};
  ins.name = reader.optional_string(options, "name", name_store, "image.name");
  ins.descr = reader.optional_string(options, "descr", descr_store, "image.descr");
  set_insert_placement(ins, pull_drawing_anchor(reader, options, FM_ANCHOR_KIND_ONE_CELL));
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t id = 0;
  if (!succeeded(fm_sheet_add_image(handle_, sheet, data.data(), data.size(), &ins, &id))) {
    return 0U;
  }
  return id;
}

void JsWorkbook::removeDrawing(uint32_t sheet, uint32_t objectId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_drawing(handle_, sheet, objectId) : kBindingInvalidHandle);
}

void JsWorkbook::setDrawingName(uint32_t sheet, uint32_t objectId, const std::string& name) {
  succeeded(handle_ != nullptr ? fm_sheet_set_drawing_name(handle_, sheet, objectId, name.c_str())
                               : kBindingInvalidHandle);
}

void JsWorkbook::setDrawingAnchor(uint32_t sheet, uint32_t objectId, emscripten::val placement) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setDrawingAnchor");
  const emscripten::val options = js_object_or_empty(placement);
  const fm_drawing_anchor_t anchor = pull_drawing_anchor(reader, options, FM_ANCHOR_KIND_ONE_CELL);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_drawing_anchor(handle_, sheet, objectId, &anchor));
}

void JsWorkbook::setDrawingZOrder(uint32_t sheet, uint32_t objectId, uint32_t index) {
  succeeded(handle_ != nullptr ? fm_sheet_set_drawing_z_order(handle_, sheet, objectId, index) : kBindingInvalidHandle);
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
  JsNarrowNumericReader reader("restoreDrawing");
  const emscripten::val options = js_object_or_empty(opts);
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return 0U;
  }
  const uint32_t flags = reader.boolean(options, "newId", false, "options.newId") ? FM_DRAWING_RESTORE_NEW_ID : 0U;
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
