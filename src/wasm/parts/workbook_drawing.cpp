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
  o.set("widthPx", info.width_px);
  o.set("heightPx", info.height_px);
  return o;
}

emscripten::val drawing_object_to_val(const fm_drawing_object_info_t& d) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", d.object_id);
  o.set("kind", d.kind);
  o.set("anchorKind", d.anchor_kind);
  o.set("editAs", d.edit_as);
  o.set("fromRow", d.from_row);
  o.set("fromCol", d.from_col);
  o.set("fromRowOffsetEmu", static_cast<double>(d.from_row_offset_emu));
  o.set("fromColOffsetEmu", static_cast<double>(d.from_col_offset_emu));
  o.set("toRow", d.to_row);
  o.set("toCol", d.to_col);
  o.set("toRowOffsetEmu", static_cast<double>(d.to_row_offset_emu));
  o.set("toColOffsetEmu", static_cast<double>(d.to_col_offset_emu));
  o.set("widthEmu", static_cast<double>(d.width_emu));
  o.set("heightEmu", static_cast<double>(d.height_emu));
  o.set("format", d.image_format);
  js_set_cstr_fields(o, {{"name", d.name}, {"description", d.description}});
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

}  // namespace

fm_drawing_anchor_t pull_drawing_anchor(JsNarrowNumericReader& reader, const emscripten::val& anchor,
                                        int32_t default_anchor_kind) {
  fm_drawing_anchor_t out{};
  out.anchor_kind = reader.i32(anchor, "anchorKind", default_anchor_kind, "anchor.anchorKind");
  out.edit_as = reader.i32(anchor, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "anchor.editAs");
  out.row = reader.u32(anchor, "row", 0U, "anchor.row");
  out.col = reader.u32(anchor, "col", 0U, "anchor.col");
  out.row_offset_emu = reader.i64(anchor, "rowOffsetEmu", 0, "anchor.rowOffsetEmu");
  out.col_offset_emu = reader.i64(anchor, "colOffsetEmu", 0, "anchor.colOffsetEmu");
  out.width_emu = reader.i64(anchor, "widthEmu", 0, "anchor.widthEmu");
  out.height_emu = reader.i64(anchor, "heightEmu", 0, "anchor.heightEmu");
  return out;
}

fm_picture_insert_spec_t pull_image_insert(JsNarrowNumericReader& reader, const emscripten::val& input,
                                           int32_t default_anchor_kind, std::string& name_store,
                                           std::string& description_store) {
  const emscripten::val options = js_object_or_empty(input);
  fm_picture_insert_spec_t ins = FM_STRUCT_INIT(fm_picture_insert_spec_t);
  ins.name = reader.optional_string(options, "name", name_store, "name");
  ins.description = reader.optional_string(options, "description", description_store, "description");
  const emscripten::val anchor = reader.value(options, "anchor", "anchor");
  ins.anchor = pull_drawing_anchor(reader, js_object_or_empty(anchor), default_anchor_kind);
  return ins;
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

uint32_t JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  std::string name_store;
  std::string description_store;
  JsNarrowNumericReader reader("Sheet.addImage");
  std::vector<uint8_t> data;
  if (!val_to_bytes_checked(bytes, &data)) {
    return 0U;
  }
  const fm_picture_insert_spec_t ins =
      pull_image_insert(reader, input, FM_ANCHOR_KIND_ONE_CELL, name_store, description_store);
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

void JsWorkbook::setDrawingName(uint32_t sheet, uint32_t objectId, const std::string& name) {
  succeeded(handle_ != nullptr ? fm_sheet_set_drawing_name(handle_, sheet, objectId, name.c_str())
                               : kBindingInvalidHandle);
}

void JsWorkbook::setDrawingAnchor(uint32_t sheet, uint32_t objectId, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.setDrawingAnchor");
  const emscripten::val options = js_object_or_empty(input);
  const fm_drawing_anchor_t anchor = pull_drawing_anchor(reader, options, FM_ANCHOR_KIND_ONE_CELL);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_drawing_anchor(handle_, sheet, objectId, &anchor));
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
  const uint32_t flags =
      reader.boolean(options, "assignNewId", false, "options.assignNewId") ? FM_DRAWING_RESTORE_NEW_ID : 0U;
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
