//
// JsWorkbook table-slicer surface: listing, adding, updating and removing the
// slicers of a worksheet.

#include <emscripten/val.h>

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

emscripten::val slicer_to_val(const fm_slicer_info_t& s) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", s.object_id);
  o.set("columnCount", s.column_count);
  o.set("sortOrder", s.sort_order);
  o.set("crossFilter", s.cross_filter);
  o.set("anchorKind", s.anchor_kind);
  o.set("fromRow", s.from_row);
  o.set("fromCol", s.from_col);
  o.set("fromRowOffsetEmu", static_cast<double>(s.from_row_offset_emu));
  o.set("fromColOffsetEmu", static_cast<double>(s.from_col_offset_emu));
  o.set("toRow", s.to_row);
  o.set("toCol", s.to_col);
  o.set("toRowOffsetEmu", static_cast<double>(s.to_row_offset_emu));
  o.set("toColOffsetEmu", static_cast<double>(s.to_col_offset_emu));
  o.set("rowHeightEmu", static_cast<double>(s.row_height_emu));
  js_set_cstr_fields(o, {{"name", s.name},
                         {"caption", s.caption},
                         {"cacheName", s.cache_name},
                         {"tableName", s.table_name},
                         {"columnName", s.column_name}});
  js_set_cstr_or_null(o, "style", s.style);
  return o;
}

struct SlicerText {
  std::string name;
  std::string caption;
  std::string style;
};

/// Reads the fields `spec` carries; an omitted field is `KEEP`, `null` on `style` is `CLEAR`.
fm_slicer_patch_t pull_slicer(JsNarrowNumericReader& reader, const emscripten::val& spec, SlicerText& store) {
  fm_slicer_patch_t s = FM_STRUCT_INIT(fm_slicer_patch_t);
  const auto has = [&](const char* key) { return reader.ok() && js_value_present(reader.value(spec, key, key)); };
  if (has("name")) {
    s.name_op = FM_PATCH_OP_SET;
    s.name = reader.optional_string(spec, "name", store.name, "slicer.name");
  }
  if (has("caption")) {
    s.caption_op = FM_PATCH_OP_SET;
    s.caption = reader.optional_string(spec, "caption", store.caption, "slicer.caption");
  }
  js_patch_string(reader, spec, "style", "slicer.style", true, store.style, &s.style_op, &s.style);
  if (has("columnCount")) {
    s.column_count_op = FM_PATCH_OP_SET;
    s.column_count = reader.u32(spec, "columnCount", 1U, "slicer.columnCount");
  }
  if (has("rowHeightEmu")) {
    s.row_height_emu_op = FM_PATCH_OP_SET;
    s.row_height_emu = reader.i64(spec, "rowHeightEmu", 0, "slicer.rowHeightEmu");
  }
  const auto offset = [&](const char* key, const char* field, fm_patch_op_t* op, int64_t* value) {
    if (has(key)) {
      *op = FM_PATCH_OP_SET;
      *value = reader.i64(spec, key, 0, field);
    }
  };
  // The two-cell anchor is all or nothing: any of its four cells makes the other three required.
  if (has("fromRow") || has("fromCol") || has("toRow") || has("toCol")) {
    s.from_row_op = FM_PATCH_OP_SET;
    s.from_row = reader.required_u32(spec, "fromRow", "slicer.fromRow");
    s.from_col_op = FM_PATCH_OP_SET;
    s.from_col = reader.required_u32(spec, "fromCol", "slicer.fromCol");
    s.to_row_op = FM_PATCH_OP_SET;
    s.to_row = reader.required_u32(spec, "toRow", "slicer.toRow");
    s.to_col_op = FM_PATCH_OP_SET;
    s.to_col = reader.required_u32(spec, "toCol", "slicer.toCol");
  }
  offset("fromRowOffsetEmu", "slicer.fromRowOffsetEmu", &s.from_row_offset_emu_op, &s.from_row_offset_emu);
  offset("fromColOffsetEmu", "slicer.fromColOffsetEmu", &s.from_col_offset_emu_op, &s.from_col_offset_emu);
  offset("toRowOffsetEmu", "slicer.toRowOffsetEmu", &s.to_row_offset_emu_op, &s.to_row_offset_emu);
  offset("toColOffsetEmu", "slicer.toColOffsetEmu", &s.to_col_offset_emu_op, &s.to_col_offset_emu);
  if (has("sortOrder")) {
    s.sort_order_op = FM_PATCH_OP_SET;
    s.sort_order = reader.i32(spec, "sortOrder", 0, "slicer.sortOrder");
  }
  if (has("crossFilter")) {
    s.cross_filter_op = FM_PATCH_OP_SET;
    s.cross_filter = reader.i32(spec, "crossFilter", 0, "slicer.crossFilter");
  }
  return s;
}

fm_status_t slicer_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_slicer_info_t s = FM_STRUCT_INIT(fm_slicer_info_t);
  const fm_status_t rc = fm_sheet_get_slicer_at(handle, sheet, index, &s);
  if (rc == 0) {
    *out = slicer_to_val(s);
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listSlicers(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, sheet, &fm_sheet_get_slicer_count, &slicer_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

uint32_t JsWorkbook::addSlicer(uint32_t sheet, uint32_t tableSheet, uint32_t table, uint32_t column,
                               emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.addSlicer");
  SlicerText store;
  const fm_slicer_patch_t slicer = pull_slicer(reader, spec, store);
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t id = 0;
  if (!succeeded(fm_sheet_add_slicer(handle_, sheet, tableSheet, table, column, &slicer, &id))) {
    return 0U;
  }
  return id;
}

emscripten::val JsWorkbook::getSlicer(uint32_t sheet, uint32_t objectId) const {
  fm_slicer_info_t s = FM_STRUCT_INIT(fm_slicer_info_t);
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_slicer(handle_, sheet, objectId, &present, &s) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? slicer_to_val(s) : emscripten::val::null());
}

void JsWorkbook::updateSlicer(uint32_t sheet, uint32_t objectId, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Slicer.update");
  SlicerText store;
  const fm_slicer_patch_t slicer = pull_slicer(reader, spec, store);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_slicer(handle_, sheet, objectId, &slicer));
}

void JsWorkbook::removeSlicer(uint32_t sheet, uint32_t objectId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_slicer(handle_, sheet, objectId) : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
