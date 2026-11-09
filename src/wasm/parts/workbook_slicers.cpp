//
// JsWorkbook table-slicer surface: reading, adding and updating the slicers of
// a worksheet.

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
  o.set("rowHeightEmu", static_cast<double>(s.row_height_emu));
  js_set_cstr_fields(o, {{"caption", s.caption},
                         {"cacheName", s.cache_name},
                         {"tableName", s.table_name},
                         {"columnName", s.column_name}});
  js_set_cstr_or_null(o, "style", s.style);
  return o;
}

struct SlicerText {
  std::string name;
  std::string description;
  std::string caption;
  std::string style;
};

/// Reads the slicer fields `spec` carries into a patch; an omitted field is `KEEP`, `null` on `style` is `CLEAR`.
fm_slicer_patch_t pull_slicer_patch(JsNarrowNumericReader& reader, const emscripten::val& spec, SlicerText& store) {
  fm_slicer_patch_t s = FM_STRUCT_INIT(fm_slicer_patch_t);
  js_patch_string(reader, spec, "caption", "slicer.caption", false, store.caption, &s.caption_op, &s.caption);
  js_patch_string(reader, spec, "style", "slicer.style", true, store.style, &s.style_op, &s.style);
  const auto has = [&](const char* key) { return reader.ok() && js_value_present(reader.value(spec, key, key)); };
  if (has("columnCount")) {
    s.column_count_op = FM_PATCH_OP_SET;
    s.column_count = reader.u32(spec, "columnCount", 1U, "slicer.columnCount");
  }
  if (has("rowHeightEmu")) {
    s.row_height_emu_op = FM_PATCH_OP_SET;
    s.row_height_emu = reader.i64(spec, "rowHeightEmu", 0, "slicer.rowHeightEmu");
  }
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

/// Reads a slicer spec: the placement keys (`name`, `description`, `hidden`, `anchor`) and the slicer fields; an
/// absent key takes the default of a new slicer.
fm_slicer_spec_t pull_slicer_spec(JsNarrowNumericReader& reader, const emscripten::val& spec, SlicerText& store) {
  fm_slicer_spec_t s = FM_STRUCT_INIT(fm_slicer_spec_t);
  s.placement = pull_drawing_placement(reader, spec, store.name, store.description);
  s.column_count = reader.u32(spec, "columnCount", 0U, "slicer.columnCount");
  s.sort_order = reader.i32(spec, "sortOrder", FM_SLICER_SORT_ORDER_ASCENDING, "slicer.sortOrder");
  s.cross_filter =
      reader.i32(spec, "crossFilter", FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_DATA_AT_TOP, "slicer.crossFilter");
  s.row_height_emu = reader.i64(spec, "rowHeightEmu", 0, "slicer.rowHeightEmu");
  s.caption = reader.optional_string(spec, "caption", store.caption, "slicer.caption");
  s.style = reader.optional_string(spec, "style", store.style, "slicer.style");
  return s;
}

}  // namespace

uint32_t JsWorkbook::addSlicer(uint32_t sheet, uint32_t tableSheet, uint32_t table, uint32_t column,
                               emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.addSlicer");
  SlicerText store;
  const fm_slicer_spec_t slicer = pull_slicer_spec(reader, spec, store);
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
  const emscripten::val options = js_object_or_empty(spec);
  const fm_slicer_patch_t slicer = pull_slicer_patch(reader, options, store);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_slicer(handle_, sheet, objectId, &slicer));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
