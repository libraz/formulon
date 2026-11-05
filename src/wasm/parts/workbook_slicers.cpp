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

emscripten::val slicer_to_val(const fm_slicer_t& s) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", s.object_id);
  o.set("columnCount", s.column_count);
  o.set("sortOrder", s.sort_order);
  o.set("crossFilter", s.cross_filter);
  o.set("anchorKind", s.anchor_kind);
  o.set("fromRow", s.from_row);
  o.set("fromCol", s.from_col);
  o.set("fromRowOff", static_cast<double>(s.from_row_off));
  o.set("fromColOff", static_cast<double>(s.from_col_off));
  o.set("toRow", s.to_row);
  o.set("toCol", s.to_col);
  o.set("toRowOff", static_cast<double>(s.to_row_off));
  o.set("toColOff", static_cast<double>(s.to_col_off));
  o.set("rowHeight", static_cast<double>(s.row_height));
  js_set_cstr_fields(o, {{"name", s.name},
                         {"caption", s.caption},
                         {"cacheName", s.cache_name},
                         {"tableName", s.table_name},
                         {"columnName", s.column_name},
                         {"style", s.style}});
  return o;
}

struct SlicerText {
  std::string name;
  std::string caption;
  std::string style;
};

/// Reads the fields `spec` carries; the `fields` mask names exactly those.
fm_slicer_t pull_slicer(JsNarrowNumericReader& reader, const emscripten::val& spec, SlicerText& store) {
  fm_slicer_t s{};
  const auto has = [&](const char* key) { return reader.ok() && js_value_present(reader.value(spec, key, key)); };
  if (has("name")) {
    s.fields |= FM_SLICER_FIELD_NAME;
    s.name = reader.optional_string(spec, "name", store.name, "slicer.name");
  }
  if (has("caption")) {
    s.fields |= FM_SLICER_FIELD_CAPTION;
    s.caption = reader.optional_string(spec, "caption", store.caption, "slicer.caption");
  }
  if (has("style")) {
    s.fields |= FM_SLICER_FIELD_STYLE;
    s.style = reader.optional_string(spec, "style", store.style, "slicer.style");
  }
  if (has("columnCount")) {
    s.fields |= FM_SLICER_FIELD_COLUMN_COUNT;
    s.column_count = reader.u32(spec, "columnCount", 1U, "slicer.columnCount");
  }
  if (has("rowHeight")) {
    s.fields |= FM_SLICER_FIELD_ROW_HEIGHT;
    s.row_height = reader.i64(spec, "rowHeight", 0, "slicer.rowHeight");
  }
  if (has("fromRow")) {
    s.fields |= FM_SLICER_FIELD_ANCHOR;
    s.from_row = reader.u32(spec, "fromRow", 0U, "slicer.fromRow");
    s.from_col = reader.u32(spec, "fromCol", 0U, "slicer.fromCol");
    s.from_row_off = reader.i64(spec, "fromRowOff", 0, "slicer.fromRowOff");
    s.from_col_off = reader.i64(spec, "fromColOff", 0, "slicer.fromColOff");
    s.to_row = reader.u32(spec, "toRow", 0U, "slicer.toRow");
    s.to_col = reader.u32(spec, "toCol", 0U, "slicer.toCol");
    s.to_row_off = reader.i64(spec, "toRowOff", 0, "slicer.toRowOff");
    s.to_col_off = reader.i64(spec, "toColOff", 0, "slicer.toColOff");
  }
  if (has("sortOrder")) {
    s.fields |= FM_SLICER_FIELD_SORT_ORDER;
    s.sort_order = reader.i32(spec, "sortOrder", 0, "slicer.sortOrder");
  }
  if (has("crossFilter")) {
    s.fields |= FM_SLICER_FIELD_CROSS_FILTER;
    s.cross_filter = reader.i32(spec, "crossFilter", 0, "slicer.crossFilter");
  }
  return s;
}

fm_status_t slicer_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_slicer_t s{};
  const fm_status_t rc = fm_sheet_get_slicer_at(handle, sheet, index, &s);
  if (rc == 0) {
    *out = slicer_to_val(s);
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listSlicers(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  arr.set("status", status_from_rc(js_collect_list(handle_, sheet, &fm_sheet_get_slicer_count, &slicer_item, arr)));
  return arr;
}

emscripten::val JsWorkbook::addSlicer(uint32_t sheet, uint32_t table, uint32_t column, emscripten::val spec) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", 0U);
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    return o;
  }
  JsNarrowNumericReader reader("addSlicer");
  SlicerText store;
  const fm_slicer_t slicer = pull_slicer(reader, spec, store);
  if (!reader.ok()) {
    o.set("status", reader.failure());
    return o;
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_add_slicer(handle_, sheet, table, column, &slicer, &id);
  o.set("objectId", rc == 0 ? id : 0U);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getSlicer(uint32_t sheet, uint32_t objectId) const {
  fm_slicer_t s{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_slicer(handle_, sheet, objectId, &present, &s) : kBindingInvalidHandle;
  return optional_result(rc, present, rc == 0 && present != 0 ? slicer_to_val(s) : emscripten::val::null());
}

JsStatus JsWorkbook::updateSlicer(uint32_t sheet, uint32_t objectId, emscripten::val spec) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("updateSlicer");
  SlicerText store;
  const fm_slicer_t slicer = pull_slicer(reader, spec, store);
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_sheet_update_slicer(handle_, sheet, objectId, &slicer));
}

JsStatus JsWorkbook::removeSlicer(uint32_t sheet, uint32_t objectId) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_remove_slicer(handle_, sheet, objectId) : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
