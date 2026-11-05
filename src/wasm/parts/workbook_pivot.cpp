//
// JsWorkbook PivotCache + PivotTable mutators. Thin wrappers over the
// `fm_workbook_pivot_*` C ABI: numeric / boolean / string args map
// straight through; the spec-style mutators accept an `emscripten::val`
// and unpack the fields into the matching `fm_pivot_*_spec_t`. The
// std::string locals keep ownership for the borrowed `const char*`
// pointers handed to the C ABI; the spec is never retained beyond the
// call. The companion `pivotCount` / `pivotLayout` readers live in
// `parts/workbook_cells.cpp` so the iteration surface stays together.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

bool build_data_field_spec_checked(JsNarrowNumericReader& reader, const emscripten::val& spec,
                                   fm_pivot_data_field_spec_t& out, std::string& name_buf, std::string& nfmt_buf) {
  // Read the optional strings before the numeric fields. Once the reader
  // rejects a numeric value, no further access to the caller's object is
  // allowed before returning the binding error.
  out.name = reader.optional_string(spec, "name", name_buf, "pivotDataField.name");
  out.number_format = reader.optional_string(spec, "numberFormat", nfmt_buf, "pivotDataField.numberFormat");
  if (!reader.ok()) {
    return false;
  }
  out.field_index = reader.u32(spec, "fieldIndex", 0U, "pivotDataField.fieldIndex");
  out.aggregation =
      static_cast<fm_pivot_aggregation_t>(reader.u32(spec, "aggregation", 0U, "pivotDataField.aggregation"));
  out.show_as = static_cast<fm_pivot_show_values_as_t>(reader.u32(spec, "showAs", 0U, "pivotDataField.showAs"));
  out.show_as_base_field = reader.i32(spec, "showAsBaseField", -1, "pivotDataField.showAsBaseField");
  out.show_as_base_item = reader.i32(spec, "showAsBaseItem", -1, "pivotDataField.showAsBaseItem");
  return reader.ok();
}

}  // namespace

// ---- PivotCache --------------------------------------------------------

namespace {

fm_status_t count_pivot_caches(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_pivot_cache_count(handle, out);
}

fm_status_t pivot_cache_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  uint32_t cache_id = 0;
  const fm_status_t rc = fm_workbook_get_pivot_cache_id_at(handle, index, &cache_id);
  if (rc == 0) {
    *out = emscripten::val(cache_id);
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listPivotCaches() const {
  emscripten::val arr = emscripten::val::array();
  arr.set("status", status_from_rc(js_collect_list(handle_, 0U, &count_pivot_caches, &pivot_cache_item, arr)));
  return arr;
}

JsAddStyleResult JsWorkbook::pivotCacheCreate(uint32_t requestedId) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_add_pivot_cache(handle_, requestedId, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotCacheRemove(uint32_t cacheId) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_remove_pivot_cache(handle_, cacheId) : kBindingInvalidHandle);
}

// The typed getters here emit their whole declared payload on every exit
// path; only `status.ok` and the values differ. See the same note in
// `parts/workbook_styles.cpp`.

emscripten::val JsWorkbook::pivotCacheGetWorksheetSource(uint32_t cacheId) const {
  emscripten::val o = emscripten::val::object();
  int32_t present = 0;
  const char* ref = nullptr;
  const char* sheet = nullptr;
  const char* name = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_pivot_cache_worksheet_source(handle_, cacheId, &present, &ref, &sheet, &name)
                         : kBindingInvalidHandle;
  if (rc != 0) {
    present = 0;
    ref = nullptr;
    sheet = nullptr;
    name = nullptr;
  }
  o.set("status", status_from_rc(rc));
  o.set("present", present != 0);
  js_set_cstr(o, "ref", ref);
  js_set_cstr(o, "sheet", sheet);
  js_set_cstr(o, "name", name);
  return o;
}

JsStatus JsWorkbook::pivotCacheSetWorksheetSource(uint32_t cacheId, emscripten::val source) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  std::string ref;
  std::string sheet;
  std::string name;
  JsNarrowNumericReader reader("setPivotCacheWorksheetSource");
  const bool present = reader.boolean(source, "present", true, "pivotCacheSource.present");
  const char* ref_ptr = reader.optional_string(source, "ref", ref, "pivotCacheSource.ref");
  const char* sheet_ptr = reader.optional_string(source, "sheet", sheet, "pivotCacheSource.sheet");
  const char* name_ptr = reader.optional_string(source, "name", name, "pivotCacheSource.name");
  if (!reader.ok()) {
    return reader.failure();
  }
  fm_status_t rc =
      fm_workbook_set_pivot_cache_worksheet_source(handle_, cacheId, present ? 1 : 0, ref_ptr, sheet_ptr, name_ptr);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotCacheFieldCount(uint32_t cacheId) const {
  std::size_t count = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_pivot_cache_field_count(handle_, cacheId, &count) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsStringResult JsWorkbook::pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIdx) const {
  const char* name = nullptr;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_pivot_cache_field_name(handle_, cacheId, fieldIdx, &name)
                                            : kBindingInvalidHandle;
  return string_result(rc, name);
}

JsAddStyleResult JsWorkbook::pivotCacheFieldAdd(uint32_t cacheId, const std::string& name) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_add_pivot_cache_field(handle_, cacheId, name.c_str(), &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotCacheFieldClear(uint32_t cacheId) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_clear_pivot_cache_fields(handle_, cacheId)
                                           : kBindingInvalidHandle);
}

JsNumberResult JsWorkbook::pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIdx) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_field_shared_item_count(handle_, cacheId, fieldIdx, &count)
                             : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsCellResult JsWorkbook::pivotCacheFieldSharedItemValue(uint32_t cacheId, uint32_t fieldIdx, uint32_t itemIdx) const {
  JsCellResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  fm_value_t v{};
  const fm_status_t rc = fm_workbook_get_pivot_cache_field_shared_item_value(handle_, cacheId, fieldIdx, itemIdx, &v);
  adopt_value(rc, v, r.status, r.value);
  return r;
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIdx, double value) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_add_pivot_cache_field_shared_item_number(handle_, cacheId, fieldIdx, value)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemDate(uint32_t cacheId, uint32_t fieldIdx, double serial) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_add_pivot_cache_field_shared_item_date(handle_, cacheId, fieldIdx, serial)
                            : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotCacheFieldSharedItemIsDate(uint32_t cacheId, uint32_t fieldIdx,
                                                            uint32_t itemIdx) const {
  int32_t isDate = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_is_pivot_cache_field_shared_item_date(
                                                  handle_, cacheId, fieldIdx, itemIdx, &isDate)
                                            : kBindingInvalidHandle;
  return bool_result(rc, isDate);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIdx, const std::string& utf8) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_add_pivot_cache_field_shared_item_text(handle_, cacheId, fieldIdx, utf8.c_str())
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIdx, bool value) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_add_pivot_cache_field_shared_item_bool(handle_, cacheId, fieldIdx, value ? 1 : 0)
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_add_pivot_cache_field_shared_item_blank(handle_, cacheId, fieldIdx)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIdx, int32_t errorCode) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_workbook_add_pivot_cache_field_shared_item_error(handle_, cacheId, fieldIdx,
                                                                       static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_clear_pivot_cache_field_shared_items(handle_, cacheId, fieldIdx)
                            : kBindingInvalidHandle);
}

JsNumberResult JsWorkbook::pivotCacheRecordCount(uint32_t cacheId) const {
  std::size_t count = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_pivot_cache_record_count(handle_, cacheId, &count) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotCacheRecordAdd(uint32_t cacheId) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_add_pivot_cache_record(handle_, cacheId, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotCacheRecordClear(uint32_t cacheId) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_clear_pivot_cache_records(handle_, cacheId)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetNumber(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, double value) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_set_pivot_cache_record_number(handle_, cacheId, recordIdx, fieldIdx, value)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetDate(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, double serial) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_set_pivot_cache_record_date(handle_, cacheId, recordIdx, fieldIdx, serial)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetText(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx,
                                             const std::string& utf8) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_set_pivot_cache_record_text(handle_, cacheId, recordIdx, fieldIdx, utf8.c_str())
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetBool(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, bool value) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_set_pivot_cache_record_bool(handle_, cacheId, recordIdx, fieldIdx, value ? 1 : 0)
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetBlank(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_set_pivot_cache_record_blank(handle_, cacheId, recordIdx, fieldIdx)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetError(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx,
                                              int32_t errorCode) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_workbook_set_pivot_cache_record_error(handle_, cacheId, recordIdx, fieldIdx,
                                                            static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
}

// ---- PivotTable --------------------------------------------------------

JsAddStyleResult JsWorkbook::pivotCreate(uint32_t sheet, const std::string& name, uint32_t cacheId, uint32_t anchorRow,
                                         uint32_t anchorCol) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot(handle_, sheet, name.c_str(), cacheId, anchorRow, anchorCol, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotRemove(uint32_t sheet, uint32_t pivotIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_remove_pivot(handle_, sheet, pivotIdx) : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotSetName(uint32_t sheet, uint32_t pivotIdx, const std::string& name) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_set_pivot_name(handle_, sheet, pivotIdx, name.c_str())
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotSetAnchor(uint32_t sheet, uint32_t pivotIdx, emscripten::val rect) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setPivotAnchor");
  const fm_cell_rect_t r = js_pull_range(rect, &reader);
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_sheet_set_pivot_anchor(handle_, sheet, pivotIdx, &r));
}

JsStatus JsWorkbook::pivotSetGrandTotals(uint32_t sheet, uint32_t pivotIdx, bool rowsEnabled, bool colsEnabled) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_sheet_set_pivot_grand_totals(handle_, sheet, pivotIdx, rowsEnabled ? 1 : 0, colsEnabled ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotSetStyleInfo(uint32_t sheet, uint32_t pivotIdx, emscripten::val info) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  std::string name;
  JsNarrowNumericReader reader("setPivotStyleInfo");
  const char* name_ptr = reader.optional_string(info, "name", name, "info.name");
  const bool row_headers = reader.boolean(info, "showRowHeaders", false, "info.showRowHeaders");
  const bool col_headers = reader.boolean(info, "showColHeaders", false, "info.showColHeaders");
  const bool row_stripes = reader.boolean(info, "showRowStripes", false, "info.showRowStripes");
  const bool col_stripes = reader.boolean(info, "showColStripes", false, "info.showColStripes");
  const bool last_column = reader.boolean(info, "showLastColumn", false, "info.showLastColumn");
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_sheet_set_pivot_style_info(handle_, sheet, pivotIdx, name_ptr, row_headers ? 1 : 0,
                                                      col_headers ? 1 : 0, row_stripes ? 1 : 0, col_stripes ? 1 : 0,
                                                      last_column ? 1 : 0));
}

JsNumberResult JsWorkbook::pivotGetLayout(uint32_t sheet, uint32_t pivotIdx) const {
  std::int32_t layout = FM_PIVOT_REPORT_LAYOUT_COMPACT;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_report_layout(handle_, sheet, pivotIdx, &layout) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(layout));
}

JsStatus JsWorkbook::pivotSetLayout(uint32_t sheet, uint32_t pivotIdx, uint32_t layout) {
  return status_from_rc(
      handle_ != nullptr ? fm_sheet_set_pivot_report_layout(handle_, sheet, pivotIdx, static_cast<std::int32_t>(layout))
                         : kBindingInvalidHandle);
}

JsAddStyleResult JsWorkbook::pivotCacheId(uint32_t sheet, uint32_t pivotIdx) const {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_sheet_get_pivot_cache_id(handle_, sheet, pivotIdx, &out);
  return index_result(rc, out);
}

JsStringResult JsWorkbook::pivotName(uint32_t sheet, uint32_t pivotIdx) const {
  const char* name = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_name(handle_, sheet, pivotIdx, &name) : kBindingInvalidHandle;
  return string_result(rc, rc == 0 ? name : nullptr);
}

emscripten::val JsWorkbook::pivotGetGrandTotals(uint32_t sheet, uint32_t pivotIdx) const {
  emscripten::val o = emscripten::val::object();
  int32_t rows = 0;
  int32_t cols = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_grand_totals(handle_, sheet, pivotIdx, &rows, &cols)
                                            : kBindingInvalidHandle;
  o.set("status", status_from_rc(rc));
  o.set("rows", rc == 0 && rows != 0);
  o.set("cols", rc == 0 && cols != 0);
  return o;
}

emscripten::val JsWorkbook::pivotGetStyleInfo(uint32_t sheet, uint32_t pivotIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_style_info_t info{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_style_info(handle_, sheet, pivotIdx, &info) : kBindingInvalidHandle;
  if (rc != 0) {
    info = fm_pivot_style_info_t{};
  }
  o.set("status", status_from_rc(rc));
  o.set("present", info.present != 0);
  js_set_cstr(o, "name", info.name);
  o.set("showRowHeaders", info.show_row_headers != 0);
  o.set("showColHeaders", info.show_col_headers != 0);
  o.set("showRowStripes", info.show_row_stripes != 0);
  o.set("showColStripes", info.show_col_stripes != 0);
  o.set("showLastColumn", info.show_last_column != 0);
  return o;
}

JsNumberResult JsWorkbook::pivotFieldCount(uint32_t sheet, uint32_t pivotIdx) const {
  std::size_t count = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_field_count(handle_, sheet, pivotIdx, &count) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  std::string source_name;
  std::string custom_name;
  std::string number_format;
  JsNarrowNumericReader reader("addPivotField");
  // Every nested read uses this per-call reader so getter/proxy failures and
  // lossy integer conversions are reported before the C ABI is called.
  const char* source_name_ptr = reader.optional_string(spec, "sourceName", source_name, "pivotField.sourceName");
  const char* custom_name_ptr = reader.optional_string(spec, "customName", custom_name, "pivotField.customName");
  const bool subtotal_top = reader.boolean(spec, "subtotalTop", true, "pivotField.subtotalTop");
  const char* number_format_ptr =
      reader.optional_string(spec, "numberFormat", number_format, "pivotField.numberFormat");

  fm_pivot_field_spec_t c_spec{};
  // `sourceName` is required; an omitted key passes NULL through rather
  // than `js_pull_string`'s empty-string default, which the C ABI's own
  // required-field check does not treat the same as a missing argument.
  c_spec.source_name = source_name_ptr;
  c_spec.custom_name = custom_name_ptr;
  c_spec.axis = static_cast<fm_pivot_axis_t>(reader.u32(spec, "axis", 0U, "pivotField.axis"));
  c_spec.subtotal_top = subtotal_top ? 1 : 0;
  c_spec.number_format = number_format_ptr;

  if (!reader.ok()) {
    r.status = reader.failure();
    return r;
  }

  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot_field(handle_, sheet, pivotIdx, &c_spec, &out);
  return index_result(rc, out);
}

emscripten::val JsWorkbook::pivotFieldAt(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_field_info_t info{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_field(handle_, sheet, pivotIdx, fieldIdx, &info) : kBindingInvalidHandle;
  if (rc != 0) {
    info = fm_pivot_field_info_t{};
    info.axis = -1;
    info.cache_field_index = -1;
    info.group_base_index = -1;
  }
  o.set("status", status_from_rc(rc));
  js_set_cstr(o, "sourceName", info.source_name);
  js_set_cstr(o, "customName", info.custom_name);
  o.set("axis", info.axis);
  o.set("subtotalTop", info.subtotal_top != 0);
  js_set_cstr(o, "numberFormat", info.number_format);
  o.set("cacheFieldIndex", info.cache_field_index);
  o.set("sortOrder", info.sort_order);
  js_set_cstr(o, "sortByField", info.sort_by_field);
  o.set("datePeriods", info.date_periods);
  o.set("dateCalendar", info.date_calendar);
  o.set("groupBaseIndex", info.group_base_index);
  o.set("defaultSubtotal", info.default_subtotal != 0);
  o.set("insertBlankRow", info.insert_blank_row != 0);
  o.set("fillDownLabels", info.fill_down_labels != 0);
  o.set("showAll", info.show_all != 0);
  emscripten::val fns = emscripten::val::array();
  for (std::uint32_t agg = 0; agg < 32U; ++agg) {
    if ((info.subtotal_fns & (1U << agg)) != 0U) {
      fns.call<void>("push", agg);
    }
  }
  o.set("subtotalFns", fns);
  return o;
}

JsNumberResult JsWorkbook::pivotFieldItemCount(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_field_item_count(handle_, sheet, pivotIdx, fieldIdx, &count)
                             : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

emscripten::val JsWorkbook::pivotFieldItemAt(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                             uint32_t itemIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_item_info_t info{};
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_field_item(handle_, sheet, pivotIdx, fieldIdx, itemIdx, &info)
                             : kBindingInvalidHandle;
  if (rc != 0) {
    info = fm_pivot_item_info_t{};
    info.cache_index = -1;
  }
  o.set("status", status_from_rc(rc));
  js_set_cstr(o, "name", info.name);
  o.set("cacheIndex", info.cache_index);
  o.set("visible", info.visible != 0);
  if (info.caption != nullptr) {
    js_set_cstr(o, "caption", info.caption);
  } else {
    o.set("caption", emscripten::val::null());
  }
  o.set("showDetail", info.show_detail != 0);
  return o;
}

JsStatus JsWorkbook::pivotFieldClear(uint32_t sheet, uint32_t pivotIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_pivot_fields(handle_, sheet, pivotIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetAxis(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, int32_t axis) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_sheet_set_pivot_field_axis(handle_, sheet, pivotIdx, fieldIdx, static_cast<std::int32_t>(axis));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetSort(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t order,
                                       const std::string& byField) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  const char* by = byField.empty() ? nullptr : byField.c_str();
  fm_status_t rc =
      fm_sheet_set_pivot_field_sort(handle_, sheet, pivotIdx, fieldIdx, static_cast<std::int32_t>(order), by);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetSubtotalTop(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool top) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_subtotal_top(handle_, sheet, pivotIdx, fieldIdx, top ? 1 : 0)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetDefaultSubtotal(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool on) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_default_subtotal(handle_, sheet, pivotIdx, fieldIdx, on ? 1 : 0)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetInsertBlankRow(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool on) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_insert_blank_row(handle_, sheet, pivotIdx, fieldIdx, on ? 1 : 0)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetFillDownLabels(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool on) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_fill_down_labels(handle_, sheet, pivotIdx, fieldIdx, on ? 1 : 0)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetShowDetail(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool on) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_show_detail(handle_, sheet, pivotIdx, fieldIdx, on ? 1 : 0)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetItemShowDetail(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t itemIdx,
                                                 bool on) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_set_pivot_field_item_show_detail(handle_, sheet, pivotIdx,
                                                                                       fieldIdx, itemIdx, on ? 1 : 0)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetItemCaption(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t itemIdx,
                                              const std::string& caption) {
  const char* text = caption.empty() ? nullptr : caption.c_str();
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_item_caption(handle_, sheet, pivotIdx, fieldIdx, itemIdx, text)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetCustomName(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                             const std::string& name) {
  const char* text = name.empty() ? nullptr : name.c_str();
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_custom_name(handle_, sheet, pivotIdx, fieldIdx, text)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetShowAll(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool on) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_show_all(handle_, sheet, pivotIdx, fieldIdx, on ? 1 : 0)
                            : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotGetOptions(uint32_t sheet, uint32_t pivotIdx) const {
  fm_pivot_options_t opt{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_options(handle_, sheet, pivotIdx, &opt) : kBindingInvalidHandle;
  if (rc != 0) {
    opt = fm_pivot_options_t{};
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("showError", opt.show_error != 0);
  js_set_cstr(o, "errorCaption", opt.error_caption);
  o.set("showMissing", opt.show_missing != 0);
  js_set_cstr(o, "missingCaption", opt.missing_caption);
  o.set("showHeaders", opt.show_headers != 0);
  o.set("showDrill", opt.show_drill != 0);
  o.set("showDataTips", opt.show_data_tips != 0);
  o.set("gridDropZones", opt.grid_drop_zones != 0);
  o.set("showEmptyRow", opt.show_empty_row != 0);
  o.set("showEmptyCol", opt.show_empty_col != 0);
  o.set("showDataDropDown", opt.show_data_drop_down != 0);
  o.set("printDrill", opt.print_drill != 0);
  o.set("fieldPrintTitles", opt.field_print_titles != 0);
  o.set("itemPrintTitles", opt.item_print_titles != 0);
  o.set("mergeItem", opt.merge_item != 0);
  o.set("indent", opt.indent);
  o.set("pageOverThenDown", opt.page_over_then_down != 0);
  o.set("pageWrap", opt.page_wrap);
  o.set("multipleFieldFilters", opt.multiple_field_filters != 0);
  o.set("preserveFormatting", opt.preserve_formatting != 0);
  o.set("customListSort", opt.custom_list_sort != 0);
  js_set_cstr(o, "altText", opt.alt_text);
  js_set_cstr(o, "altTextSummary", opt.alt_text_summary);
  js_set_cstr(o, "grandTotalCaption", opt.grand_total_caption);
  o.set("autoFormatting", opt.auto_formatting != 0);
  o.set("enableDrill", opt.enable_drill != 0);
  return o;
}

JsStatus JsWorkbook::pivotSetOptions(uint32_t sheet, uint32_t pivotIdx, emscripten::val options) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  if (!js_value_present(options)) {
    return binding_error_status(kInvalidArgument, "setPivotOptions: `options` must be an object");
  }
  // The C call writes every member, so members the caller omits keep the
  // pivot's current value.
  fm_pivot_options_t opt{};
  const fm_status_t get_rc = fm_sheet_get_pivot_options(handle_, sheet, pivotIdx, &opt);
  if (get_rc != 0) {
    return status_from_rc(get_rc);
  }
  std::string error_caption = string_from_cstr(opt.error_caption);
  std::string missing_caption = string_from_cstr(opt.missing_caption);
  std::string alt_text = string_from_cstr(opt.alt_text);
  std::string alt_summary = string_from_cstr(opt.alt_text_summary);
  std::string grand_total = string_from_cstr(opt.grand_total_caption);

  JsNarrowNumericReader reader("setPivotOptions");
  const auto flag = [&](const char* key, int32_t current) {
    const std::string field = std::string("pivotOptions.") + key;
    return reader.boolean(options, key, current != 0, field.c_str()) ? 1 : 0;
  };
  const auto text = [&](const char* key, std::string& storage) {
    std::string value;
    const std::string field = std::string("pivotOptions.") + key;
    if (reader.optional_string(options, key, value, field.c_str()) != nullptr) {
      storage = value;
    }
  };
  opt.show_error = flag("showError", opt.show_error);
  text("errorCaption", error_caption);
  opt.show_missing = flag("showMissing", opt.show_missing);
  text("missingCaption", missing_caption);
  opt.show_headers = flag("showHeaders", opt.show_headers);
  opt.show_drill = flag("showDrill", opt.show_drill);
  opt.show_data_tips = flag("showDataTips", opt.show_data_tips);
  opt.grid_drop_zones = flag("gridDropZones", opt.grid_drop_zones);
  opt.show_empty_row = flag("showEmptyRow", opt.show_empty_row);
  opt.show_empty_col = flag("showEmptyCol", opt.show_empty_col);
  opt.show_data_drop_down = flag("showDataDropDown", opt.show_data_drop_down);
  opt.print_drill = flag("printDrill", opt.print_drill);
  opt.field_print_titles = flag("fieldPrintTitles", opt.field_print_titles);
  opt.item_print_titles = flag("itemPrintTitles", opt.item_print_titles);
  opt.merge_item = flag("mergeItem", opt.merge_item);
  opt.indent = reader.u32(options, "indent", opt.indent, "pivotOptions.indent");
  opt.page_over_then_down = flag("pageOverThenDown", opt.page_over_then_down);
  opt.page_wrap = reader.u32(options, "pageWrap", opt.page_wrap, "pivotOptions.pageWrap");
  opt.multiple_field_filters = flag("multipleFieldFilters", opt.multiple_field_filters);
  opt.preserve_formatting = flag("preserveFormatting", opt.preserve_formatting);
  opt.custom_list_sort = flag("customListSort", opt.custom_list_sort);
  text("altText", alt_text);
  text("altTextSummary", alt_summary);
  text("grandTotalCaption", grand_total);
  opt.auto_formatting = flag("autoFormatting", opt.auto_formatting);
  opt.enable_drill = flag("enableDrill", opt.enable_drill);
  if (!reader.ok()) {
    return reader.failure();
  }
  const auto view = [](const std::string& s) { return s.empty() ? nullptr : s.c_str(); };
  opt.error_caption = view(error_caption);
  opt.missing_caption = view(missing_caption);
  opt.alt_text = view(alt_text);
  opt.alt_text_summary = view(alt_summary);
  opt.grand_total_caption = view(grand_total);
  return status_from_rc(fm_sheet_set_pivot_options(handle_, sheet, pivotIdx, &opt));
}

emscripten::val JsWorkbook::pivotCacheGetOptions(uint32_t cacheId) const {
  fm_pivot_cache_options_t opt{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_pivot_cache_options(handle_, cacheId, &opt) : kBindingInvalidHandle;
  if (rc != 0) {
    opt = fm_pivot_cache_options_t{};
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("saveData", opt.save_data != 0);
  o.set("missingItemsLimit", opt.missing_items_limit);
  o.set("refreshOnLoad", opt.refresh_on_load != 0);
  return o;
}

JsStatus JsWorkbook::pivotCacheSetOptions(uint32_t cacheId, emscripten::val options) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  if (!js_value_present(options)) {
    return binding_error_status(kInvalidArgument, "setPivotCacheOptions: `options` must be an object");
  }
  fm_pivot_cache_options_t opt{};
  const fm_status_t get_rc = fm_workbook_get_pivot_cache_options(handle_, cacheId, &opt);
  if (get_rc != 0) {
    return status_from_rc(get_rc);
  }
  JsNarrowNumericReader reader("setPivotCacheOptions");
  opt.save_data = reader.boolean(options, "saveData", opt.save_data != 0, "pivotCacheOptions.saveData") ? 1 : 0;
  opt.missing_items_limit =
      reader.i32(options, "missingItemsLimit", opt.missing_items_limit, "pivotCacheOptions.missingItemsLimit");
  opt.refresh_on_load =
      reader.boolean(options, "refreshOnLoad", opt.refresh_on_load != 0, "pivotCacheOptions.refreshOnLoad") ? 1 : 0;
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_workbook_set_pivot_cache_options(handle_, cacheId, &opt));
}

JsAddStyleResult JsWorkbook::pivotCacheCalculatedFieldAdd(uint32_t cacheId, const std::string& name,
                                                          const std::string& formula) {
  std::size_t index = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_add_pivot_cache_calculated_field(
                                                  handle_, cacheId, name.c_str(), formula.c_str(), &index)
                                            : kBindingInvalidHandle;
  return index_result(rc, index);
}

JsStatus JsWorkbook::pivotCacheCalculatedFieldSetFormula(uint32_t cacheId, uint32_t fieldIdx,
                                                         const std::string& formula) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_set_pivot_cache_calculated_field_formula(
                                                 handle_, cacheId, fieldIdx, formula.c_str())
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheCalculatedFieldRemove(uint32_t cacheId, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_remove_pivot_cache_calculated_field(handle_, cacheId, fieldIdx)
                                           : kBindingInvalidHandle);
}

JsStringResult JsWorkbook::pivotCacheFieldFormula(uint32_t cacheId, uint32_t fieldIdx) const {
  const char* formula = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_field_formula(handle_, cacheId, fieldIdx, &formula)
                             : kBindingInvalidHandle;
  return string_result(rc, rc == 0 ? formula : nullptr);
}

JsAddStyleResult JsWorkbook::pivotCacheCalculatedItemAdd(uint32_t cacheId, uint32_t fieldIdx, const std::string& name,
                                                         const std::string& formula) {
  std::size_t out = 0;
  return index_result(handle_ != nullptr ? fm_workbook_add_pivot_cache_calculated_item(
                                               handle_, cacheId, fieldIdx, name.c_str(), formula.c_str(), &out)
                                         : kBindingInvalidHandle,
                      out);
}

JsNumberResult JsWorkbook::pivotCacheCalculatedItemCount(uint32_t cacheId) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_calculated_item_count(handle_, cacheId, &count)
                             : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

emscripten::val JsWorkbook::pivotCacheCalculatedItemAt(uint32_t cacheId, uint32_t idx) const {
  fm_pivot_calculated_item_t item{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_pivot_cache_calculated_item(handle_, cacheId, idx, &item)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    item = fm_pivot_calculated_item_t{};
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("fieldIndex", item.field_index);
  js_set_cstr(o, "name", item.name);
  js_set_cstr(o, "formula", item.formula);
  o.set("scoped", item.scoped != 0);
  return o;
}

JsStatus JsWorkbook::pivotCacheCalculatedItemSetFormula(uint32_t cacheId, uint32_t idx, const std::string& formula) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_set_pivot_cache_calculated_item_formula(handle_, cacheId, idx, formula.c_str())
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheCalculatedItemMove(uint32_t cacheId, uint32_t idx, uint32_t newIdx) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_move_pivot_cache_calculated_item(handle_, cacheId, idx, newIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheCalculatedItemRemove(uint32_t cacheId, uint32_t idx) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_remove_pivot_cache_calculated_item(handle_, cacheId, idx)
                                           : kBindingInvalidHandle);
}

JsStringResult JsWorkbook::pivotDataFieldDefaultName(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                                     int32_t aggregation, std::size_t excludeDataField) const {
  const char* name = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_data_field_default_name(handle_, sheet, pivotIdx, fieldIdx,
                                                                          aggregation, excludeDataField, &name)
                             : kBindingInvalidHandle;
  return string_result(rc, rc == 0 ? name : nullptr);
}

JsStatus JsWorkbook::pivotProject(uint32_t sheet, uint32_t pivotIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_project_pivot(handle_, sheet, pivotIdx) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotProjectionPending(uint32_t sheet, uint32_t pivotIdx) const {
  int32_t pending = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_is_pivot_projection_pending(handle_, sheet, pivotIdx, &pending)
                                            : kBindingInvalidHandle;
  return bool_result(rc, pending);
}

JsAddStyleResult JsWorkbook::pivotShowDetails(uint32_t sheet, uint32_t pivotIdx, uint32_t row, uint32_t col) {
  std::size_t newSheet = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_show_pivot_details(handle_, sheet, pivotIdx, row, col, &newSheet)
                                            : kBindingInvalidHandle;
  return index_result(rc, newSheet);
}

JsAddStyleResult JsWorkbook::pivotFieldAddItem(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                               const std::string& name, bool visible) {
  std::size_t out = 0;
  return index_result(handle_ != nullptr ? fm_sheet_add_pivot_field_item(handle_, sheet, pivotIdx, fieldIdx,
                                                                         name.c_str(), visible ? 1 : 0, &out)
                                         : kBindingInvalidHandle,
                      out);
}

JsAddStyleResult JsWorkbook::pivotFieldAddItemAt(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                                 uint32_t cacheIndex, bool visible) {
  std::size_t out = 0;
  return index_result(handle_ != nullptr ? fm_sheet_add_pivot_field_shared_item(handle_, sheet, pivotIdx, fieldIdx,
                                                                                cacheIndex, visible ? 1 : 0, &out)
                                         : kBindingInvalidHandle,
                      out);
}

JsStatus JsWorkbook::pivotFieldClearItems(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_pivot_field_items(handle_, sheet, pivotIdx, fieldIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetItemVisible(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t itemIdx,
                                              bool visible) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_sheet_set_pivot_field_item_visible(handle_, sheet, pivotIdx, fieldIdx, itemIdx, visible ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldAddSubtotalFn(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t agg) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc =
      fm_sheet_add_pivot_field_subtotal_fn(handle_, sheet, pivotIdx, fieldIdx, static_cast<std::int32_t>(agg));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldClearSubtotalFns(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_pivot_field_subtotal_fns(handle_, sheet, pivotIdx, fieldIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetDateGroup(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t periods,
                                            uint32_t calendar, uint32_t intervalDays, double start, double end) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_sheet_set_pivot_field_date_group(handle_, sheet, pivotIdx, fieldIdx, periods,
                                                       static_cast<std::int32_t>(calendar), intervalDays, start, end);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldClearDateGroup(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_pivot_field_date_group(handle_, sheet, pivotIdx, fieldIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFieldSetNumberFormat(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                               const std::string& utf8) {
  return status_from_rc(handle_ != nullptr
                            ? fm_sheet_set_pivot_field_number_format(handle_, sheet, pivotIdx, fieldIdx, utf8.c_str())
                            : kBindingInvalidHandle);
}

namespace {

// The JS spelling of `FM_PIVOT_VALUES_FIELD` (`PIVOT_VALUES_FIELD`).
constexpr int32_t kPivotValuesFieldJs = -2;

// Shared body of `pivotSetRowFieldOrder` / `pivotSetColFieldOrder`.
using PivotFieldOrderFn = fm_status_t (*)(fm_workbook_t*, size_t, size_t, const uint32_t*, size_t);

JsStatus set_pivot_field_order(fm_workbook_t* wb, uint32_t sheet, uint32_t pivotIdx, const emscripten::val& indices,
                               PivotFieldOrderFn fn) {
  if (wb == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("pivotFieldOrder");
  if (!reader.is_array(indices, "pivotFieldOrder")) {
    return binding_error_status(kInvalidArgument, "pivot field order must be an array");
  }
  const uint32_t length = reader.length(indices, "pivotFieldOrder");
  if (!reader.ok()) {
    return reader.failure();
  }
  std::vector<uint32_t> v;
  v.reserve(length);
  for (uint32_t i = 0; i < length && reader.ok(); ++i) {
    // Read the raw element once through the guarded property reader, then
    // let the checked conversion reject a symbol, fraction, non-finite value,
    // or uint32 overflow. No subsequent JS reads occur after failure.
    const std::string key = std::to_string(i);
    const emscripten::val value = reader.value(indices, key.c_str(), "pivotFieldOrder.indices");
    // -2 is the Values level; the checked i32 read still rejects other negatives.
    const int32_t index = reader.i32_value(value, 0, "pivotFieldOrder.indices");
    if (index == kPivotValuesFieldJs) {
      v.push_back(FM_PIVOT_VALUES_FIELD);
    } else if (index < 0) {
      return binding_error_status(kInvalidArgument, "pivot field order index must be >= 0 or PIVOT_VALUES_FIELD");
    } else {
      v.push_back(static_cast<uint32_t>(index));
    }
  }
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fn(wb, sheet, pivotIdx, v.empty() ? nullptr : v.data(), v.size()));
}

}  // namespace

JsStatus JsWorkbook::pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices) {
  return set_pivot_field_order(handle_, sheet, pivotIdx, indices, &fm_sheet_set_pivot_row_field_order);
}

JsStatus JsWorkbook::pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices) {
  return set_pivot_field_order(handle_, sheet, pivotIdx, indices, &fm_sheet_set_pivot_col_field_order);
}

namespace {

// Shared body of `pivotRowFieldOrder` / `pivotColFieldOrder`.
using PivotFieldOrderGetter = fm_status_t (*)(const fm_workbook_t*, size_t, size_t, const uint32_t**, size_t*);

emscripten::val get_pivot_field_order(const fm_workbook_t* wb, uint32_t sheet, uint32_t pivotIdx,
                                      PivotFieldOrderGetter fn) {
  emscripten::val o = emscripten::val::object();
  emscripten::val indices = emscripten::val::array();
  const uint32_t* data = nullptr;
  std::size_t count = 0;
  const fm_status_t rc = wb != nullptr ? fn(wb, sheet, pivotIdx, &data, &count) : kBindingInvalidHandle;
  for (std::size_t i = 0; rc == 0 && i < count; ++i) {
    indices.call<void>("push", data[i] == FM_PIVOT_VALUES_FIELD ? kPivotValuesFieldJs : static_cast<int32_t>(data[i]));
  }
  o.set("status", status_from_rc(rc));
  o.set("indices", indices);
  return o;
}

}  // namespace

emscripten::val JsWorkbook::pivotRowFieldOrder(uint32_t sheet, uint32_t pivotIdx) const {
  return get_pivot_field_order(handle_, sheet, pivotIdx, &fm_sheet_get_pivot_row_field_order);
}

emscripten::val JsWorkbook::pivotColFieldOrder(uint32_t sheet, uint32_t pivotIdx) const {
  return get_pivot_field_order(handle_, sheet, pivotIdx, &fm_sheet_get_pivot_col_field_order);
}

JsNumberResult JsWorkbook::pivotDataFieldCount(uint32_t sheet, uint32_t pivotIdx) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_data_field_count(handle_, sheet, pivotIdx, &count)
                                            : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotDataFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  JsNarrowNumericReader reader("addPivotDataField");
  if (!build_data_field_spec_checked(reader, spec, c_spec, name_buf, nfmt_buf)) {
    r.status = reader.failure();
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot_data_field(handle_, sheet, pivotIdx, &c_spec, &out);
  return index_result(rc, out);
}

emscripten::val JsWorkbook::pivotDataFieldAt(uint32_t sheet, uint32_t pivotIdx, uint32_t dataFieldIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_data_field_spec_t spec{};
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_data_field(handle_, sheet, pivotIdx, dataFieldIdx, &spec)
                             : kBindingInvalidHandle;
  if (rc != 0) {
    spec = fm_pivot_data_field_spec_t{};
    spec.show_as_base_field = -1;
    spec.show_as_base_item = -1;
  }
  o.set("status", status_from_rc(rc));
  js_set_cstr(o, "name", spec.name);
  o.set("fieldIndex", spec.field_index);
  o.set("aggregation", static_cast<int32_t>(spec.aggregation));
  js_set_cstr(o, "numberFormat", spec.number_format);
  o.set("showAs", static_cast<int32_t>(spec.show_as));
  o.set("showAsBaseField", spec.show_as_base_field);
  o.set("showAsBaseItem", spec.show_as_base_item);
  return o;
}

JsStatus JsWorkbook::pivotDataFieldClear(uint32_t sheet, uint32_t pivotIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_pivot_data_fields(handle_, sheet, pivotIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotDataFieldSet(uint32_t sheet, uint32_t pivotIdx, uint32_t dataFieldIdx, emscripten::val spec) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  JsNarrowNumericReader reader("setPivotDataField");
  if (!build_data_field_spec_checked(reader, spec, c_spec, name_buf, nfmt_buf)) {
    return reader.failure();
  }
  fm_status_t rc = fm_sheet_set_pivot_data_field(handle_, sheet, pivotIdx, dataFieldIdx, &c_spec);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotFilterCount(uint32_t sheet, uint32_t pivotIdx) const {
  std::size_t count = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_filter_count(handle_, sheet, pivotIdx, &count) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotFilterAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  if (handle_ == nullptr) {
    return {invalid_handle_status()};
  }
  // `fieldName` is required; an omitted key passes NULL through rather
  // than `js_pull_string`'s empty-string default, which the C ABI's own
  // required-field check does not treat the same as a missing argument.
  std::string field_name;
  std::string value_text;
  JsNarrowNumericReader reader("addPivotFilter");
  const char* field_name_ptr = reader.optional_string(spec, "fieldName", field_name, "pivotFilter.fieldName");
  const double value_double = reader.number(spec, "valueDouble", 0.0, "pivotFilter.valueDouble");
  const char* value_text_ptr = reader.optional_string(spec, "valueText", value_text, "pivotFilter.valueText");
  const double value_high_double = reader.number(spec, "valueHighDouble", 0.0, "pivotFilter.valueHighDouble");
  std::string value_high_text;
  const char* value_high_text_ptr =
      reader.optional_string(spec, "valueHighText", value_high_text, "pivotFilter.valueHighText");

  fm_pivot_filter_spec_t c_spec{};
  c_spec.axis = static_cast<fm_pivot_axis_t>(reader.u32(spec, "axis", 0U, "pivotFilter.axis"));
  c_spec.field_name = field_name_ptr;
  c_spec.type = static_cast<fm_pivot_filter_type_t>(reader.u32(spec, "type", 0U, "pivotFilter.type"));
  // valueKind defaults to NONE (-1) when omitted; the C ABI rejects NONE
  // for non-range filters.
  c_spec.value_kind = static_cast<fm_pivot_filter_value_kind_t>(
      reader.i32(spec, "valueKind", FM_PIVOT_FILTER_VALUE_NONE, "pivotFilter.valueKind"));
  c_spec.value_int = reader.i32(spec, "valueInt", 0, "pivotFilter.valueInt");
  c_spec.value_double = value_double;
  c_spec.value_text = value_text_ptr;
  c_spec.value_high_kind = static_cast<fm_pivot_filter_value_kind_t>(
      reader.i32(spec, "valueHighKind", FM_PIVOT_FILTER_VALUE_NONE, "pivotFilter.valueHighKind"));
  c_spec.value_high_int = reader.i32(spec, "valueHighInt", 0, "pivotFilter.valueHighInt");
  c_spec.value_high_double = value_high_double;
  c_spec.value_high_text = value_high_text_ptr;
  c_spec.data_field_index = reader.u32(spec, "dataFieldIndex", 0U, "pivotFilter.dataFieldIndex");

  if (!reader.ok()) {
    return {reader.failure()};
  }

  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot_filter(handle_, sheet, pivotIdx, &c_spec, &out);
  return index_result(rc, out);
}

emscripten::val JsWorkbook::pivotFilterAt(uint32_t sheet, uint32_t pivotIdx, uint32_t filterIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_filter_spec_t spec{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_filter(handle_, sheet, pivotIdx, filterIdx, &spec)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    spec = fm_pivot_filter_spec_t{};
  }
  o.set("status", status_from_rc(rc));
  o.set("axis", static_cast<int32_t>(spec.axis));
  js_set_cstr(o, "fieldName", spec.field_name);
  o.set("type", static_cast<int32_t>(spec.type));
  o.set("dataFieldIndex", spec.data_field_index);
  o.set("valueKind", static_cast<int32_t>(spec.value_kind));
  o.set("valueInt", spec.value_int);
  o.set("valueDouble", spec.value_double);
  js_set_cstr(o, "valueText", spec.value_text);
  o.set("valueHighKind", static_cast<int32_t>(spec.value_high_kind));
  o.set("valueHighInt", spec.value_high_int);
  o.set("valueHighDouble", spec.value_high_double);
  js_set_cstr(o, "valueHighText", spec.value_high_text);
  return o;
}

JsStatus JsWorkbook::pivotFilterClear(uint32_t sheet, uint32_t pivotIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_pivot_filters(handle_, sheet, pivotIdx)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotFilterRemoveAt(uint32_t sheet, uint32_t pivotIdx, uint32_t filterIdx) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_remove_pivot_filter(handle_, sheet, pivotIdx, filterIdx)
                                           : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
