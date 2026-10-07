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
                                   fm_pivot_data_field_spec_t& out, std::string& name_buf, std::string& nfmt_buf,
                                   bool& has_nfmt) {
  // Read the optional strings before the numeric fields. Once the reader
  // rejects a numeric value, no further access to the caller's object is
  // allowed before returning the binding error.
  out.name = reader.optional_string(spec, "name", name_buf, "pivotDataField.name");
  out.number_format = reader.optional_string(spec, "numberFormat", nfmt_buf, "pivotDataField.numberFormat");
  has_nfmt = out.number_format != nullptr;
  if (!reader.ok()) {
    return false;
  }
  out.field_index = reader.u32(spec, "fieldIndex", 0U, "pivotDataField.fieldIndex");
  out.aggregation =
      static_cast<fm_pivot_aggregation_t>(reader.u32(spec, "aggregation", 0U, "pivotDataField.aggregation"));
  out.show_as = static_cast<fm_pivot_show_as_t>(reader.u32(spec, "showAs", 0U, "pivotDataField.showAs"));
  out.show_as_base_field = reader.i32(spec, "showAsBaseField", -1, "pivotDataField.showAsBaseField");
  out.show_as_base_item = reader.i32(spec, "showAsBaseItem", -1, "pivotDataField.showAsBaseItem");
  return reader.ok();
}

}  // namespace

// ---- PivotCache --------------------------------------------------------

JsNumberResult JsWorkbook::pivotCacheCount() const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_cache_count(handle_, &count);
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotCacheIdAt(uint32_t idx) const {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_pivot_cache_id_at(handle_, idx, &out);
  return index_result(rc, out);
}

JsAddStyleResult JsWorkbook::pivotCacheCreate(uint32_t requestedId) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_pivot_cache_create(handle_, requestedId, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotCacheRemove(uint32_t cacheId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_remove(handle_, cacheId);
  return status_from_rc(rc);
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
      handle_ != nullptr ? fm_workbook_pivot_cache_get_worksheet_source(handle_, cacheId, &present, &ref, &sheet, &name)
                         : 7000;
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
    return error_status(7000);
  }
  std::string ref;
  std::string sheet;
  std::string name;
  JsNarrowNumericReader reader("pivotCacheSetWorksheetSource");
  const bool present = reader.boolean(source, "present", true, "pivotCacheSource.present");
  const char* ref_ptr = reader.optional_string(source, "ref", ref, "pivotCacheSource.ref");
  const char* sheet_ptr = reader.optional_string(source, "sheet", sheet, "pivotCacheSource.sheet");
  const char* name_ptr = reader.optional_string(source, "name", name, "pivotCacheSource.name");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  fm_status_t rc =
      fm_workbook_pivot_cache_set_worksheet_source(handle_, cacheId, present ? 1 : 0, ref_ptr, sheet_ptr, name_ptr);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotCacheFieldCount(uint32_t cacheId) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_cache_field_count(handle_, cacheId, &count);
  return number_result(rc, static_cast<double>(count));
}

JsStringResult JsWorkbook::pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIdx) const {
  const char* name = nullptr;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_pivot_cache_field_name(handle_, cacheId, fieldIdx, &name)
                                            : kBindingInvalidHandle;
  return string_result(rc, name);
}

JsAddStyleResult JsWorkbook::pivotCacheFieldAdd(uint32_t cacheId, const std::string& name) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_field_add(handle_, cacheId, name.c_str(), &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotCacheFieldClear(uint32_t cacheId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_clear(handle_, cacheId);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIdx) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_cache_field_shared_item_count(handle_, cacheId, fieldIdx, &count);
  return number_result(rc, static_cast<double>(count));
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIdx, double value) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_number(handle_, cacheId, fieldIdx, value);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIdx, const std::string& utf8) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_text(handle_, cacheId, fieldIdx, utf8.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIdx, bool value) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_bool(handle_, cacheId, fieldIdx, value ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_blank(handle_, cacheId, fieldIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIdx, int32_t errorCode) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_error(handle_, cacheId, fieldIdx,
                                                                       static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_clear_shared_items(handle_, cacheId, fieldIdx);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotCacheRecordCount(uint32_t cacheId) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_cache_record_count(handle_, cacheId, &count);
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotCacheRecordAdd(uint32_t cacheId) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_record_add(handle_, cacheId, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotCacheRecordClear(uint32_t cacheId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_record_clear(handle_, cacheId);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheRecordSetNumber(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, double value) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_record_set_number(handle_, cacheId, recordIdx, fieldIdx, value);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheRecordSetText(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx,
                                             const std::string& utf8) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_record_set_text(handle_, cacheId, recordIdx, fieldIdx, utf8.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheRecordSetBool(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx, bool value) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_record_set_bool(handle_, cacheId, recordIdx, fieldIdx, value ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheRecordSetBlank(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_record_set_blank(handle_, cacheId, recordIdx, fieldIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheRecordSetError(uint32_t cacheId, uint32_t recordIdx, uint32_t fieldIdx,
                                              int32_t errorCode) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_record_set_error(handle_, cacheId, recordIdx, fieldIdx,
                                                            static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
}

// ---- PivotTable --------------------------------------------------------

JsAddStyleResult JsWorkbook::pivotCreate(uint32_t sheet, const std::string& name, uint32_t cacheId, uint32_t anchorRow,
                                         uint32_t anchorCol) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_create(handle_, sheet, name.c_str(), cacheId, anchorRow, anchorCol, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotRemove(uint32_t sheet, uint32_t pivotIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_remove(handle_, sheet, pivotIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotSetName(uint32_t sheet, uint32_t pivotIdx, const std::string& name) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_set_name(handle_, sheet, pivotIdx, name.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotSetAnchor(uint32_t sheet, uint32_t pivotIdx, uint32_t anchorRow, uint32_t anchorCol,
                                    uint32_t spanRows, uint32_t spanCols) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_set_anchor(handle_, sheet, pivotIdx, anchorRow, anchorCol, spanRows, spanCols);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotSetGrandTotals(uint32_t sheet, uint32_t pivotIdx, bool rowsEnabled, bool colsEnabled) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc =
      fm_workbook_pivot_set_grand_totals(handle_, sheet, pivotIdx, rowsEnabled ? 1 : 0, colsEnabled ? 1 : 0);
  return status_from_rc(rc);
}

emscripten::val JsWorkbook::pivotGetLayout(uint32_t sheet, uint32_t pivotIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_layout_t layout = FM_PIVOT_LAYOUT_COMPACT;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_pivot_get_layout(handle_, sheet, pivotIdx, &layout) : 7000;
  if (rc != 0) {
    layout = FM_PIVOT_LAYOUT_COMPACT;
  }
  o.set("status", status_from_rc(rc));
  o.set("layout", static_cast<uint32_t>(layout));
  return o;
}

JsStatus JsWorkbook::pivotSetLayout(uint32_t sheet, uint32_t pivotIdx, uint32_t layout) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_set_layout(handle_, sheet, pivotIdx, static_cast<std::int32_t>(layout));
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotFieldCount(uint32_t sheet, uint32_t pivotIdx) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_field_count(handle_, sheet, pivotIdx, &count);
  return number_result(rc, static_cast<double>(count));
}

JsAddStyleResult JsWorkbook::pivotFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::string source_name;
  std::string custom_name;
  std::string number_format;
  JsNarrowNumericReader reader("pivotFieldAdd");
  // Every nested read uses this per-call reader so getter/proxy failures and
  // lossy integer conversions are reported before the C ABI is called.
  const char* source_name_ptr = reader.optional_string(spec, "sourceName", source_name, "pivotField.sourceName");
  const char* custom_name_ptr = reader.optional_string(spec, "customName", custom_name, "pivotField.customName");
  const bool subtotal_top = reader.boolean(spec, "subtotalTop", false, "pivotField.subtotalTop");
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
    r.status = binding_error_status(kInvalidArgument, reader.message().c_str());
    return r;
  }

  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_field_add(handle_, sheet, pivotIdx, &c_spec, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotFieldClear(uint32_t sheet, uint32_t pivotIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_clear(handle_, sheet, pivotIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetAxis(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t axis) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc =
      fm_workbook_pivot_field_set_axis(handle_, sheet, pivotIdx, fieldIdx, static_cast<std::int32_t>(axis));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetSort(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool ascending,
                                       const std::string& byField) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const char* by = byField.empty() ? nullptr : byField.c_str();
  fm_status_t rc = fm_workbook_pivot_field_set_sort(handle_, sheet, pivotIdx, fieldIdx, ascending ? 1 : 0, by);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetSubtotalTop(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, bool top) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_set_subtotal_top(handle_, sheet, pivotIdx, fieldIdx, top ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldAddItem(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, const std::string& name,
                                       bool visible) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_add_item(handle_, sheet, pivotIdx, fieldIdx, name.c_str(), visible ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldAddItemAt(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t cacheIndex,
                                         bool visible) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_add_item_at(handle_, sheet, pivotIdx, fieldIdx, cacheIndex, visible ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldClearItems(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_clear_items(handle_, sheet, pivotIdx, fieldIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetItemVisible(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t itemIdx,
                                              bool visible) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc =
      fm_workbook_pivot_field_set_item_visible(handle_, sheet, pivotIdx, fieldIdx, itemIdx, visible ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldAddSubtotalFn(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t agg) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc =
      fm_workbook_pivot_field_add_subtotal_fn(handle_, sheet, pivotIdx, fieldIdx, static_cast<std::int32_t>(agg));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldClearSubtotalFns(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_clear_subtotal_fns(handle_, sheet, pivotIdx, fieldIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetDateGroup(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t granularity,
                                            uint32_t calendar, int32_t startYear, int32_t endYear,
                                            uint32_t intervalDays, double startSerial, double endSerial) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_set_date_group(
      handle_, sheet, pivotIdx, fieldIdx, static_cast<std::int32_t>(granularity), static_cast<std::int32_t>(calendar),
      startYear, endYear, intervalDays, startSerial, endSerial);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldClearDateGroup(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_clear_date_group(handle_, sheet, pivotIdx, fieldIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldSetNumberFormat(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx,
                                               const std::string& utf8) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_set_number_format(handle_, sheet, pivotIdx, fieldIdx, utf8.c_str());
  return status_from_rc(rc);
}

namespace {

// Shared body of `pivotSetRowFieldOrder` / `pivotSetColFieldOrder`.
using PivotFieldOrderFn = fm_status_t (*)(fm_workbook_t*, size_t, size_t, const uint32_t*, size_t);

JsStatus set_pivot_field_order(fm_workbook_t* wb, uint32_t sheet, uint32_t pivotIdx, const emscripten::val& indices,
                               PivotFieldOrderFn fn) {
  if (wb == nullptr) {
    return error_status(7000);
  }
  JsNarrowNumericReader reader("pivotFieldOrder");
  if (indices.isUndefined() || indices.isNull()) {
    return status_from_rc(fn(wb, sheet, pivotIdx, nullptr, 0));
  }
  if (!reader.is_array(indices, "pivotFieldOrder")) {
    return binding_error_status(kInvalidArgument, "pivot field order must be an array");
  }
  const uint32_t length = reader.length(indices, "pivotFieldOrder");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  std::vector<uint32_t> v;
  v.reserve(length);
  for (uint32_t i = 0; i < length && reader.ok(); ++i) {
    // Read the raw element once through the guarded property reader, then
    // let the checked conversion reject a symbol, fraction, non-finite value,
    // or uint32 overflow. No subsequent JS reads occur after failure.
    const std::string key = std::to_string(i);
    const emscripten::val value = reader.value(indices, key.c_str(), "pivotFieldOrder.indices");
    v.push_back(reader.u32_value(value, 0U, "pivotFieldOrder.indices"));
  }
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fn(wb, sheet, pivotIdx, v.empty() ? nullptr : v.data(), v.size()));
}

}  // namespace

JsStatus JsWorkbook::pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices) {
  return set_pivot_field_order(handle_, sheet, pivotIdx, indices, &fm_workbook_pivot_set_row_field_order);
}

JsStatus JsWorkbook::pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices) {
  return set_pivot_field_order(handle_, sheet, pivotIdx, indices, &fm_workbook_pivot_set_col_field_order);
}

JsNumberResult JsWorkbook::pivotDataFieldCount(uint32_t sheet, uint32_t pivotIdx) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_data_field_count(handle_, sheet, pivotIdx, &count);
  return number_result(rc, static_cast<double>(count));
}

void JsWorkbook::build_data_field_spec(emscripten::val spec, fm_pivot_data_field_spec_t& out, std::string& name_buf,
                                       std::string& nfmt_buf, bool& has_nfmt) {
  JsNarrowNumericReader reader("buildDataFieldSpec");
  (void)build_data_field_spec_checked(reader, spec, out, name_buf, nfmt_buf, has_nfmt);
}

JsAddStyleResult JsWorkbook::pivotDataFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  JsNarrowNumericReader reader("pivotDataFieldAdd");
  if (!build_data_field_spec_checked(reader, spec, c_spec, name_buf, nfmt_buf, has_nfmt)) {
    r.status = binding_error_status(kInvalidArgument, reader.message().c_str());
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_data_field_add(handle_, sheet, pivotIdx, &c_spec, &out);
  return index_result(rc, out);
}

JsStatus JsWorkbook::pivotDataFieldClear(uint32_t sheet, uint32_t pivotIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_data_field_clear(handle_, sheet, pivotIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotDataFieldSet(uint32_t sheet, uint32_t pivotIdx, uint32_t dataFieldIdx, emscripten::val spec) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_pivot_data_field_spec_t c_spec{};
  std::string name_buf;
  std::string nfmt_buf;
  bool has_nfmt = false;
  JsNarrowNumericReader reader("pivotDataFieldSet");
  if (!build_data_field_spec_checked(reader, spec, c_spec, name_buf, nfmt_buf, has_nfmt)) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  fm_status_t rc = fm_workbook_pivot_data_field_set(handle_, sheet, pivotIdx, dataFieldIdx, &c_spec);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::pivotFilterCount(uint32_t sheet, uint32_t pivotIdx) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_pivot_filter_count(handle_, sheet, pivotIdx, &count);
  return number_result(rc, static_cast<double>(count));
}

JsStatus JsWorkbook::pivotFilterAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  // `fieldName` is required; an omitted key passes NULL through rather
  // than `js_pull_string`'s empty-string default, which the C ABI's own
  // required-field check does not treat the same as a missing argument.
  std::string field_name;
  std::string value_text;
  JsNarrowNumericReader reader("pivotFilterAdd");
  const char* field_name_ptr = reader.optional_string(spec, "fieldName", field_name, "pivotFilter.fieldName");
  const double value_double = reader.number(spec, "valueDouble", 0.0, "pivotFilter.valueDouble");
  const char* value_text_ptr = reader.optional_string(spec, "valueText", value_text, "pivotFilter.valueText");
  const double value_high_double = reader.number(spec, "valueHighDouble", 0.0, "pivotFilter.valueHighDouble");

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
  c_spec.data_field_index = reader.u32(spec, "dataFieldIndex", 0U, "pivotFilter.dataFieldIndex");

  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }

  fm_status_t rc = fm_workbook_pivot_filter_add(handle_, sheet, pivotIdx, &c_spec);
  return status_from_rc(rc);
}

emscripten::val JsWorkbook::pivotFilterAt(uint32_t sheet, uint32_t pivotIdx, uint32_t filterIdx) const {
  emscripten::val o = emscripten::val::object();
  fm_pivot_filter_spec_t spec{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_pivot_filter_at(handle_, sheet, pivotIdx, filterIdx, &spec) : 7000;
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
  return o;
}

JsStatus JsWorkbook::pivotFilterClear(uint32_t sheet, uint32_t pivotIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_filter_clear(handle_, sheet, pivotIdx);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFilterRemoveAt(uint32_t sheet, uint32_t pivotIdx, uint32_t filterIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_filter_remove_at(handle_, sheet, pivotIdx, filterIdx);
  return status_from_rc(rc);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
