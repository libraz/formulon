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

// ---- PivotCache --------------------------------------------------------

uint32_t JsWorkbook::pivotCacheCount() const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_count(handle_, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

JsAddStyleResult JsWorkbook::pivotCacheIdAt(uint32_t idx) const {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  uint32_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_id_at(handle_, idx, &out);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = out;
  return r;
}

JsAddStyleResult JsWorkbook::pivotCacheCreate(uint32_t requestedId) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  uint32_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_create(handle_, requestedId, &out);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = out;
  return r;
}

JsStatus JsWorkbook::pivotCacheRemove(uint32_t cacheId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_remove(handle_, cacheId);
  return status_from_rc(rc);
}

emscripten::val JsWorkbook::pivotCacheGetWorksheetSource(uint32_t cacheId) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  int32_t present = 0;
  const char* ref = nullptr;
  const char* sheet = nullptr;
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_pivot_cache_get_worksheet_source(handle_, cacheId, &present, &ref, &sheet, &name);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("present", present != 0);
  o.set("ref", std::string(ref != nullptr ? ref : ""));
  o.set("sheet", std::string(sheet != nullptr ? sheet : ""));
  o.set("name", std::string(name != nullptr ? name : ""));
  return o;
}

JsStatus JsWorkbook::pivotCacheSetWorksheetSource(uint32_t cacheId, emscripten::val source) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const bool present = js_pull_bool(source, "present", true);
  const bool has_ref = !source["ref"].isUndefined() && !source["ref"].isNull();
  const bool has_sheet = !source["sheet"].isUndefined() && !source["sheet"].isNull();
  const bool has_name = !source["name"].isUndefined() && !source["name"].isNull();
  const std::string ref = has_ref ? source["ref"].as<std::string>() : std::string();
  const std::string sheet = has_sheet ? source["sheet"].as<std::string>() : std::string();
  const std::string name = has_name ? source["name"].as<std::string>() : std::string();
  fm_status_t rc = fm_workbook_pivot_cache_set_worksheet_source(
      handle_, cacheId, present ? 1 : 0, has_ref ? ref.c_str() : nullptr, has_sheet ? sheet.c_str() : nullptr,
      has_name ? name.c_str() : nullptr);
  return status_from_rc(rc);
}

uint32_t JsWorkbook::pivotCacheFieldCount(uint32_t cacheId) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_field_count(handle_, cacheId, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

JsStringResult JsWorkbook::pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIdx) const {
  JsStringResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_pivot_cache_field_name(handle_, cacheId, fieldIdx, &name);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.value = (name != nullptr) ? std::string(name) : std::string();
  return r;
}

JsAddStyleResult JsWorkbook::pivotCacheFieldAdd(uint32_t cacheId, const std::string& name) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_field_add(handle_, cacheId, name.c_str(), &out);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = static_cast<uint32_t>(out);
  return r;
}

JsStatus JsWorkbook::pivotCacheFieldClear(uint32_t cacheId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_cache_field_clear(handle_, cacheId);
  return status_from_rc(rc);
}

uint32_t JsWorkbook::pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIdx) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_field_shared_item_count(handle_, cacheId, fieldIdx, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
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

uint32_t JsWorkbook::pivotCacheRecordCount(uint32_t cacheId) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_record_count(handle_, cacheId, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

JsAddStyleResult JsWorkbook::pivotCacheRecordAdd(uint32_t cacheId) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_record_add(handle_, cacheId, &out);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = static_cast<uint32_t>(out);
  return r;
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
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = static_cast<uint32_t>(out);
  return r;
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
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_pivot_layout_t layout = FM_PIVOT_LAYOUT_COMPACT;
  fm_status_t rc = fm_workbook_pivot_get_layout(handle_, sheet, pivotIdx, &layout);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("layout", static_cast<uint32_t>(layout));
  return o;
}

JsStatus JsWorkbook::pivotSetLayout(uint32_t sheet, uint32_t pivotIdx, uint32_t layout) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_set_layout(handle_, sheet, pivotIdx, static_cast<fm_pivot_layout_t>(layout));
  return status_from_rc(rc);
}

uint32_t JsWorkbook::pivotFieldCount(uint32_t sheet, uint32_t pivotIdx) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_field_count(handle_, sheet, pivotIdx, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

JsAddStyleResult JsWorkbook::pivotFieldAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  const std::string source_name = js_pull_string(spec, "sourceName");
  const bool has_custom = !spec["customName"].isUndefined() && !spec["customName"].isNull();
  const std::string custom_name = has_custom ? spec["customName"].as<std::string>() : std::string();
  const bool has_nfmt = !spec["numberFormat"].isUndefined() && !spec["numberFormat"].isNull();
  const std::string number_format = has_nfmt ? spec["numberFormat"].as<std::string>() : std::string();

  fm_pivot_field_spec_t c_spec{};
  c_spec.source_name = source_name.c_str();
  c_spec.custom_name = has_custom ? custom_name.c_str() : nullptr;
  c_spec.axis = static_cast<fm_pivot_axis_t>(js_pull_u32(spec, "axis", 0U));
  c_spec.subtotal_top = js_pull_bool(spec, "subtotalTop", false) ? 1 : 0;
  c_spec.number_format = has_nfmt ? number_format.c_str() : nullptr;

  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_field_add(handle_, sheet, pivotIdx, &c_spec, &out);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = static_cast<uint32_t>(out);
  return r;
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
      fm_workbook_pivot_field_set_axis(handle_, sheet, pivotIdx, fieldIdx, static_cast<fm_pivot_axis_t>(axis));
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

JsStatus JsWorkbook::pivotFieldAddAggregation(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx, uint32_t agg) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_add_aggregation(handle_, sheet, pivotIdx, fieldIdx,
                                                           static_cast<fm_pivot_aggregation_t>(agg));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotFieldClearAggregations(uint32_t sheet, uint32_t pivotIdx, uint32_t fieldIdx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_clear_aggregations(handle_, sheet, pivotIdx, fieldIdx);
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
  fm_status_t rc = fm_workbook_pivot_field_add_subtotal_fn(handle_, sheet, pivotIdx, fieldIdx,
                                                           static_cast<fm_pivot_aggregation_t>(agg));
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
                                            uint32_t calendar, int32_t startYear, int32_t endYear) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_pivot_field_set_date_group(
      handle_, sheet, pivotIdx, fieldIdx, static_cast<fm_pivot_date_grouping_t>(granularity),
      static_cast<fm_pivot_calendar_t>(calendar), startYear, endYear);
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

JsStatus JsWorkbook::pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  std::vector<uint32_t> v;
  if (!indices.isUndefined() && !indices.isNull()) {
    const uint32_t len = indices["length"].as<uint32_t>();
    v.reserve(len);
    for (uint32_t i = 0; i < len; ++i) {
      v.push_back(indices[i].as<uint32_t>());
    }
  }
  fm_status_t rc =
      fm_workbook_pivot_set_row_field_order(handle_, sheet, pivotIdx, v.empty() ? nullptr : v.data(), v.size());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIdx, emscripten::val indices) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  std::vector<uint32_t> v;
  if (!indices.isUndefined() && !indices.isNull()) {
    const uint32_t len = indices["length"].as<uint32_t>();
    v.reserve(len);
    for (uint32_t i = 0; i < len; ++i) {
      v.push_back(indices[i].as<uint32_t>());
    }
  }
  fm_status_t rc =
      fm_workbook_pivot_set_col_field_order(handle_, sheet, pivotIdx, v.empty() ? nullptr : v.data(), v.size());
  return status_from_rc(rc);
}

uint32_t JsWorkbook::pivotDataFieldCount(uint32_t sheet, uint32_t pivotIdx) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_data_field_count(handle_, sheet, pivotIdx, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

void JsWorkbook::build_data_field_spec(emscripten::val spec, fm_pivot_data_field_spec_t& out, std::string& name_buf,
                                       std::string& nfmt_buf, bool& has_nfmt) {
  name_buf = js_pull_string(spec, "name");
  has_nfmt = !spec["numberFormat"].isUndefined() && !spec["numberFormat"].isNull();
  nfmt_buf = has_nfmt ? spec["numberFormat"].as<std::string>() : std::string();
  out.name = name_buf.c_str();
  out.field_index = js_pull_u32(spec, "fieldIndex", 0U);
  out.aggregation = static_cast<fm_pivot_aggregation_t>(js_pull_u32(spec, "aggregation", 0U));
  out.number_format = has_nfmt ? nfmt_buf.c_str() : nullptr;
  out.show_as = static_cast<fm_pivot_show_as_t>(js_pull_u32(spec, "showAs", 0U));
  // -1 sentinels are explicitly modelled as int32; treat absent/null as -1.
  if (!spec["showAsBaseField"].isUndefined() && !spec["showAsBaseField"].isNull()) {
    out.show_as_base_field = spec["showAsBaseField"].as<int32_t>();
  } else {
    out.show_as_base_field = -1;
  }
  if (!spec["showAsBaseItem"].isUndefined() && !spec["showAsBaseItem"].isNull()) {
    out.show_as_base_item = spec["showAsBaseItem"].as<int32_t>();
  } else {
    out.show_as_base_item = -1;
  }
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
  build_data_field_spec(spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_data_field_add(handle_, sheet, pivotIdx, &c_spec, &out);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = static_cast<uint32_t>(out);
  return r;
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
  build_data_field_spec(spec, c_spec, name_buf, nfmt_buf, has_nfmt);
  fm_status_t rc = fm_workbook_pivot_data_field_set(handle_, sheet, pivotIdx, dataFieldIdx, &c_spec);
  return status_from_rc(rc);
}

uint32_t JsWorkbook::pivotFilterCount(uint32_t sheet, uint32_t pivotIdx) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_filter_count(handle_, sheet, pivotIdx, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

JsStatus JsWorkbook::pivotFilterAdd(uint32_t sheet, uint32_t pivotIdx, emscripten::val spec) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const std::string field_name = js_pull_string(spec, "fieldName");
  const bool has_text = !spec["valueText"].isUndefined() && !spec["valueText"].isNull();
  const std::string value_text = has_text ? spec["valueText"].as<std::string>() : std::string();

  fm_pivot_filter_spec_t c_spec{};
  c_spec.axis = static_cast<fm_pivot_axis_t>(js_pull_u32(spec, "axis", 0U));
  c_spec.field_name = field_name.c_str();
  c_spec.type = static_cast<fm_pivot_filter_type_t>(js_pull_u32(spec, "type", 0U));
  // valueKind defaults to NONE (-1) when omitted; the C ABI rejects NONE
  // for non-range filters.
  if (!spec["valueKind"].isUndefined() && !spec["valueKind"].isNull()) {
    c_spec.value_kind = static_cast<fm_pivot_filter_value_kind_t>(spec["valueKind"].as<int32_t>());
  } else {
    c_spec.value_kind = FM_PIVOT_FILTER_VALUE_NONE;
  }
  c_spec.value_int =
      (!spec["valueInt"].isUndefined() && !spec["valueInt"].isNull()) ? spec["valueInt"].as<int32_t>() : 0;
  c_spec.value_double =
      (!spec["valueDouble"].isUndefined() && !spec["valueDouble"].isNull()) ? spec["valueDouble"].as<double>() : 0.0;
  c_spec.value_text = has_text ? value_text.c_str() : nullptr;
  if (!spec["valueHighKind"].isUndefined() && !spec["valueHighKind"].isNull()) {
    c_spec.value_high_kind = static_cast<fm_pivot_filter_value_kind_t>(spec["valueHighKind"].as<int32_t>());
  } else {
    c_spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  }
  c_spec.value_high_int =
      (!spec["valueHighInt"].isUndefined() && !spec["valueHighInt"].isNull()) ? spec["valueHighInt"].as<int32_t>() : 0;
  c_spec.value_high_double = (!spec["valueHighDouble"].isUndefined() && !spec["valueHighDouble"].isNull())
                                 ? spec["valueHighDouble"].as<double>()
                                 : 0.0;

  fm_status_t rc = fm_workbook_pivot_filter_add(handle_, sheet, pivotIdx, &c_spec);
  return status_from_rc(rc);
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
