//
// JsWorkbook PivotCache surface: cache list / create / remove, worksheet
// source, field and shared-item mutators, record cells, cache options, and
// calculated fields / items. Thin wrappers over the
// `fm_workbook_*_pivot_cache_*` C ABI; the PivotTable side lives in
// `parts/workbook_pivot.cpp`.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

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

JsStringResult JsWorkbook::pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIndex) const {
  const char* name = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_field_name(handle_, cacheId, fieldIndex, &name)
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

JsNumberResult JsWorkbook::pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIndex) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_field_shared_item_count(handle_, cacheId, fieldIndex, &count)
                             : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(count));
}

JsCellResult JsWorkbook::pivotCacheFieldSharedItemValue(uint32_t cacheId, uint32_t fieldIndex,
                                                        uint32_t itemIndex) const {
  JsCellResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  fm_value_t v{};
  const fm_status_t rc =
      fm_workbook_get_pivot_cache_field_shared_item_value(handle_, cacheId, fieldIndex, itemIndex, &v);
  adopt_value(rc, v, r.status, r.value);
  return r;
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIndex, double value) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_add_pivot_cache_field_shared_item_number(handle_, cacheId, fieldIndex, value)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemDate(uint32_t cacheId, uint32_t fieldIndex, double serial) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_add_pivot_cache_field_shared_item_date(handle_, cacheId, fieldIndex, serial)
                            : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotCacheFieldSharedItemIsDate(uint32_t cacheId, uint32_t fieldIndex,
                                                            uint32_t itemIndex) const {
  int32_t isDate = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_is_pivot_cache_field_shared_item_date(
                                                  handle_, cacheId, fieldIndex, itemIndex, &isDate)
                                            : kBindingInvalidHandle;
  return bool_result(rc, isDate);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIndex, const std::string& utf8) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_add_pivot_cache_field_shared_item_text(
                                                 handle_, cacheId, fieldIndex, utf8.c_str())
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIndex, bool value) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_add_pivot_cache_field_shared_item_bool(
                                                 handle_, cacheId, fieldIndex, value ? 1 : 0)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIndex) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_add_pivot_cache_field_shared_item_blank(handle_, cacheId, fieldIndex)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIndex, int32_t errorCode) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_workbook_add_pivot_cache_field_shared_item_error(handle_, cacheId, fieldIndex,
                                                                       static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIndex) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_clear_pivot_cache_field_shared_items(handle_, cacheId, fieldIndex)
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

JsStatus JsWorkbook::pivotCacheRecordSetNumber(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex,
                                               double value) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_set_pivot_cache_record_number(handle_, cacheId, recordIndex, fieldIndex, value)
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetDate(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex,
                                             double serial) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_set_pivot_cache_record_date(handle_, cacheId, recordIndex, fieldIndex, serial)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetText(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex,
                                             const std::string& utf8) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_set_pivot_cache_record_text(handle_, cacheId, recordIndex,
                                                                                     fieldIndex, utf8.c_str())
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetBool(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, bool value) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_set_pivot_cache_record_bool(handle_, cacheId, recordIndex,
                                                                                     fieldIndex, value ? 1 : 0)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetBlank(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_set_pivot_cache_record_blank(handle_, cacheId, recordIndex, fieldIndex)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheRecordSetError(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex,
                                              int32_t errorCode) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_status_t rc = fm_workbook_set_pivot_cache_record_error(handle_, cacheId, recordIndex, fieldIndex,
                                                            static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
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

JsStatus JsWorkbook::pivotCacheCalculatedFieldSetFormula(uint32_t cacheId, uint32_t fieldIndex,
                                                         const std::string& formula) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_set_pivot_cache_calculated_field_formula(
                                                 handle_, cacheId, fieldIndex, formula.c_str())
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheCalculatedFieldRemove(uint32_t cacheId, uint32_t fieldIndex) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_remove_pivot_cache_calculated_field(handle_, cacheId, fieldIndex)
                            : kBindingInvalidHandle);
}

JsStringResult JsWorkbook::pivotCacheFieldFormula(uint32_t cacheId, uint32_t fieldIndex) const {
  const char* formula = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_field_formula(handle_, cacheId, fieldIndex, &formula)
                             : kBindingInvalidHandle;
  return string_result(rc, rc == 0 ? formula : nullptr);
}

JsAddStyleResult JsWorkbook::pivotCacheCalculatedItemAdd(uint32_t cacheId, uint32_t fieldIndex, const std::string& name,
                                                         const std::string& formula) {
  std::size_t out = 0;
  return index_result(handle_ != nullptr ? fm_workbook_add_pivot_cache_calculated_item(
                                               handle_, cacheId, fieldIndex, name.c_str(), formula.c_str(), &out)
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

JsStatus JsWorkbook::pivotCacheCalculatedItemMove(uint32_t cacheId, uint32_t idx, uint32_t newIndex) {
  return status_from_rc(handle_ != nullptr
                            ? fm_workbook_move_pivot_cache_calculated_item(handle_, cacheId, idx, newIndex)
                            : kBindingInvalidHandle);
}

JsStatus JsWorkbook::pivotCacheCalculatedItemRemove(uint32_t cacheId, uint32_t idx) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_remove_pivot_cache_calculated_item(handle_, cacheId, idx)
                                           : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
