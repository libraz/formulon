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
  if (!succeeded(js_collect_list(handle_, 0U, &count_pivot_caches, &pivot_cache_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

uint32_t JsWorkbook::pivotCacheCreate(uint32_t requestedId) {
  uint32_t out = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_pivot_cache(handle_, requestedId, &out)
                                    : kBindingInvalidHandle)) {
    return 0U;
  }
  return out;
}

void JsWorkbook::pivotCacheRemove(uint32_t cacheId) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_pivot_cache(handle_, cacheId) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotCacheGetWorksheetSource(uint32_t cacheId) const {
  int32_t present = 0;
  fm_pivot_worksheet_source_t source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_pivot_cache_worksheet_source(handle_, cacheId, &present, &source)
                             : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  if (present == 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_or_null_fields(o, {{"ref", source.ref}, {"sheet", source.sheet}, {"name", source.name}});
  return o;
}

void JsWorkbook::pivotCacheSetWorksheetSource(uint32_t cacheId, emscripten::val source) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  std::string ref;
  std::string sheet;
  std::string name;
  JsNarrowNumericReader reader("PivotCache.setWorksheetSource");
  // `null` removes the source element.
  if (source.isNull()) {
    succeeded(fm_workbook_set_pivot_cache_worksheet_source(handle_, cacheId, nullptr));
    return;
  }
  if (!reader.require_object(source, "source")) {
    return;
  }
  const char* ref_ptr = reader.optional_string(source, "ref", ref, "ref");
  const char* sheet_ptr = reader.optional_string(source, "sheet", sheet, "sheet");
  const char* name_ptr = reader.optional_string(source, "name", name, "name");
  if (!reader.ok()) {
    return;
  }
  // A NULL string is an absent attribute.
  fm_pivot_worksheet_source_t c_source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  c_source.ref = ref_ptr;
  c_source.sheet = sheet_ptr;
  c_source.name = name_ptr;
  succeeded(fm_workbook_set_pivot_cache_worksheet_source(handle_, cacheId, &c_source));
}

double JsWorkbook::pivotCacheFieldCount(uint32_t cacheId) const {
  std::size_t count = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pivot_cache_field_count(handle_, cacheId, &count)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

std::string JsWorkbook::pivotCacheFieldName(uint32_t cacheId, uint32_t fieldIndex) const {
  const char* name = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pivot_cache_field_name(handle_, cacheId, fieldIndex, &name)
                                    : kBindingInvalidHandle)) {
    return std::string();
  }
  return string_from_cstr(name);
}

uint32_t JsWorkbook::pivotCacheFieldAdd(uint32_t cacheId, const std::string& name) {
  std::size_t out = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_pivot_cache_field(handle_, cacheId, name.c_str(), &out)
                                    : kBindingInvalidHandle)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::pivotCacheFieldClear(uint32_t cacheId) {
  succeeded(handle_ != nullptr ? fm_workbook_clear_pivot_cache_fields(handle_, cacheId) : kBindingInvalidHandle);
}

double JsWorkbook::pivotCacheFieldSharedItemCount(uint32_t cacheId, uint32_t fieldIndex) const {
  std::size_t count = 0;
  if (!succeeded(handle_ != nullptr
                     ? fm_workbook_get_pivot_cache_field_shared_item_count(handle_, cacheId, fieldIndex, &count)
                     : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

emscripten::val JsWorkbook::pivotCacheFieldSharedItemValue(uint32_t cacheId, uint32_t fieldIndex,
                                                           uint32_t itemIndex) const {
  fm_value_t v{};
  if (!succeeded(handle_ != nullptr
                     ? fm_workbook_get_pivot_cache_field_shared_item_value(handle_, cacheId, fieldIndex, itemIndex, &v)
                     : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return value_to_val(v);
}

double JsWorkbook::pivotCacheFieldAddSharedItemNumber(uint32_t cacheId, uint32_t fieldIndex, double value) {
  std::size_t index = 0;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  if (!succeeded(fm_workbook_add_pivot_cache_field_shared_item_number(handle_, cacheId, fieldIndex, value, &index))) {
    return 0.0;
  }
  return static_cast<double>(index);
}

double JsWorkbook::pivotCacheFieldAddSharedItemDate(uint32_t cacheId, uint32_t fieldIndex, double serial) {
  std::size_t index = 0;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  if (!succeeded(fm_workbook_add_pivot_cache_field_shared_item_date(handle_, cacheId, fieldIndex, serial, &index))) {
    return 0.0;
  }
  return static_cast<double>(index);
}

bool JsWorkbook::pivotCacheFieldSharedItemIsDate(uint32_t cacheId, uint32_t fieldIndex, uint32_t itemIndex) const {
  int32_t isDate = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_is_pivot_cache_field_shared_item_date(handle_, cacheId, fieldIndex,
                                                                                        itemIndex, &isDate)
                                    : kBindingInvalidHandle)) {
    return false;
  }
  return isDate != 0;
}

double JsWorkbook::pivotCacheFieldAddSharedItemText(uint32_t cacheId, uint32_t fieldIndex, const std::string& utf8) {
  std::size_t index = 0;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  if (!succeeded(
          fm_workbook_add_pivot_cache_field_shared_item_text(handle_, cacheId, fieldIndex, utf8.c_str(), &index))) {
    return 0.0;
  }
  return static_cast<double>(index);
}

double JsWorkbook::pivotCacheFieldAddSharedItemBool(uint32_t cacheId, uint32_t fieldIndex, bool value) {
  std::size_t index = 0;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  if (!succeeded(
          fm_workbook_add_pivot_cache_field_shared_item_bool(handle_, cacheId, fieldIndex, value ? 1 : 0, &index))) {
    return 0.0;
  }
  return static_cast<double>(index);
}

double JsWorkbook::pivotCacheFieldAddSharedItemBlank(uint32_t cacheId, uint32_t fieldIndex) {
  std::size_t index = 0;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  if (!succeeded(fm_workbook_add_pivot_cache_field_shared_item_blank(handle_, cacheId, fieldIndex, &index))) {
    return 0.0;
  }
  return static_cast<double>(index);
}

double JsWorkbook::pivotCacheFieldAddSharedItemError(uint32_t cacheId, uint32_t fieldIndex, int32_t errorCode) {
  std::size_t index = 0;
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  if (!succeeded(fm_workbook_add_pivot_cache_field_shared_item_error(
          handle_, cacheId, fieldIndex, static_cast<fm_error_code_t>(errorCode), &index))) {
    return 0.0;
  }
  return static_cast<double>(index);
}

void JsWorkbook::pivotCacheFieldClearSharedItems(uint32_t cacheId, uint32_t fieldIndex) {
  succeeded(handle_ != nullptr ? fm_workbook_clear_pivot_cache_field_shared_items(handle_, cacheId, fieldIndex)
                               : kBindingInvalidHandle);
}

double JsWorkbook::pivotCacheRecordCount(uint32_t cacheId) const {
  std::size_t count = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pivot_cache_record_count(handle_, cacheId, &count)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

uint32_t JsWorkbook::pivotCacheRecordAdd(uint32_t cacheId) {
  std::size_t out = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_pivot_cache_record(handle_, cacheId, &out)
                                    : kBindingInvalidHandle)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::pivotCacheRecordClear(uint32_t cacheId) {
  succeeded(handle_ != nullptr ? fm_workbook_clear_pivot_cache_records(handle_, cacheId) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotCacheRecordValue(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  fm_value_t v{};
  if (!succeeded(fm_workbook_get_pivot_cache_record_value(handle_, cacheId, recordIndex, fieldIndex, &v))) {
    return emscripten::val::undefined();
  }
  // Convert the value first: its text is a read-scratch view the flag call may reuse.
  emscripten::val value = value_to_val(v);
  int32_t is_date = 0;
  if (!succeeded(fm_workbook_is_pivot_cache_record_value_date(handle_, cacheId, recordIndex, fieldIndex, &is_date))) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("value", value);
  o.set("isDate", is_date != 0);
  return o;
}

void JsWorkbook::pivotCacheRecordSetNumber(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, double value) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_pivot_cache_record_number(handle_, cacheId, recordIndex, fieldIndex, value)
                : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheRecordSetDate(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, double serial) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_pivot_cache_record_date(handle_, cacheId, recordIndex, fieldIndex, serial)
                : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheRecordSetText(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex,
                                         const std::string& utf8) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_pivot_cache_record_text(handle_, cacheId, recordIndex, fieldIndex, utf8.c_str())
                : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheRecordSetBool(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex, bool value) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_pivot_cache_record_bool(handle_, cacheId, recordIndex, fieldIndex, value ? 1 : 0)
                : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheRecordSetBlank(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex) {
  succeeded(handle_ != nullptr ? fm_workbook_set_pivot_cache_record_blank(handle_, cacheId, recordIndex, fieldIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheRecordSetError(uint32_t cacheId, uint32_t recordIndex, uint32_t fieldIndex,
                                          int32_t errorCode) {
  succeeded(handle_ != nullptr ? fm_workbook_set_pivot_cache_record_error(handle_, cacheId, recordIndex, fieldIndex,
                                                                          static_cast<fm_error_code_t>(errorCode))
                               : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotCacheGetOptions(uint32_t cacheId) const {
  fm_pivot_cache_options_t opt = FM_STRUCT_INIT(fm_pivot_cache_options_t);
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pivot_cache_options(handle_, cacheId, &opt)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("saveData", opt.save_data != 0);
  js_set_present_or_null(o, "missingItemsLimit", opt.has_missing_items_limit, opt.missing_items_limit);
  o.set("refreshOnLoad", opt.refresh_on_load != 0);
  return o;
}

void JsWorkbook::pivotCacheSetOptions(uint32_t cacheId, emscripten::val options) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (!js_value_present(options)) {
    record_type_error("PivotCache.updateOptions patch must be an object");
    return;
  }
  fm_pivot_cache_options_patch_t opt = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  JsNarrowNumericReader reader("PivotCache.updateOptions");
  js_patch_flag(reader, options, "saveData", "saveData", &opt.save_data_op, &opt.save_data);
  // `null` is the automatic limit (no attribute).
  js_patch_u32(reader, options, "missingItemsLimit", "missingItemsLimit", true, &opt.missing_items_limit_op,
               &opt.missing_items_limit);
  js_patch_flag(reader, options, "refreshOnLoad", "refreshOnLoad", &opt.refresh_on_load_op, &opt.refresh_on_load);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_update_pivot_cache_options(handle_, cacheId, &opt));
}

uint32_t JsWorkbook::pivotCacheCalculatedFieldAdd(uint32_t cacheId, const std::string& name,
                                                  const std::string& formula) {
  std::size_t index = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_pivot_cache_calculated_field(handle_, cacheId, name.c_str(),
                                                                                   formula.c_str(), &index)
                                    : kBindingInvalidHandle)) {
    return 0U;
  }
  return static_cast<uint32_t>(index);
}

void JsWorkbook::pivotCacheCalculatedFieldSetFormula(uint32_t cacheId, uint32_t fieldIndex,
                                                     const std::string& formula) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_pivot_cache_calculated_field_formula(handle_, cacheId, fieldIndex, formula.c_str())
                : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheCalculatedFieldRemove(uint32_t cacheId, uint32_t fieldIndex) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_pivot_cache_calculated_field(handle_, cacheId, fieldIndex)
                               : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::pivotCacheFieldFormula(uint32_t cacheId, uint32_t fieldIndex) const {
  const char* formula = nullptr;
  int32_t present = 0;
  if (!succeeded(handle_ != nullptr
                     ? fm_workbook_get_pivot_cache_field_formula(handle_, cacheId, fieldIndex, &present, &formula)
                     : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return present != 0 ? cstr_or_null(formula) : emscripten::val::null();
}

uint32_t JsWorkbook::pivotCacheCalculatedItemAdd(uint32_t cacheId, uint32_t fieldIndex, const std::string& name,
                                                 const std::string& formula) {
  std::size_t out = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_pivot_cache_calculated_item(handle_, cacheId, fieldIndex,
                                                                                  name.c_str(), formula.c_str(), &out)
                                    : kBindingInvalidHandle)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

double JsWorkbook::pivotCacheCalculatedItemCount(uint32_t cacheId) const {
  std::size_t count = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pivot_cache_calculated_item_count(handle_, cacheId, &count)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

emscripten::val JsWorkbook::pivotCacheCalculatedItemAt(uint32_t cacheId, uint32_t idx) const {
  fm_pivot_calculated_item_info_t item = FM_STRUCT_INIT(fm_pivot_calculated_item_info_t);
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_pivot_cache_calculated_item_at(handle_, cacheId, idx, &item)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("fieldIndex", item.field_index);
  js_set_cstr(o, "name", item.name);
  js_set_cstr(o, "formula", item.formula);
  o.set("scoped", item.scoped != 0);
  return o;
}

void JsWorkbook::pivotCacheCalculatedItemSetFormula(uint32_t cacheId, uint32_t idx, const std::string& formula) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_pivot_cache_calculated_item_formula(handle_, cacheId, idx, formula.c_str())
                : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheCalculatedItemMove(uint32_t cacheId, uint32_t idx, uint32_t newIndex) {
  succeeded(handle_ != nullptr ? fm_workbook_move_pivot_cache_calculated_item(handle_, cacheId, idx, newIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::pivotCacheCalculatedItemRemove(uint32_t cacheId, uint32_t idx) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_pivot_cache_calculated_item(handle_, cacheId, idx)
                               : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
