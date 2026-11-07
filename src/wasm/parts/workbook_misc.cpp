//
// JsWorkbook trace / spill surfaces and the workbook-free catalog module
// functions: `precedents` / `dependents` and their shared `trace_to_val`
// bridge, `function_metadata` / `function_names` / `localize_function_name` /
// `canonicalize_function_name` / `localize_formula` / `canonicalize_formula` /
// `locale_facts` / `list_locale_error_names`, `compare_text` /
// `collation_classes`, `spillInfo`, and the legacy array formula accessors.

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

using ProfileTextMapFn = fm_status_t (*)(const char*, const char*, const char**);

std::string map_profile_text(ProfileTextMapFn fn, const std::string& text, const std::string& profile_id) {
  const char* out = nullptr;
  if (!succeeded(fn(text.c_str(), profile_id.c_str(), &out))) {
    return std::string();
  }
  return string_from_cstr(out);
}

}  // namespace

// ---- Trace helpers -----------------------------------------------------
//
// Shared bridge for `precedents` / `dependents`: invokes the C ABI
// entry point, copies the result into an array of {sheet, row, col}
// value-objects, and frees the C-owned handle.

emscripten::val JsWorkbook::trace_to_val(TraceFn fn, uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  fm_cell_nodes_t* nodes = nullptr;
  if (!succeeded(fn(handle_, sheet, row, col, depth, &nodes))) {
    return emscripten::val::undefined();
  }
  emscripten::val arr = emscripten::val::array();
  std::size_t count = 0;
  fm_status_t rc = fm_cell_nodes_get_node_count(nodes, &count);
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_node_t n = FM_STRUCT_INIT(fm_cell_node_t);
    rc = fm_cell_nodes_get_node_at(nodes, i, &n);
    if (rc != 0) {
      break;
    }
    emscripten::val item = emscripten::val::object();
    item.set("sheet", n.sheet);
    item.set("row", n.row);
    item.set("col", n.col);
    arr.set(static_cast<uint32_t>(i), item);
  }
  fm_cell_nodes_destroy(nodes);
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return arr;
}

emscripten::val JsWorkbook::precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  return trace_to_val(fm_sheet_get_precedents, sheet, row, col, depth);
}

emscripten::val JsWorkbook::dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  return trace_to_val(fm_sheet_get_dependents, sheet, row, col, depth);
}

// ---- Function catalog --------------------------------------------------

emscripten::val function_metadata(const std::string& name) {
  int32_t present = 0;
  fm_function_metadata_t md = FM_STRUCT_INIT(fm_function_metadata_t);
  const fm_status_t rc = fm_get_function_metadata(name.c_str(), &present, &md);
  if (rc != 0 || present == 0) {
    return optional_or_null(rc, present, emscripten::val::null());
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", md.canonical_name);
  o.set("minArity", md.min_arity);
  // `0xFFFFFFFF` is the unbounded / unknown-arity sentinel; surface it as
  // `null` so JS callers do not mistake it for a concrete upper bound.
  if (md.max_arity == 0xFFFFFFFFu) {
    o.set("maxArity", emscripten::val::null());
  } else {
    o.set("maxArity", md.max_arity);
  }
  o.set("availability", static_cast<uint32_t>(md.availability));
  if (md.signature_template != nullptr) {
    o.set("signatureTemplate", std::string(md.signature_template));
  }
  if (md.description != nullptr) {
    o.set("description", std::string(md.description));
  }
  return optional_or_null(rc, present, o);
}

emscripten::val function_names() {
  emscripten::val arr = emscripten::val::array();
  std::size_t n = 0;
  fm_status_t rc = fm_get_function_count(&n);
  for (std::size_t i = 0; rc == 0 && i < n; ++i) {
    const char* name = nullptr;
    rc = fm_get_function_name_at(i, &name);
    if (rc != 0) {
      break;
    }
    arr.set(static_cast<uint32_t>(i), string_from_cstr(name));
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return arr;
}

std::string localize_function_name(const std::string& canonical_name, const std::string& profile_id) {
  return map_profile_text(fm_localize_function_name, canonical_name, profile_id);
}

std::string canonicalize_function_name(const std::string& localized_name, const std::string& profile_id) {
  return map_profile_text(fm_canonicalize_function_name, localized_name, profile_id);
}

std::string localize_formula(const std::string& formula, const std::string& profile_id) {
  return map_profile_text(fm_localize_formula, formula, profile_id);
}

std::string canonicalize_formula(const std::string& formula, const std::string& profile_id) {
  return map_profile_text(fm_canonicalize_formula, formula, profile_id);
}

namespace {

// Appends profile `profile_id`'s localized error names to `list`, in error-code order.
fm_status_t collect_locale_error_names(const std::string& profile_id, emscripten::val& list) {
  std::size_t n = 0;
  fm_status_t rc = fm_get_locale_error_name_count(&n);
  for (std::size_t i = 0; i < n && rc == 0; ++i) {
    const char* canonical = nullptr;
    const char* localized = nullptr;
    std::int32_t measured = 0;
    rc = fm_get_locale_error_name_at(profile_id.c_str(), i, &canonical, &localized, &measured);
    if (rc != 0) {
      break;
    }
    emscripten::val item = emscripten::val::object();
    js_set_cstr(item, "canonical", canonical);
    js_set_cstr(item, "localized", localized);
    item.set("measured", measured != 0);
    list.set(static_cast<uint32_t>(i), item);
  }
  return rc;
}

}  // namespace

emscripten::val locale_facts(const std::string& profile_id) {
  fm_locale_facts_t f = FM_STRUCT_INIT(fm_locale_facts_t);
  if (!succeeded(fm_get_locale_facts(profile_id.c_str(), &f))) {
    return emscripten::val::undefined();
  }
  emscripten::val facts = emscripten::val::object();
  js_set_cstr(facts, "decimalSeparator", f.decimal_separator);
  js_set_cstr(facts, "groupSeparator", f.group_separator);
  js_set_cstr(facts, "listSeparator", f.list_separator);
  js_set_cstr(facts, "arrayColumnSeparator", f.array_column_separator);
  js_set_cstr(facts, "arrayRowSeparator", f.array_row_separator);
  js_set_cstr(facts, "trueName", f.true_name);
  js_set_cstr(facts, "falseName", f.false_name);
  facts.set("dateOrder", static_cast<int32_t>(f.date_order));
  js_set_cstr(facts, "currencySymbol", f.currency_symbol);
  facts.set("currencySuffix", f.currency_suffix != 0);
  facts.set("currencySpace", f.currency_space != 0);
  facts.set("currencyDefaultDecimals", f.currency_default_decimals);
  facts.set("measured", f.measured != 0);
  facts.set("collationMeasured", f.collation_measured != 0);
  emscripten::val errors = emscripten::val::array();
  if (!succeeded(collect_locale_error_names(profile_id, errors))) {
    return emscripten::val::undefined();
  }
  facts.set("errorNames", errors);
  return facts;
}

emscripten::val list_locale_error_names(const std::string& profile_id) {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(collect_locale_error_names(profile_id, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

// ---- Spill info --------------------------------------------------------

emscripten::val JsWorkbook::spillInfo(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_spill_info(handle_, sheet, row, col, &info)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val item = emscripten::val::object();
  item.set("engaged", info.engaged != 0);
  item.set("blocked", info.blocked != 0);
  item.set("rect", info.engaged != 0 || info.blocked != 0 ? cell_rect_to_val(info.rect) : emscripten::val::null());
  return item;
}

// ---- Legacy array (CSE) formulas ---------------------------------------

emscripten::val JsWorkbook::getArrayFormula(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_array_formula(handle_, sheet, row, col, &info)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  if (info.engaged == 0) {
    return emscripten::val::null();
  }
  emscripten::val array = emscripten::val::object();
  array.set("rect", cell_rect_to_val(info.rect));
  array.set("formula", std::string(info.formula));
  return array;
}

void JsWorkbook::setArrayFormula(uint32_t sheet, emscripten::val range, const std::string& formula) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setArrayFormula");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_array_formula(handle_, sheet, &rect, formula.c_str()));
}

double compare_text(const std::string& a, const std::string& b, int32_t strength, const std::string& profile_id) {
  int32_t order = 0;
  if (!succeeded(fm_compare_text(a.c_str(), b.c_str(), strength, profile_id.c_str(), &order))) {
    return 0.0;
  }
  return order;
}

double convert_civil_to_serial(emscripten::val civil, bool date1904) {
  JsNarrowNumericReader reader("convertCivilToSerial");
  if (!reader.require_object(civil, "civil")) {
    return 0.0;
  }
  fm_civil_time_t value{};
  value.year = reader.i32(civil, "year", 0, "civil.year");
  value.month = reader.i32(civil, "month", 0, "civil.month");
  value.day = reader.i32(civil, "day", 0, "civil.day");
  value.hour = reader.i32(civil, "hour", 0, "civil.hour");
  value.minute = reader.i32(civil, "minute", 0, "civil.minute");
  value.second = reader.i32(civil, "second", 0, "civil.second");
  if (!reader.ok()) {
    return 0.0;
  }
  double serial = 0.0;
  if (!succeeded(fm_convert_civil_to_serial(&value, date1904 ? 1 : 0, &serial))) {
    return 0.0;
  }
  return serial;
}

emscripten::val convert_serial_to_civil(double serial, bool date1904) {
  fm_civil_time_t value{};
  if (!succeeded(fm_convert_serial_to_civil(serial, date1904 ? 1 : 0, &value))) {
    return emscripten::val::undefined();
  }
  emscripten::val civil = emscripten::val::object();
  civil.set("year", value.year);
  civil.set("month", value.month);
  civil.set("day", value.day);
  civil.set("hour", value.hour);
  civil.set("minute", value.minute);
  civil.set("second", value.second);
  return civil;
}

emscripten::val collation_classes(emscripten::val values, int32_t strength, const std::string& profile_id) {
  JsNarrowNumericReader reader("getCollationClasses");
  if (!reader.is_array(values, "values")) {
    if (reader.ok()) {
      record_type_error("getCollationClasses: values must be an array");
    }
    return emscripten::val::undefined();
  }
  const uint32_t n = reader.length(values, "values");
  // Text payloads must keep their addresses until the C call returns.
  std::vector<std::string> texts(n);
  std::vector<fm_value_t> items(n);
  for (uint32_t i = 0; i < n && reader.ok(); ++i) {
    const emscripten::val element = reader.array_element(values, i, "values[]");
    if (!reader.ok()) {
      break;
    }
    items[i] = fm_value_t{};
    if (!val_to_value(element, &items[i], texts[i])) {
      return emscripten::val::undefined();
    }
  }
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  std::vector<uint32_t> ids(n);
  uint32_t class_count = 0;
  if (!succeeded(fm_get_collation_classes(items.data(), n, strength, profile_id.c_str(), n == 0 ? nullptr : ids.data(),
                                          &class_count))) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  emscripten::val u32 = emscripten::val::global("Uint32Array").new_(n);
  if (n != 0) {
    // The destination is a standalone JS buffer; `set` copies the transient WASM view.
    u32.call<void>("set", emscripten::val(emscripten::typed_memory_view(n, ids.data())));
  }
  o.set("classIds", u32);
  o.set("classCount", class_count);
  return o;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
