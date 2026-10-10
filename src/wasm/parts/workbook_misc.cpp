//
// JsWorkbook trace / function-catalog / spill surfaces:
// `precedents` / `dependents` and their shared `trace_to_val` bridge,
// `functionMetadata` / `functionNames` / `localizeFunctionName` /
// `canonicalizeFunctionName` / `localizeFormula` / `canonicalizeFormula` /
// `localeFacts`, and `spillInfo`.

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

namespace {

using ProfileTextMapFn = fm_status_t (*)(const char*, const char*, const char**);

JsStringResult map_profile_text(ProfileTextMapFn fn, const std::string& text, const std::string& profile_id) {
  const char* out = nullptr;
  const fm_status_t rc = fn(text.c_str(), profile_id.c_str(), &out);
  return string_result(rc, out);
}

}  // namespace

// ---- Trace helpers -----------------------------------------------------
//
// Shared bridge for `precedents` / `dependents`: invokes the C ABI
// entry point, copies the result into a `ListResult` of {sheet, row, col}
// value-objects, and frees the C-owned handle.

emscripten::val JsWorkbook::trace_to_val(TraceFn fn, uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(kBindingInvalidHandle));
    return arr;
  }
  fm_cell_nodes_t* nodes = nullptr;
  fm_status_t rc = fn(handle_, sheet, row, col, depth, &nodes);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  const std::size_t count = fm_cell_nodes_count(nodes);
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_node_t n{};
    rc = fm_cell_nodes_at(nodes, i, &n);
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
  arr.set("status", status_from_rc(rc));
  return arr;
}

emscripten::val JsWorkbook::precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  return trace_to_val(fm_workbook_precedents, sheet, row, col, depth);
}

emscripten::val JsWorkbook::dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  return trace_to_val(fm_workbook_dependents, sheet, row, col, depth);
}

// ---- Function catalog --------------------------------------------------

emscripten::val JsWorkbook::functionMetadata(const std::string& name) const {
  emscripten::val o = emscripten::val::object();
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata(name.c_str(), &md);
  if (rc != 0) {
    o.set("ok", false);
    return o;
  }
  o.set("ok", true);
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
  return o;
}

emscripten::val JsWorkbook::functionNames() const {
  emscripten::val arr = emscripten::val::array();
  const std::size_t n = fm_function_count();
  fm_status_t rc = 0;
  for (std::size_t i = 0; i < n; ++i) {
    const char* name = nullptr;
    rc = fm_function_name_at(i, &name);
    if (rc != 0) {
      break;
    }
    arr.set(static_cast<uint32_t>(i), string_from_cstr(name));
  }
  arr.set("status", status_from_rc(rc));
  return arr;
}

JsStringResult JsWorkbook::localizeFunctionName(const std::string& canonical_name,
                                                const std::string& profile_id) const {
  return map_profile_text(fm_function_localize, canonical_name, profile_id);
}

JsStringResult JsWorkbook::canonicalizeFunctionName(const std::string& localized_name,
                                                    const std::string& profile_id) const {
  return map_profile_text(fm_function_canonicalize, localized_name, profile_id);
}

JsStringResult JsWorkbook::localizeFormula(const std::string& formula, const std::string& profile_id) const {
  return map_profile_text(fm_formula_localize, formula, profile_id);
}

JsStringResult JsWorkbook::canonicalizeFormula(const std::string& formula, const std::string& profile_id) const {
  return map_profile_text(fm_formula_canonicalize, formula, profile_id);
}

emscripten::val JsWorkbook::localeFacts(const std::string& profile_id) const {
  emscripten::val o = emscripten::val::object();
  o.set("facts", emscripten::val::null());
  fm_locale_facts_t f{};
  fm_status_t rc = fm_locale_facts(profile_id.c_str(), &f);
  if (rc != 0) {
    o.set("status", status_from_rc(rc));
    return o;
  }
  static const char* const kDateOrders[] = {"mdy", "ymd", "dmy"};
  emscripten::val facts = emscripten::val::object();
  js_set_cstr(facts, "decimalSeparator", f.decimal_separator);
  js_set_cstr(facts, "groupSeparator", f.group_separator);
  js_set_cstr(facts, "listSeparator", f.list_separator);
  js_set_cstr(facts, "arrayColumnSeparator", f.array_column_separator);
  js_set_cstr(facts, "arrayRowSeparator", f.array_row_separator);
  js_set_cstr(facts, "trueName", f.true_name);
  js_set_cstr(facts, "falseName", f.false_name);
  facts.set("dateOrder", std::string(kDateOrders[static_cast<std::size_t>(f.date_order)]));
  js_set_cstr(facts, "currencySymbol", f.currency_symbol);
  facts.set("currencySuffix", f.currency_suffix != 0);
  facts.set("currencySpace", f.currency_space != 0);
  facts.set("currencyDefaultDecimals", f.currency_default_decimals);
  facts.set("measured", f.measured != 0);
  emscripten::val errors = emscripten::val::array();
  const std::size_t n = fm_locale_error_name_count();
  for (std::size_t i = 0; i < n && rc == 0; ++i) {
    const char* canonical = nullptr;
    const char* localized = nullptr;
    std::int32_t measured = 0;
    rc = fm_locale_error_name(profile_id.c_str(), i, &canonical, &localized, &measured);
    if (rc != 0) {
      break;
    }
    emscripten::val item = emscripten::val::object();
    js_set_cstr(item, "canonical", canonical);
    js_set_cstr(item, "localized", localized);
    item.set("measured", measured != 0);
    errors.set(static_cast<uint32_t>(i), item);
  }
  facts.set("errorNames", errors);
  o.set("status", status_from_rc(rc));
  if (rc == 0) {
    o.set("facts", facts);
  }
  return o;
}

// ---- Spill info --------------------------------------------------------

emscripten::val JsWorkbook::spillInfo(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val item = emscripten::val::object();
  item.set("engaged", false);
  item.set("anchorRow", 0U);
  item.set("anchorCol", 0U);
  item.set("rows", 0U);
  item.set("cols", 0U);
  if (handle_ == nullptr) {
    item.set("status", error_status(kBindingInvalidHandle));
    return item;
  }
  fm_spill_info_t info{};
  const fm_status_t rc = fm_workbook_spill_info(handle_, sheet, row, col, &info);
  item.set("status", status_from_rc(rc));
  if (rc != 0) {
    return item;
  }
  item.set("engaged", info.engaged != 0);
  item.set("anchorRow", info.anchor_row);
  item.set("anchorCol", info.anchor_col);
  item.set("rows", info.rows);
  item.set("cols", info.cols);
  return item;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
