// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// JsWorkbook trace / function-catalog / spill surfaces:
// `precedents` / `dependents` and their shared `trace_to_val` bridge,
// `functionMetadata` / `functionNames` / `localizeFunctionName` /
// `canonicalizeFunctionName`, and `spillInfo`.

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

// ---- Trace helpers -----------------------------------------------------
//
// Shared bridge for `precedents` / `dependents`: invokes the C ABI
// entry point, copies the result into a JS array of {sheet, row, col}
// value-objects, and frees the C-owned handle. Returns an empty array
// on any error so the JS side does not need a separate failure path.

emscripten::val JsWorkbook::trace_to_val(TraceFn fn, uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    return arr;
  }
  fm_cell_nodes_t* nodes = nullptr;
  if (fn(handle_, sheet, row, col, depth, &nodes) != 0) {
    return arr;
  }
  const std::size_t count = fm_cell_nodes_count(nodes);
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_node_t n{};
    if (fm_cell_nodes_at(nodes, i, &n) != 0) {
      continue;
    }
    emscripten::val item = emscripten::val::object();
    item.set("sheet", n.sheet);
    item.set("row", n.row);
    item.set("col", n.col);
    arr.set(static_cast<uint32_t>(i), item);
  }
  fm_cell_nodes_destroy(nodes);
  return arr;
}

emscripten::val JsWorkbook::precedents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  return trace_to_val(fm_workbook_precedents, sheet, row, col, depth);
}

emscripten::val JsWorkbook::dependents(uint32_t sheet, uint32_t row, uint32_t col, uint32_t depth) const {
  return trace_to_val(fm_workbook_dependents, sheet, row, col, depth);
}

// ---- Function catalog --------------------------------------------------

emscripten::val JsWorkbook::functionMetadata(const std::string& name, uint32_t locale) const {
  emscripten::val o = emscripten::val::object();
  fm_function_metadata_t md{};
  fm_status_t rc = fm_function_metadata(name.c_str(), static_cast<fm_locale_t>(locale), &md);
  if (rc != 0) {
    o.set("ok", false);
    return o;
  }
  o.set("ok", true);
  o.set("name", md.canonical_name != nullptr ? std::string(md.canonical_name) : std::string());
  o.set("minArity", md.min_arity);
  o.set("maxArity", md.max_arity);
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
  for (std::size_t i = 0; i < n; ++i) {
    const char* name = nullptr;
    if (fm_function_name_at(i, &name) != 0 || name == nullptr) {
      continue;
    }
    arr.set(static_cast<uint32_t>(i), std::string(name));
  }
  return arr;
}

std::string JsWorkbook::localizeFunctionName(const std::string& canonical_name, uint32_t locale) const {
  const char* out = nullptr;
  if (fm_function_localize(canonical_name.c_str(), static_cast<fm_locale_t>(locale), &out) != 0 || out == nullptr) {
    return std::string();
  }
  return std::string(out);
}

std::string JsWorkbook::canonicalizeFunctionName(const std::string& localized_name, uint32_t locale) const {
  const char* out = nullptr;
  if (fm_function_canonicalize(localized_name.c_str(), static_cast<fm_locale_t>(locale), &out) != 0 || out == nullptr) {
    return std::string();
  }
  return std::string(out);
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
    return item;
  }
  fm_spill_info_t info{};
  if (fm_workbook_spill_info(handle_, sheet, row, col, &info) != 0) {
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
