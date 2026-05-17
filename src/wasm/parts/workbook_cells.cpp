// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// JsWorkbook cell-value mutators / readers and iteration accessors:
// `setNumber` / `setBool` / `setText` / `setBlank` / `setFormula` /
// `getValue` / `getLambdaText`, plus the `cellCount` / `cellAt` /
// `definedName*` / `table*` / `passthrough*` / `pivotCount` /
// `pivotLayout` / `getExternalLinks` iteration surface.

#include <emscripten/val.h>

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- Cell value / formula setters ---------------------------------------

JsStatus JsWorkbook::setNumber(uint32_t sheet, uint32_t row, uint32_t col, double value) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_number(handle_, sheet, row, col, value);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setBool(uint32_t sheet, uint32_t row, uint32_t col, bool value) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_bool(handle_, sheet, row, col, value ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_text(handle_, sheet, row, col, text.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setBlank(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_blank(handle_, sheet, row, col);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_formula(handle_, sheet, row, col, formula.c_str());
  return status_from_rc(rc);
}

JsCellResult JsWorkbook::getValue(uint32_t sheet, uint32_t row, uint32_t col) const {
  JsCellResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(handle_, sheet, row, col, &v);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.value = translate_value(v);
  r.status = ok_status();
  return r;
}

emscripten::val JsWorkbook::getLambdaText(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    o.set("text", std::string());
    return o;
  }
  const char* text = nullptr;
  fm_status_t rc = fm_workbook_lambda_text_at(handle_, sheet, row, col, &text);
  if (rc != 0) {
    o.set("status", error_status(rc));
    o.set("text", std::string());
    return o;
  }
  o.set("status", ok_status());
  o.set("text", std::string(text != nullptr ? text : ""));
  return o;
}

// ---- Iteration / metadata accessors -------------------------------------

uint32_t JsWorkbook::cellCount(uint32_t sheet) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_cell_count(handle_, sheet, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

emscripten::val JsWorkbook::cellAt(uint32_t sheet, uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  uint32_t row = 0;
  uint32_t col = 0;
  const char* formula = nullptr;
  fm_value_t v{};
  fm_status_t rc = fm_workbook_cell_at(handle_, sheet, idx, &row, &col, &formula, &v);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("row", row);
  o.set("col", col);
  o.set("formula", formula != nullptr ? emscripten::val(std::string(formula)) : emscripten::val::null());
  o.set("value", translate_value(v));
  return o;
}

uint32_t JsWorkbook::definedNameCount() const {
  if (handle_ == nullptr) {
    return 0;
  }
  return static_cast<uint32_t>(fm_workbook_defined_name_count(handle_));
}

emscripten::val JsWorkbook::definedNameAt(uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  const char* name = nullptr;
  const char* formula = nullptr;
  fm_status_t rc = fm_workbook_defined_name_at(handle_, idx, &name, &formula);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", name != nullptr ? std::string(name) : std::string());
  o.set("formula", formula != nullptr ? std::string(formula) : std::string());
  return o;
}

uint32_t JsWorkbook::tableCount() const {
  if (handle_ == nullptr) {
    return 0;
  }
  return static_cast<uint32_t>(fm_workbook_table_count(handle_));
}

emscripten::val JsWorkbook::tableAt(uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  const char* name = nullptr;
  const char* display = nullptr;
  const char* ref = nullptr;
  std::size_t sheet_index = 0;
  fm_status_t rc = fm_workbook_table_at(handle_, idx, &name, &display, &ref, &sheet_index);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", name != nullptr ? std::string(name) : std::string());
  o.set("displayName", display != nullptr ? std::string(display) : std::string());
  o.set("ref", ref != nullptr ? std::string(ref) : std::string());
  o.set("sheetIndex", static_cast<uint32_t>(sheet_index));
  return o;
}

uint32_t JsWorkbook::passthroughCount() const {
  if (handle_ == nullptr) {
    return 0;
  }
  return static_cast<uint32_t>(fm_workbook_passthrough_count(handle_));
}

emscripten::val JsWorkbook::passthroughAt(uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  const char* path = nullptr;
  fm_status_t rc = fm_workbook_passthrough_at(handle_, idx, &path);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("path", path != nullptr ? std::string(path) : std::string());
  return o;
}

uint32_t JsWorkbook::pivotCount(uint32_t sheet) const {
  if (handle_ == nullptr) {
    return 0;
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_count(handle_, sheet, &count) != 0) {
    return 0;
  }
  return static_cast<uint32_t>(count);
}

emscripten::val JsWorkbook::pivotLayout(uint32_t sheet, uint32_t pivotIndex) const {
  if (handle_ == nullptr) {
    return empty_pivot_layout_result(error_status(7000));
  }

  fm_pivot_cells_t* cells = nullptr;
  fm_status_t rc = fm_workbook_pivot_layout(handle_, sheet, pivotIndex, &cells);
  if (rc != 0) {
    return empty_pivot_layout_result(error_status(rc));
  }

  uint32_t top = 0;
  uint32_t left = 0;
  uint32_t rows = 0;
  uint32_t cols = 0;
  rc = fm_pivot_cells_bounds(cells, &top, &left, &rows, &cols);
  if (rc != 0) {
    fm_pivot_cells_destroy(cells);
    return empty_pivot_layout_result(error_status(rc));
  }

  emscripten::val arr = emscripten::val::array();
  const std::size_t count = fm_pivot_cells_count(cells);
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_t cell{};
    if (fm_pivot_cells_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.set(i, pivot_cell_to_val(cell));
  }

  fm_pivot_cells_destroy(cells);
  emscripten::val o = emscripten::val::object();
  o.set("status", ok_status());
  o.set("top", top);
  o.set("left", left);
  o.set("rows", rows);
  o.set("cols", cols);
  o.set("cells", arr);
  return o;
}

emscripten::val JsWorkbook::getExternalLinks() const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    return arr;
  }
  uint32_t count = 0;
  if (fm_workbook_external_link_count(handle_, &count) != 0) {
    return arr;
  }
  for (uint32_t i = 0; i < count; ++i) {
    fm_external_link_record_t rec{};
    if (fm_workbook_external_link_at(handle_, i, &rec) != 0) {
      continue;
    }
    emscripten::val item = emscripten::val::object();
    item.set("index", rec.index);
    item.set("relId", std::string(rec.rel_id != nullptr ? rec.rel_id : ""));
    item.set("partPath", std::string(rec.part_path != nullptr ? rec.part_path : ""));
    item.set("target", std::string(rec.target != nullptr ? rec.target : ""));
    item.set("targetExternal", rec.target_external != 0);
    item.set("kind", rec.kind);
    arr.set(i, item);
  }
  return arr;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
