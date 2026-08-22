//
// JsWorkbook cell-value mutators / readers and iteration accessors:
// `setNumber` / `setBool` / `setText` / `setBlank` / `setFormula` /
// `getValue` / `getLambdaText`, plus the `cellCount` / `cellAt` /
// `definedName*` / `table*` / `passthrough*` / `pivotCount` /
// `pivotLayout` / `getExternalLinks` iteration surface.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "utils/error.h"
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

JsStatus JsWorkbook::setError(uint32_t sheet, uint32_t row, uint32_t col, int32_t errorCode) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_error(handle_, sheet, row, col, static_cast<fm_error_code_t>(errorCode));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_text(handle_, sheet, row, col, text.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col, const std::string& phonetic) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_workbook_set_cell_phonetic(handle_, sheet, row, col, phonetic.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val runs) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  if (!runs.isArray()) {
    return binding_error_status(static_cast<int32_t>(formulon::FormulonErrorCode::kInvalidArgument),
                                "setCellPhoneticRuns: `runs` must be an array of { sb, eb, text }");
  }
  const uint32_t count = runs["length"].as<uint32_t>();
  // Two passes for the same reason `createTable` needs them: no `c_str()`
  // may be taken before `texts` has finished growing.
  std::vector<std::string> texts;
  texts.reserve(count);
  std::vector<fm_phonetic_run_t> records;
  records.reserve(count);
  for (uint32_t i = 0; i < count; ++i) {
    const emscripten::val run = runs[i];
    texts.push_back(js_pull_string(run, "text"));
    records.push_back(fm_phonetic_run_t{js_pull_u32(run, "sb", 0U), js_pull_u32(run, "eb", 0U), nullptr});
  }
  for (uint32_t i = 0; i < count; ++i) {
    records[i].text = texts[i].c_str();
  }
  const fm_status_t rc = fm_workbook_set_cell_phonetic_runs(handle_, sheet, row, col, records.data(), records.size());
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

emscripten::val JsWorkbook::getCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    o.set("value", std::string());
    return o;
  }
  const char* text = nullptr;
  fm_status_t rc = fm_workbook_get_cell_phonetic(handle_, sheet, row, col, &text);
  if (rc != 0) {
    o.set("status", error_status(rc));
    o.set("value", std::string());
    return o;
  }
  o.set("status", ok_status());
  o.set("value", std::string(text != nullptr ? text : ""));
  return o;
}

emscripten::val JsWorkbook::getCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  emscripten::val out = emscripten::val::array();
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    o.set("runs", out);
    return o;
  }
  uint32_t count = 0;
  fm_status_t rc = fm_workbook_get_cell_phonetic_run_count(handle_, sheet, row, col, &count);
  for (uint32_t i = 0; rc == 0 && i < count; ++i) {
    fm_phonetic_run_t run{};
    rc = fm_workbook_get_cell_phonetic_run(handle_, sheet, row, col, i, &run);
    if (rc != 0) {
      break;
    }
    emscripten::val entry = emscripten::val::object();
    entry.set("sb", run.sb);
    entry.set("eb", run.eb);
    // Copied immediately: each read refreshes the handle's scratch, so the
    // previous run's pointer is dead by the time the next one lands.
    entry.set("text", std::string(run.text != nullptr ? run.text : ""));
    out.call<void>("push", entry);
  }
  if (rc != 0) {
    o.set("status", error_status(rc));
    o.set("runs", emscripten::val::array());
    return o;
  }
  o.set("status", ok_status());
  o.set("runs", out);
  return o;
}

JsEvalResult JsWorkbook::evaluateFormulaText(uint32_t sheet, uint32_t row, uint32_t col,
                                             const std::string& formula) const {
  JsEvalResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_value_t v{};
  fm_status_t rc = fm_workbook_evaluate_formula(handle_, sheet, row, col, formula.c_str(), &v);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.value = translate_value(v);
  r.status = ok_status();
  return r;
}

JsEvalResult JsWorkbook::evaluateConditionalFormula(uint32_t sheet, uint32_t row, uint32_t col, uint32_t anchorRow,
                                                    uint32_t anchorCol, const std::string& formula) const {
  JsEvalResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_value_t v{};
  fm_status_t rc = fm_workbook_evaluate_cf_formula(handle_, sheet, row, col, anchorRow, anchorCol, formula.c_str(), &v);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.value = translate_value(v);
  r.status = ok_status();
  return r;
}

emscripten::val JsWorkbook::evaluateFormulaArray(uint32_t sheet, uint32_t row, uint32_t col,
                                                 const std::string& formula) const {
  emscripten::val o = emscripten::val::object();
  auto fail = [&o](fm_status_t rc) {
    o.set("status", error_status(rc));
    o.set("rows", 0);
    o.set("cols", 0);
    o.set("cells", emscripten::val::array());
    return o;
  };
  if (handle_ == nullptr) {
    return fail(7000);
  }
  uint32_t rows = 0;
  uint32_t cols = 0;
  fm_status_t rc = fm_workbook_evaluate_formula_array(handle_, sheet, row, col, formula.c_str(), &rows, &cols);
  if (rc != 0) {
    return fail(rc);
  }
  // Build a rows x cols nested array of Value objects (row-major).
  emscripten::val cells = emscripten::val::array();
  for (uint32_t r = 0; r < rows; ++r) {
    emscripten::val js_row = emscripten::val::array();
    for (uint32_t c = 0; c < cols; ++c) {
      const uint32_t index = r * cols + c;
      fm_value_t v{};
      fm_status_t cell_rc = fm_workbook_evaluate_formula_array_cell(handle_, index, &v);
      if (cell_rc != 0) {
        return fail(cell_rc);
      }
      js_row.set(c, emscripten::val(translate_value(v)));
    }
    cells.set(r, js_row);
  }
  o.set("status", ok_status());
  o.set("rows", rows);
  o.set("cols", cols);
  o.set("cells", cells);
  return o;
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
//
// `cellCount` / `definedNameCount` / `tableCount` / `passthroughCount`
// / `pivotCount` are now emitted by the binding codegen (see
// `src/wasm/generated/workbook_counts.cpp` and
// `src/wasm/generated/sheet_counts.cpp`).

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

emscripten::val JsWorkbook::definedNameAt(uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  const char* name = nullptr;
  const char* formula = nullptr;
  int32_t local_sheet_id = -1;
  fm_status_t rc = fm_workbook_defined_name_at(handle_, idx, &name, &formula, &local_sheet_id);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", name != nullptr ? std::string(name) : std::string());
  o.set("formula", formula != nullptr ? std::string(formula) : std::string());
  o.set("localSheetId", local_sheet_id);
  return o;
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

JsAddStyleResult JsWorkbook::createTable(emscripten::val spec) {
  JsAddStyleResult out;
  if (handle_ == nullptr) {
    out.status = error_status(7000);
    return out;
  }
  const uint32_t sheet = js_pull_u32(spec, "sheetIndex", 0U);
  const std::string ref = js_pull_string(spec, "ref");
  const std::string name = js_pull_string(spec, "name");
  std::string display_name = js_pull_string(spec, "displayName");
  if (display_name.empty()) {
    display_name = name;
  }
  const std::string style_name = js_pull_string(spec, "styleName");
  const bool header_row = js_pull_bool(spec, "headerRow", true);
  const bool totals_row = js_pull_bool(spec, "totalsRow", false);
  emscripten::val columns = spec["columns"];
  if (!columns.isArray()) {
    out.status = binding_error_status(static_cast<int32_t>(formulon::FormulonErrorCode::kInvalidArgument),
                                      "createTable: `columns` must be an array of column names");
    return out;
  }
  const uint32_t count = columns["length"].as<uint32_t>();
  // The pointer vector is filled in a second pass so that no `c_str()` is
  // taken before `names` has finished growing.
  std::vector<std::string> names;
  names.reserve(count);
  for (uint32_t i = 0; i < count; ++i) {
    names.push_back(columns[i].as<std::string>());
  }
  std::vector<const char*> pointers;
  pointers.reserve(count);
  for (const std::string& column : names) {
    pointers.push_back(column.c_str());
  }
  size_t index = 0;
  const fm_status_t rc =
      fm_workbook_table_create(handle_, sheet, ref.c_str(), name.c_str(), display_name.c_str(), pointers.data(),
                               pointers.size(), style_name.c_str(), header_row ? 1 : 0, totals_row ? 1 : 0, &index);
  if (rc != 0) {
    out.status = error_status(rc);
    return out;
  }
  out.status = ok_status();
  out.index = static_cast<uint32_t>(index);
  return out;
}

JsStatus JsWorkbook::updateTable(uint32_t idx, emscripten::val spec) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }

  // The C ABI keeps `ref` non-null, so an omitted ref is resolved to the
  // current value before forwarding the partial update. Other omitted fields
  // use the C ABI's preservation sentinels directly.
  emscripten::val ref_value = spec["ref"];
  std::string ref;
  if (ref_value.isUndefined() || ref_value.isNull()) {
    const char* current_ref = nullptr;
    const char* ignored_name = nullptr;
    const char* ignored_display_name = nullptr;
    std::size_t ignored_sheet = 0;
    const fm_status_t lookup_rc =
        fm_workbook_table_at(handle_, idx, &ignored_name, &ignored_display_name, &current_ref, &ignored_sheet);
    if (lookup_rc != 0) {
      return status_from_rc(lookup_rc);
    }
    ref = current_ref != nullptr ? current_ref : std::string();
  } else {
    ref = ref_value.as<std::string>();
  }

  emscripten::val style_value = spec["styleName"];
  std::string style_name;
  const char* style_name_ptr = nullptr;
  if (!style_value.isUndefined() && !style_value.isNull()) {
    style_name = style_value.as<std::string>();
    style_name_ptr = style_name.c_str();
  }

  const auto optional_bool = [&spec](const char* key) {
    const emscripten::val value = spec[key];
    if (value.isUndefined() || value.isNull()) {
      return int32_t{-1};
    }
    return value.as<bool>() ? int32_t{1} : int32_t{0};
  };
  return status_from_rc(fm_workbook_table_update(handle_, idx, ref.c_str(), style_name_ptr, optional_bool("headerRow"),
                                                 optional_bool("totalsRow")));
}

JsStatus JsWorkbook::removeTable(uint32_t idx) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_workbook_table_remove(handle_, idx));
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
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_t cell{};
    if (fm_pivot_cells_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.set(emitted, pivot_cell_to_val(cell));
    ++emitted;
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
    arr.set("status", error_status(7000));
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = fm_workbook_external_link_count(handle_, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  uint32_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_external_link_record_t rec{};
    rc = fm_workbook_external_link_at(handle_, i, &rec);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
    }
    emscripten::val item = emscripten::val::object();
    item.set("index", rec.index);
    item.set("relId", std::string(rec.rel_id != nullptr ? rec.rel_id : ""));
    item.set("partPath", std::string(rec.part_path != nullptr ? rec.part_path : ""));
    item.set("target", std::string(rec.target != nullptr ? rec.target : ""));
    item.set("targetExternal", rec.target_external != 0);
    item.set("kind", rec.kind);
    arr.set(emitted, item);
    ++emitted;
  }
  arr.set("status", ok_status());
  return arr;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
