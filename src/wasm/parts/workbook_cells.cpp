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
    return error_status(kBindingInvalidHandle);
  }
  fm_status_t rc = fm_workbook_set_number(handle_, sheet, row, col, value);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setBool(uint32_t sheet, uint32_t row, uint32_t col, bool value) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  fm_status_t rc = fm_workbook_set_bool(handle_, sheet, row, col, value ? 1 : 0);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setError(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val errorCode) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  // Unlike setNumber/setBool/setText, there is no sane zero-value default
  // here: an omitted `errorCode` silently writing `#NULL!` (error code 0)
  // would mask a caller bug instead of surfacing it. `errorCode` is taken
  // as `emscripten::val` rather than `int32_t` so a missing argument is
  // still distinguishable from a literal `0` at this point -- matching
  // the Node binding's `info.Length() < 4` rejection for the same call.
  if (!js_value_present(errorCode)) {
    return binding_error_status(static_cast<int32_t>(formulon::FormulonErrorCode::kBindingNullPointer),
                                "setError: `errorCode` is required");
  }
  fm_status_t rc =
      fm_workbook_set_error(handle_, sheet, row, col, static_cast<fm_error_code_t>(errorCode.as<int32_t>()));
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  fm_status_t rc = fm_workbook_set_text(handle_, sheet, row, col, text.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col, const std::string& phonetic) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  fm_status_t rc = fm_workbook_set_cell_phonetic(handle_, sheet, row, col, phonetic.c_str());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val runs) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setCellPhoneticRuns");
  if (!reader.is_array(runs, "setCellPhoneticRuns.runs")) {
    return binding_error_status(static_cast<int32_t>(formulon::FormulonErrorCode::kInvalidArgument),
                                "setCellPhoneticRuns: `runs` must be an array of { sb, eb, text }");
  }
  const uint32_t count = reader.length(runs, "setCellPhoneticRuns.runs.length");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  // Two passes for the same reason `createTable` needs them: no `c_str()`
  // may be taken before `texts` has finished growing.
  std::vector<std::string> texts;
  texts.reserve(count);
  std::vector<fm_phonetic_run_t> records;
  records.reserve(count);
  for (uint32_t i = 0; i < count && reader.ok(); ++i) {
    const emscripten::val run = reader.array_element(runs, i, "setCellPhoneticRuns.run");
    const emscripten::val sb_value = reader.value(run, "sb", "setCellPhoneticRuns.sb");
    const emscripten::val eb_value = reader.value(run, "eb", "setCellPhoneticRuns.eb");
    const emscripten::val text_value = reader.value(run, "text", "setCellPhoneticRuns.text");
    const uint32_t sb = reader.u32_value(sb_value, 0U, "setCellPhoneticRuns.sb");
    const uint32_t eb = reader.u32_value(eb_value, 0U, "setCellPhoneticRuns.eb");
    texts.push_back(reader.string_value(text_value, "setCellPhoneticRuns.text"));
    records.push_back(fm_phonetic_run_t{sb, eb, nullptr});
  }
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  for (uint32_t i = 0; i < count; ++i) {
    records[i].text = texts[i].c_str();
  }
  const fm_status_t rc = fm_workbook_set_cell_phonetic_runs(handle_, sheet, row, col, records.data(), records.size());
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val properties) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setCellPhoneticProperties");
  const uint32_t font_id = reader.u32(properties, "fontId", 0U, "phoneticProperties.fontId");
  const uint32_t type = reader.u32(properties, "type", 0U, "phoneticProperties.type");
  const uint32_t alignment = reader.u32(properties, "alignment", 0U, "phoneticProperties.alignment");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  const fm_status_t rc = fm_workbook_set_cell_phonetic_properties(handle_, sheet, row, col, font_id, type, alignment);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setBlank(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  fm_status_t rc = fm_workbook_set_blank(handle_, sheet, row, col);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  fm_status_t rc = fm_workbook_set_formula(handle_, sheet, row, col, formula.c_str());
  return status_from_rc(rc);
}

namespace {

// Moves a value-producing C call's outcome into a result envelope.
void adopt_value(fm_status_t rc, const fm_value_t& v, JsStatus& status, JsValue& value) {
  if (rc != 0) {
    status = error_status(rc);
    return;
  }
  value = translate_value(v);
  status = ok_status();
}

}  // namespace

JsCellResult JsWorkbook::getValue(uint32_t sheet, uint32_t row, uint32_t col) const {
  JsCellResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(handle_, sheet, row, col, &v);
  adopt_value(rc, v, r.status, r.value);
  return r;
}

emscripten::val JsWorkbook::getCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* text = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_phonetic(handle_, sheet, row, col, &text) : kBindingInvalidHandle;
  return js_text_result(rc, "value", text);
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
    js_set_cstr(entry, "text", run.text);
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

emscripten::val JsWorkbook::getCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col) const {
  uint32_t font_id = 0;
  uint32_t type = 0;
  uint32_t alignment = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_cell_phonetic_properties(handle_, sheet, row, col,
                                                                                       &font_id, &type, &alignment)
                                            : kBindingInvalidHandle;
  // The payload keys are declared unconditionally, so a failure reports the
  // defaults beside the status rather than dropping them.
  if (rc != 0) {
    font_id = 0;
    type = 0;
    alignment = 0;
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", rc == 0 ? ok_status() : error_status(rc));
  o.set("fontId", font_id);
  o.set("type", type);
  o.set("alignment", alignment);
  return o;
}

JsEvalResult JsWorkbook::evaluateFormulaText(uint32_t sheet, uint32_t row, uint32_t col,
                                             const std::string& formula) const {
  JsEvalResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  fm_value_t v{};
  fm_status_t rc = fm_workbook_evaluate_formula(handle_, sheet, row, col, formula.c_str(), &v);
  adopt_value(rc, v, r.status, r.value);
  return r;
}

JsEvalResult JsWorkbook::evaluateConditionalFormula(uint32_t sheet, uint32_t row, uint32_t col, uint32_t anchorRow,
                                                    uint32_t anchorCol, const std::string& formula) const {
  JsEvalResult r;
  if (handle_ == nullptr) {
    r.status = error_status(kBindingInvalidHandle);
    return r;
  }
  fm_value_t v{};
  fm_status_t rc = fm_workbook_evaluate_cf_formula(handle_, sheet, row, col, anchorRow, anchorCol, formula.c_str(), &v);
  adopt_value(rc, v, r.status, r.value);
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
    return fail(kBindingInvalidHandle);
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
  const char* text = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_lambda_text_at(handle_, sheet, row, col, &text) : kBindingInvalidHandle;
  return js_text_result(rc, "text", text);
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
    o.set("status", error_status(kBindingInvalidHandle));
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
    o.set("status", error_status(kBindingInvalidHandle));
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
  js_set_cstr(o, "name", name);
  js_set_cstr(o, "formula", formula);
  o.set("localSheetId", local_sheet_id);
  return o;
}

emscripten::val JsWorkbook::tableAt(uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
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
  js_set_cstr(o, "name", name);
  js_set_cstr(o, "displayName", display);
  js_set_cstr(o, "ref", ref);
  o.set("sheetIndex", static_cast<uint32_t>(sheet_index));
  return o;
}

JsAddStyleResult JsWorkbook::createTable(emscripten::val spec) {
  JsAddStyleResult out;
  if (handle_ == nullptr) {
    out.status = error_status(kBindingInvalidHandle);
    return out;
  }
  JsNarrowNumericReader reader("createTable");
  const uint32_t sheet = reader.u32(spec, "sheetIndex", 0U, "createTable.sheetIndex");
  const std::string ref = reader.string(spec, "ref", "createTable.ref");
  const std::string name = reader.string(spec, "name", "createTable.name");
  std::string display_name = reader.string(spec, "displayName", "createTable.displayName");
  if (display_name.empty()) {
    display_name = name;
  }
  const std::string style_name = reader.string(spec, "styleName", "createTable.styleName");
  const bool header_row = reader.boolean(spec, "headerRow", true, "createTable.headerRow");
  const bool totals_row = reader.boolean(spec, "totalsRow", false, "createTable.totalsRow");
  const emscripten::val columns = reader.value(spec, "columns", "createTable.columns");
  if (!reader.ok()) {
    out.status = binding_error_status(kInvalidArgument, reader.message().c_str());
    return out;
  }
  if (!reader.is_array(columns, "createTable.columns")) {
    out.status = binding_error_status(kInvalidArgument, "createTable: `columns` must be an array of column names");
    return out;
  }
  const uint32_t count = reader.length(columns, "createTable.columns.length");
  if (!reader.ok()) {
    out.status = binding_error_status(kInvalidArgument, reader.message().c_str());
    return out;
  }
  // The pointer vector is filled in a second pass so that no `c_str()` is
  // taken before `names` has finished growing.
  std::vector<std::string> names;
  names.reserve(count);
  for (uint32_t i = 0; i < count && reader.ok(); ++i) {
    const emscripten::val column = reader.array_element(columns, i, "createTable.columns[]");
    names.push_back(reader.string_value(column, "createTable.columns[]"));
  }
  if (!reader.ok()) {
    out.status = binding_error_status(kInvalidArgument, reader.message().c_str());
    return out;
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
    return error_status(kBindingInvalidHandle);
  }

  // The C ABI keeps `ref` non-null, so an omitted ref is resolved to the
  // current value before forwarding the partial update. Other omitted fields
  // use the C ABI's preservation sentinels directly.
  JsNarrowNumericReader reader("updateTable");
  std::string ref;
  std::string style_name;
  const char* ref_ptr = reader.optional_string(spec, "ref", ref, "updateTable.ref");
  const char* style_name_ptr = reader.optional_string(spec, "styleName", style_name, "updateTable.styleName");
  const auto optional_bool = [&reader, &spec](const char* key, const char* field) {
    const emscripten::val value = reader.value(spec, key, field);
    if (!reader.ok() || !js_value_present(value)) {
      return int32_t{-1};
    }
    return reader.boolean_value(value, false, field) ? int32_t{1} : int32_t{0};
  };
  const int32_t header_row = optional_bool("headerRow", "updateTable.headerRow");
  const int32_t totals_row = optional_bool("totalsRow", "updateTable.totalsRow");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }

  std::string effective_ref;
  if (ref_ptr == nullptr) {
    const char* current_ref = nullptr;
    const char* ignored_name = nullptr;
    const char* ignored_display_name = nullptr;
    std::size_t ignored_sheet = 0;
    const fm_status_t lookup_rc =
        fm_workbook_table_at(handle_, idx, &ignored_name, &ignored_display_name, &current_ref, &ignored_sheet);
    if (lookup_rc != 0) {
      return status_from_rc(lookup_rc);
    }
    effective_ref = current_ref != nullptr ? current_ref : std::string();
    ref_ptr = effective_ref.c_str();
  }
  return status_from_rc(fm_workbook_table_update(handle_, idx, ref_ptr, style_name_ptr, header_row, totals_row));
}

JsStatus JsWorkbook::removeTable(uint32_t idx) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_workbook_table_remove(handle_, idx));
}

emscripten::val JsWorkbook::passthroughAt(uint32_t idx) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  const char* path = nullptr;
  fm_status_t rc = fm_workbook_passthrough_at(handle_, idx, &path);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  js_set_cstr(o, "path", path);
  return o;
}

emscripten::val JsWorkbook::pivotLayout(uint32_t sheet, uint32_t pivotIndex) const {
  if (handle_ == nullptr) {
    return empty_pivot_layout_result(error_status(kBindingInvalidHandle));
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
    arr.set("status", error_status(kBindingInvalidHandle));
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
    js_set_cstr(item, "relId", rec.rel_id);
    js_set_cstr(item, "partPath", rec.part_path);
    js_set_cstr(item, "target", rec.target);
    item.set("targetExternal", rec.target_external != 0);
    item.set("kind", rec.kind);
    arr.set(emitted, item);
    ++emitted;
  }
  arr.set("status", ok_status());
  return arr;
}

// ---- Formula text / range enumeration / display text -------------------

namespace {

emscripten::val formula_envelope(fm_status_t rc, const char* formula) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  const bool present = rc == 0 && formula != nullptr && formula[0] != '\0';
  o.set("formula", present ? emscripten::val(std::string(formula)) : emscripten::val::null());
  return o;
}

emscripten::val display_envelope(fm_status_t rc, const char* text, int32_t display_status) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  js_set_cstr(o, "text", rc == 0 ? text : nullptr);
  o.set("displayStatus", rc == 0 ? display_status : 0);
  return o;
}

/// Reads an optional non-negative integer argument (`undefined` / `null`
/// keep `dflt`). Returns false for anything else that is not a safe integer.
bool js_optional_index(const emscripten::val& v, double dflt, double max, double* out) {
  if (!js_value_present(v)) {
    *out = dflt;
    return true;
  }
  const emscripten::val number = emscripten::val::global("Number");
  if (!number.call<bool>("isInteger", v)) {
    return false;
  }
  const double d = v.as<double>();
  *out = d;
  return d >= 0.0 && d <= max;
}

/// Reads a `{kind, number, boolean, text, errorCode}` value record into
/// `out`; `text` owns the storage a `FM_VAL_TEXT` payload points at. Returns
/// false when any nested field fails validation.
bool pull_value(JsNarrowNumericReader& reader, const emscripten::val& value, std::string& text, fm_value_t* out) {
  out->kind = static_cast<fm_value_kind_t>(reader.i32(value, "kind", 0, "value.kind"));
  text = reader.string(value, "text", "value.text");
  switch (out->kind) {
    case FM_VAL_NUMBER:
      out->u.number = reader.number(value, "number", 0.0, "value.number");
      break;
    case FM_VAL_BOOL:
      out->u.boolean = reader.i32(value, "boolean", 0, "value.boolean") != 0 ? 1 : 0;
      break;
    case FM_VAL_TEXT:
      out->u.text = text.c_str();
      break;
    case FM_VAL_ERROR:
      out->u.error_code = reader.i32(value, "errorCode", 0, "value.errorCode");
      break;
    default:
      break;
  }
  return reader.ok();
}

/// Fills `o.cells` / `o.nextCursor` from a cell-range page and destroys it.
void fill_cells_page(emscripten::val& o, fm_cell_range_t* page) {
  std::size_t count = 0;
  uint64_t next = 0;
  fm_cell_range_count(page, &count);
  fm_cell_range_next_cursor(page, &next);
  emscripten::val cells = emscripten::val::array();
  for (std::size_t i = 0; i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t v{};
    if (fm_cell_range_at(page, i, &row, &col, &formula, &v) != 0) {
      continue;
    }
    emscripten::val cell = emscripten::val::object();
    cell.set("row", row);
    cell.set("col", col);
    cell.set("formula", formula != nullptr ? emscripten::val(std::string(formula)) : emscripten::val::null());
    cell.set("value", translate_value(v));
    cells.call<void>("push", cell);
  }
  fm_cell_range_destroy(page);
  o.set("status", ok_status());
  o.set("cells", cells);
  if (next != UINT64_MAX) {
    o.set("nextCursor", static_cast<double>(next));
  }
}

}  // namespace

emscripten::val JsWorkbook::getFormula(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* formula = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_formula(handle_, sheet, row, col, &formula) : kBindingInvalidHandle;
  return formula_envelope(rc, formula);
}

emscripten::val JsWorkbook::getFormulaR1C1(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* formula = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_formula_r1c1(handle_, sheet, row, col, &formula) : kBindingInvalidHandle;
  return formula_envelope(rc, formula);
}

emscripten::val JsWorkbook::getCellsInRange(uint32_t sheet, emscripten::val range, emscripten::val cursor,
                                            emscripten::val limit) const {
  emscripten::val o = emscripten::val::object();
  o.set("cells", emscripten::val::array());
  o.set("nextCursor", emscripten::val::null());
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  double cursor_value = 0.0;
  double limit_value = 0.0;
  if (!js_optional_index(cursor, 0.0, 9007199254740991.0, &cursor_value) ||
      !js_optional_index(limit, 0.0, 4294967295.0, &limit_value)) {
    o.set("status", binding_error_status(static_cast<int32_t>(formulon::FormulonErrorCode::kInvalidArgument),
                                         "getCellsInRange: `cursor` and `limit` must be non-negative integers"));
    return o;
  }
  JsNarrowNumericReader reader("getCellsInRange");
  const fm_merge_range bounds = js_pull_range(range, &reader);
  if (!reader.ok()) {
    o.set("status", binding_error_status(kInvalidArgument, reader.message().c_str()));
    return o;
  }
  fm_cell_range_t* page = nullptr;
  const fm_status_t rc =
      fm_sheet_cells_in_range(handle_, sheet, bounds.first_row, bounds.first_col, bounds.last_row, bounds.last_col,
                              static_cast<uint64_t>(cursor_value), static_cast<uint32_t>(limit_value), &page);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  fill_cells_page(o, page);
  return o;
}

emscripten::val JsWorkbook::listInvalidCells(uint32_t sheet, emscripten::val cursor, emscripten::val limit) const {
  emscripten::val o = emscripten::val::object();
  o.set("cells", emscripten::val::array());
  o.set("nextCursor", emscripten::val::null());
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  double cursor_value = 0.0;
  double limit_value = 0.0;
  if (!js_optional_index(cursor, 0.0, 9007199254740991.0, &cursor_value) ||
      !js_optional_index(limit, 0.0, 4294967295.0, &limit_value)) {
    o.set("status", binding_error_status(kInvalidArgument,
                                         "listInvalidCells: `cursor` and `limit` must be non-negative integers"));
    return o;
  }
  fm_cell_range_t* page = nullptr;
  const fm_status_t rc = fm_sheet_list_invalid_cells(handle_, sheet, static_cast<uint64_t>(cursor_value),
                                                     static_cast<uint32_t>(limit_value), &page);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  fill_cells_page(o, page);
  return o;
}

emscripten::val JsWorkbook::validateValue(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val value) const {
  emscripten::val o = emscripten::val::object();
  o.set("hasRule", false);
  o.set("valid", true);
  o.set("ruleIndex", 0U);
  o.set("errorStyle", 0U);
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    return o;
  }
  std::string text;
  fm_value_t proposed{};
  JsNarrowNumericReader reader("validateValue");
  if (!pull_value(reader, value, text, &proposed)) {
    o.set("status", binding_error_status(kInvalidArgument, reader.message().c_str()));
    return o;
  }
  fm_validation_outcome outcome{};
  const fm_status_t rc = fm_sheet_validate_value(handle_, sheet, row, col, &proposed, &outcome);
  o.set("status", status_from_rc(rc));
  if (rc == 0) {
    o.set("hasRule", outcome.has_rule != 0);
    o.set("valid", outcome.valid != 0);
    o.set("ruleIndex", outcome.rule_index);
    o.set("errorStyle", static_cast<uint32_t>(outcome.error_style));
  }
  return o;
}

emscripten::val JsWorkbook::getDisplayText(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* text = nullptr;
  int32_t display_status = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_display_text(handle_, sheet, row, col, &text, &display_status)
                             : kBindingInvalidHandle;
  return display_envelope(rc, text, display_status);
}

emscripten::val JsWorkbook::formatValue(emscripten::val value, const std::string& formatCode) const {
  if (handle_ == nullptr) {
    return display_envelope(kBindingInvalidHandle, nullptr, 0);
  }
  std::string text;
  fm_value_t v{};
  JsNarrowNumericReader reader("formatValue");
  if (!pull_value(reader, value, text, &v)) {
    emscripten::val out = emscripten::val::object();
    out.set("status", binding_error_status(kInvalidArgument, reader.message().c_str()));
    out.set("text", std::string());
    out.set("displayStatus", 0);
    return out;
  }
  const char* out = nullptr;
  int32_t display_status = 0;
  const fm_status_t rc = fm_workbook_format_value(handle_, &v, formatCode.c_str(), &out, &display_status);
  return display_envelope(rc, out, display_status);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
