//
// JsWorkbook cell-value mutators / readers and iteration accessors:
// `setNumber` / `setBool` / `setText` / `setBlank` / `setFormula` /
// `getValue` / `getLambdaText`, plus the `cellCount` / `cellAt` /
// `definedName*` / `table*` / `passthrough*` / `pivotCount` /
// `pivotLayout` / `getExternalLink` iteration surface.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- Cell value / formula setters ---------------------------------------

void JsWorkbook::setNumber(uint32_t sheet, uint32_t row, uint32_t col, double value) {
  succeeded(handle_ != nullptr ? fm_sheet_set_number(handle_, sheet, row, col, value) : kBindingInvalidHandle);
}

void JsWorkbook::setBool(uint32_t sheet, uint32_t row, uint32_t col, bool value) {
  succeeded(handle_ != nullptr ? fm_sheet_set_bool(handle_, sheet, row, col, value ? 1 : 0) : kBindingInvalidHandle);
}

void JsWorkbook::setError(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val errorCode) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // An omitted `errorCode` silently writing `#NULL!` (error code 0) would
  // mask a caller bug. The pre-js guard rejects it as the Node guard does;
  // `emscripten::val` keeps a missing argument distinguishable from `0`
  // here as well.
  if (!js_value_present(errorCode)) {
    record_type_error("Sheet.setError error is required");
    return;
  }
  fm_status_t rc = fm_sheet_set_error(handle_, sheet, row, col, static_cast<fm_error_code_t>(errorCode.as<int32_t>()));
  succeeded(rc);
}

void JsWorkbook::setText(uint32_t sheet, uint32_t row, uint32_t col, const std::string& text) {
  succeeded(handle_ != nullptr ? fm_sheet_set_text(handle_, sheet, row, col, text.c_str()) : kBindingInvalidHandle);
}

void JsWorkbook::setCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col, const std::string& phonetic) {
  succeeded(handle_ != nullptr ? fm_sheet_set_cell_phonetic(handle_, sheet, row, col, phonetic.c_str())
                               : kBindingInvalidHandle);
}

void JsWorkbook::setCellPhoneticRuns(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val runs) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.setCellPhoneticRuns");
  if (!reader.is_array(runs, "runs")) {
    record_type_error("Sheet.setCellPhoneticRuns runs must be an array");
    return;
  }
  const uint32_t count = reader.length(runs, "runs.length");
  if (!reader.ok()) {
    return;
  }
  // Two passes for the same reason `createTable` needs them: no `c_str()`
  // may be taken before `texts` has finished growing.
  std::vector<std::string> texts;
  texts.reserve(count);
  std::vector<fm_phonetic_run_t> records;
  records.reserve(count);
  for (uint32_t i = 0; i < count && reader.ok(); ++i) {
    const emscripten::val run = reader.array_element(runs, i, "runs[]");
    const emscripten::val start_value = reader.value(run, "start", "runs[].start");
    const emscripten::val end_value = reader.value(run, "end", "runs[].end");
    const emscripten::val text_value = reader.value(run, "text", "runs[].text");
    const uint32_t start = reader.u32_value(start_value, 0U, "runs[].start");
    const uint32_t end = reader.u32_value(end_value, 0U, "runs[].end");
    texts.push_back(reader.string_value(text_value, "runs[].text"));
    records.push_back(fm_phonetic_run_t{start, end, nullptr});
  }
  if (!reader.ok()) {
    return;
  }
  for (uint32_t i = 0; i < count; ++i) {
    records[i].text = texts[i].c_str();
  }
  const fm_status_t rc = fm_sheet_set_cell_phonetic_runs(handle_, sheet, row, col, records.data(), records.size());
  succeeded(rc);
}

void JsWorkbook::setCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val properties) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // `null` resets the cell to the unstated triple; `stated` is read-only and ignored.
  if (properties.isNull()) {
    succeeded(fm_sheet_set_cell_phonetic_properties(handle_, sheet, row, col, nullptr));
    return;
  }
  JsNarrowNumericReader reader("Sheet.setCellPhoneticProperties");
  if (!reader.require_object(properties, "properties")) {
    return;
  }
  fm_phonetic_properties_t record = FM_STRUCT_INIT(fm_phonetic_properties_t);
  record.font_index = reader.u32(properties, "fontIndex", 0U, "properties.fontIndex");
  record.type = reader.i32(properties, "type", 0, "properties.type");
  record.alignment = reader.i32(properties, "alignment", 0, "properties.alignment");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_cell_phonetic_properties(handle_, sheet, row, col, &record));
}

void JsWorkbook::setBlank(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_set_blank(handle_, sheet, row, col) : kBindingInvalidHandle);
}

void JsWorkbook::setFormula(uint32_t sheet, uint32_t row, uint32_t col, const std::string& formula) {
  succeeded(handle_ != nullptr ? fm_sheet_set_formula(handle_, sheet, row, col, formula.c_str())
                               : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getValue(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_value_t v{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_value(handle_, sheet, row, col, &v) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return value_to_val(v);
}

emscripten::val JsWorkbook::getCellPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* text = nullptr;
  int32_t present = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_phonetic(handle_, sheet, row, col, &present, &text)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return present != 0 ? cstr_or_null(text) : emscripten::val::null();
}

double JsWorkbook::getCellPhoneticRunCount(uint32_t sheet, uint32_t row, uint32_t col) const {
  std::size_t count = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_phonetic_run_count(handle_, sheet, row, col, &count)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

emscripten::val JsWorkbook::getCellPhoneticRun(uint32_t sheet, uint32_t row, uint32_t col,
                                               uint32_t phoneticRunIndex) const {
  fm_phonetic_run_t run{};
  if (!succeeded(handle_ != nullptr
                     ? fm_sheet_get_cell_phonetic_run_at(handle_, sheet, row, col, phoneticRunIndex, &run)
                     : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("start", run.start);
  o.set("end", run.end);
  js_set_cstr(o, "text", run.text);
  return o;
}

emscripten::val JsWorkbook::getCellPhoneticProperties(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_phonetic_properties_t record = FM_STRUCT_INIT(fm_phonetic_properties_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_phonetic_properties(handle_, sheet, row, col, &record)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("fontIndex", record.font_index);
  o.set("type", record.type);
  o.set("alignment", record.alignment);
  o.set("stated", record.stated != 0);
  return o;
}

bool JsWorkbook::getCellShowPhonetic(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t show = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_show_phonetic(handle_, sheet, row, col, &show)
                                    : kBindingInvalidHandle)) {
    return false;
  }
  return show != 0;
}

void JsWorkbook::setCellShowPhonetic(uint32_t sheet, uint32_t row, uint32_t col, bool show) {
  succeeded(handle_ != nullptr ? fm_sheet_set_cell_show_phonetic(handle_, sheet, row, col, show ? 1 : 0)
                               : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::evaluateConditionalFormula(uint32_t sheet, uint32_t row, uint32_t col, uint32_t anchorRow,
                                                       uint32_t anchorCol, const std::string& formula) const {
  fm_value_t v{};
  if (!succeeded(handle_ != nullptr
                     ? fm_sheet_evaluate_cf_formula(handle_, sheet, row, col, anchorRow, anchorCol, formula.c_str(), &v)
                     : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return emscripten::val(v.u.boolean != 0);
}

emscripten::val JsWorkbook::evaluateFormulaArray(uint32_t sheet, uint32_t row, uint32_t col,
                                                 const std::string& formula) const {
  fm_value_array_t* array = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_sheet_evaluate_formula_array(handle_, sheet, row, col, formula.c_str(), &array)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val out = value_array_to_val(array);
  fm_value_array_destroy(array);
  return out;
}

emscripten::val JsWorkbook::getLambdaText(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t present = 0;
  const char* text = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_lambda_text(handle_, sheet, row, col, &present, &text) : kBindingInvalidHandle;
  return optional_or_null(rc, present,
                          rc == 0 && present != 0 ? emscripten::val(std::string(text)) : emscripten::val::null());
}

// ---- Iteration / metadata accessors -------------------------------------
//
// `cellCount` / `tableCount` / `pivotCount` are now emitted by the
// binding codegen (see
// `src/wasm/generated/workbook_counts.cpp` and
// `src/wasm/generated/sheet_counts.cpp`).

emscripten::val JsWorkbook::cellAt(uint32_t sheet, uint32_t idx) const {
  uint32_t row = 0;
  uint32_t col = 0;
  const char* formula = nullptr;
  fm_value_t v{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_at(handle_, sheet, idx, &row, &col, &formula, &v)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("row", row);
  o.set("col", col);
  o.set("formula", formula != nullptr ? emscripten::val(std::string(formula)) : emscripten::val::null());
  o.set("value", value_to_val(v));
  return o;
}

namespace {

emscripten::val defined_name_to_val(const char* name, const char* formula, int32_t local_sheet_id, const char* comment,
                                    int32_t hidden) {
  emscripten::val item = emscripten::val::object();
  js_set_cstr_fields(item, {{"name", name}, {"formula", formula}});
  js_set_cstr_or_null(item, "comment", comment);
  js_set_present_or_null(item, "scope", local_sheet_id != FM_DEFINED_NAME_SCOPE_WORKBOOK ? 1 : 0, local_sheet_id);
  item.set("hidden", hidden != 0);
  return item;
}

fm_status_t count_defined_names(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_defined_name_count(handle, out);
}

fm_status_t defined_name_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  fm_defined_name_t d = FM_STRUCT_INIT(fm_defined_name_t);
  const fm_status_t rc = fm_workbook_get_defined_name_at(handle, index, &d);
  if (rc == 0) {
    *out = defined_name_to_val(d.name, d.formula, d.local_sheet_id, d.comment, d.hidden);
  }
  return rc;
}

fm_status_t count_passthroughs(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_passthrough_count(handle, out);
}

fm_status_t passthrough_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  const char* path = nullptr;
  const fm_status_t rc = fm_workbook_get_passthrough_at(handle, index, &path);
  if (rc == 0) {
    *out = emscripten::val(string_from_cstr(path));
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listDefinedNames() const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, 0U, &count_defined_names, &defined_name_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

double JsWorkbook::tableCount(uint32_t sheet) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0.0;
  }
  std::size_t count = 0;
  if (!succeeded(fm_sheet_get_table_count(handle_, sheet, &count))) {
    return 0.0;
  }
  return static_cast<double>(count);
}

uint32_t JsWorkbook::tableRuntimeId(uint32_t sheet, uint32_t idx) const {
  uint32_t out = 0;
  succeeded(handle_ != nullptr ? fm_sheet_get_table_runtime_id(handle_, sheet, idx, &out) : kBindingInvalidHandle);
  return out;
}

emscripten::val JsWorkbook::resolveTableRuntimeId(uint32_t id) const {
  std::size_t sheet_index = 0;
  std::size_t table_index = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_resolve_table_runtime_id(handle_, id, &sheet_index, &table_index)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("sheet", static_cast<uint32_t>(sheet_index));
  o.set("index", static_cast<uint32_t>(table_index));
  return o;
}

emscripten::val JsWorkbook::findTable(const std::string& name) const {
  std::size_t sheet_index = 0;
  std::size_t table_index = 0;
  int32_t present = 0;
  if (!succeeded(handle_ != nullptr
                     ? fm_workbook_find_table(handle_, name.c_str(), &present, &sheet_index, &table_index)
                     : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  if (present == 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  o.set("sheet", static_cast<uint32_t>(sheet_index));
  o.set("index", static_cast<uint32_t>(table_index));
  return o;
}

emscripten::val JsWorkbook::tableAt(uint32_t sheet, uint32_t idx) const {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_table_at(handle_, sheet, idx, &info) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  fm_status_t rc = 0;
  emscripten::val columns = emscripten::val::array();
  for (uint32_t i = 0; rc == 0 && i < info.column_count; ++i) {
    fm_table_column_info_t column = FM_STRUCT_INIT(fm_table_column_info_t);
    rc = fm_sheet_get_table_column_at(handle_, sheet, idx, i, &column);
    if (rc == 0) {
      emscripten::val c = emscripten::val::object();
      js_set_cstr(c, "name", column.name);
      c.set("totalsRowFunction", column.totals_row_function);
      js_set_cstr_or_null(c, "totalsRowLabel", column.totals_row_label);
      js_set_cstr_or_null(c, "totalsRowFormula", column.totals_row_formula);
      columns.call<void>("push", c);
    }
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", info.name);
  emscripten::val range = emscripten::val::object();
  range.set("sheet", sheet);
  range.set("rect", cell_rect_to_val(info.range));
  o.set("range", range);
  o.set("headerRow", info.header_row_count != 0);
  o.set("totalsRow", info.totals_row_count != 0);
  o.set("columns", columns);
  return o;
}

emscripten::val JsWorkbook::getTableStyleOptions(uint32_t sheet, uint32_t idx) const {
  fm_table_style_options_t info = FM_STRUCT_INIT(fm_table_style_options_t);
  int32_t present = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_table_style_options(handle_, sheet, idx, &present, &info)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  if (present == 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_or_null(o, "styleName", info.style_name);
  o.set("showFirstColumn", info.show_first_column != 0);
  o.set("showLastColumn", info.show_last_column != 0);
  o.set("showRowStripes", info.show_row_stripes != 0);
  o.set("showColumnStripes", info.show_column_stripes != 0);
  return o;
}

namespace {

/// A boolean toggle of a table style patch: `KEEP` when `key` is omitted, else `SET`.
void js_table_style_toggle(JsNarrowNumericReader& reader, const emscripten::val& spec, const char* key,
                           const char* field, fm_patch_op_t* op, int32_t* value) {
  const emscripten::val v = reader.value(spec, key, field);
  if (!reader.ok() || !js_value_present(v)) {
    return;
  }
  *op = FM_PATCH_OP_SET;
  *value = reader.boolean_value(v, false) ? 1 : 0;
}

}  // namespace

void JsWorkbook::updateTableStyleOptions(uint32_t sheet, uint32_t idx, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Table.updateStyleOptions");
  fm_table_style_options_patch_t p = FM_STRUCT_INIT(fm_table_style_options_patch_t);
  std::string style_name;
  // `null` removes the style name; a string replaces it; an omitted key keeps it.
  const emscripten::val name_value = reader.value(patch, "styleName", "styleName");
  switch (js_patch_key(reader, name_value, "styleName", true)) {
    case JsPatchKey::kClear:
      p.style_name_op = FM_PATCH_OP_CLEAR;
      break;
    case JsPatchKey::kSet:
      p.style_name_op = FM_PATCH_OP_SET;
      p.style_name = reader.optional_string_value(name_value, style_name, "styleName");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  js_table_style_toggle(reader, patch, "showFirstColumn", "showFirstColumn", &p.show_first_column_op,
                        &p.show_first_column);
  js_table_style_toggle(reader, patch, "showLastColumn", "showLastColumn", &p.show_last_column_op, &p.show_last_column);
  js_table_style_toggle(reader, patch, "showRowStripes", "showRowStripes", &p.show_row_stripes_op, &p.show_row_stripes);
  js_table_style_toggle(reader, patch, "showColumnStripes", "showColumnStripes", &p.show_column_stripes_op,
                        &p.show_column_stripes);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_table_style_options(handle_, sheet, idx, &p));
}

uint32_t JsWorkbook::createTable(uint32_t sheet, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.addTable");
  fm_cell_rect_t rect{};
  const emscripten::val range = reader.value(spec, "range", "range");
  if (reader.require_object(range, "range")) {
    rect = js_pull_range(range, &reader);
  }
  const std::string name = reader.string(spec, "name", "name");
  // An omitted `styleName` applies the workbook default; `styled: false` creates no style element.
  std::string style_store;
  const char* style_name = reader.optional_string(spec, "styleName", style_store, "styleName");
  const emscripten::val columns = reader.value(spec, "columns", "columns");
  fm_table_spec_t table = FM_STRUCT_INIT(fm_table_spec_t);
  table.styled = js_tristate(reader, spec, "styled", "styled");
  table.header_row = js_tristate(reader, spec, "headerRow", "headerRow");
  table.totals_row = js_tristate(reader, spec, "totalsRow", "totalsRow");
  table.show_first_column = js_tristate(reader, spec, "showFirstColumn", "showFirstColumn");
  table.show_last_column = js_tristate(reader, spec, "showLastColumn", "showLastColumn");
  table.show_row_stripes = js_tristate(reader, spec, "showRowStripes", "showRowStripes");
  table.show_column_stripes = js_tristate(reader, spec, "showColumnStripes", "showColumnStripes");
  if (!reader.ok()) {
    return 0U;
  }
  if (!reader.is_array(columns, "columns")) {
    record_type_error("Sheet.addTable input.columns must be an array of column names");
    return 0U;
  }
  const uint32_t count = reader.length(columns, "columns.length");
  if (!reader.ok()) {
    return 0U;
  }
  // The pointer vector is filled in a second pass so that no `c_str()` is
  // taken before `names` has finished growing.
  std::vector<std::string> names;
  names.reserve(count);
  for (uint32_t i = 0; i < count && reader.ok(); ++i) {
    const emscripten::val column = reader.array_element(columns, i, "columns[]");
    names.push_back(reader.string_value(column, "columns[]"));
  }
  if (!reader.ok()) {
    return 0U;
  }
  std::vector<const char*> pointers;
  pointers.reserve(count);
  for (const std::string& column : names) {
    pointers.push_back(column.c_str());
  }
  table.range = rect;
  table.name = name.c_str();
  table.column_names = pointers.empty() ? nullptr : pointers.data();
  table.column_count = static_cast<uint32_t>(pointers.size());
  table.style_name = style_name;
  size_t index = 0;
  if (!succeeded(fm_sheet_add_table(handle_, sheet, &table, &index))) {
    return 0U;
  }
  return static_cast<uint32_t>(index);
}

void JsWorkbook::updateTable(uint32_t sheet, uint32_t idx, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }

  JsNarrowNumericReader reader("Table.update");
  fm_table_patch_t patch = FM_STRUCT_INIT(fm_table_patch_t);
  const emscripten::val rect_value = reader.value(spec, "rect", "rect");
  if (reader.ok() && js_value_present(rect_value)) {
    patch.range_op = FM_PATCH_OP_SET;
    patch.range = js_pull_range(rect_value, &reader);
  }
  js_table_style_toggle(reader, spec, "headerRow", "headerRow", &patch.header_row_op, &patch.header_row);
  js_table_style_toggle(reader, spec, "totalsRow", "totalsRow", &patch.totals_row_op, &patch.totals_row);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_table(handle_, sheet, idx, &patch));
}

void JsWorkbook::renameTable(uint32_t sheet, uint32_t idx, const std::string& name) {
  succeeded(handle_ != nullptr ? fm_sheet_rename_table(handle_, sheet, idx, name.c_str()) : kBindingInvalidHandle);
}

void JsWorkbook::setTableColumnTotals(uint32_t sheet, uint32_t idx, uint32_t column, emscripten::val totals) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Table.setColumnTotals");
  std::string label;
  std::string formula;
  const int32_t function =
      reader.i32(totals, "totalsRowFunction", FM_TABLE_TOTALS_ROW_FUNCTION_NONE, "totals.totalsRowFunction");
  const char* label_ptr = reader.optional_string(totals, "totalsRowLabel", label, "totals.totalsRowLabel");
  const char* formula_ptr = reader.optional_string(totals, "totalsRowFormula", formula, "totals.totalsRowFormula");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_table_column_totals(handle_, sheet, idx, column, function, label_ptr, formula_ptr));
}

void JsWorkbook::removeTable(uint32_t sheet, uint32_t idx) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_table(handle_, sheet, idx) : kBindingInvalidHandle);
}

std::string JsWorkbook::renameTableColumn(uint32_t sheet, uint32_t idx, uint32_t column, const std::string& name) {
  const char* final_name = nullptr;
  if (!succeeded(handle_ != nullptr
                     ? fm_sheet_rename_table_column(handle_, sheet, idx, column, name.c_str(), &final_name)
                     : kBindingInvalidHandle)) {
    return std::string();
  }
  return string_from_cstr(final_name);
}

emscripten::val JsWorkbook::getDefinedName(const std::string& name, int32_t localSheetId) const {
  fm_defined_name_t d = FM_STRUCT_INIT(fm_defined_name_t);
  int32_t present = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_defined_name(handle_, name.c_str(), localSheetId, &present, &d)
                             : kBindingInvalidHandle;
  return optional_or_null(rc, present,
                          rc == 0 && present != 0
                              ? defined_name_to_val(d.name, d.formula, localSheetId, d.comment, d.hidden)
                              : emscripten::val::null());
}

void JsWorkbook::removeDefinedName(const std::string& name, int32_t localSheetId) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_defined_name(handle_, name.c_str(), localSheetId)
                               : kBindingInvalidHandle);
}

bool JsWorkbook::hasPassthrough(const std::string& path) const {
  int32_t present = 0;
  succeeded(handle_ != nullptr ? fm_workbook_has_passthrough(handle_, path.c_str(), &present) : kBindingInvalidHandle);
  return present != 0;
}

emscripten::val JsWorkbook::listPassthroughs() const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, 0U, &count_passthroughs, &passthrough_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

emscripten::val JsWorkbook::pivotLayout(uint32_t sheet, uint32_t pivotIndex) const {
  fm_pivot_report_t* cells = nullptr;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_pivot_layout(handle_, sheet, pivotIndex, &cells)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }

  int32_t present = 0;
  fm_cell_rect_t bounds{};
  fm_status_t rc = fm_pivot_report_get_bounds(cells, &present, &bounds);
  if (rc != 0) {
    fm_pivot_report_destroy(cells);
    record_failure(rc);
    return emscripten::val::undefined();
  }

  emscripten::val arr = emscripten::val::array();
  std::size_t count = 0;
  rc = fm_pivot_report_get_cell_count(cells, &count);
  if (rc != 0) {
    fm_pivot_report_destroy(cells);
    record_failure(rc);
    return emscripten::val::undefined();
  }
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_info_t cell = FM_STRUCT_INIT(fm_pivot_cell_info_t);
    if (fm_pivot_report_get_cell_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.set(emitted, pivot_cell_to_val(cell));
    ++emitted;
  }

  fm_pivot_report_destroy(cells);
  emscripten::val o = emscripten::val::object();
  o.set("rect", present != 0 ? cell_rect_to_val(bounds) : emscripten::val::null());
  o.set("cells", arr);
  return o;
}

namespace {

fm_status_t count_external_links(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_external_link_count(handle, out);
}

emscripten::val external_link_to_val(const fm_external_link_info_t& rec) {
  emscripten::val item = emscripten::val::object();
  item.set("bookNumber", rec.book_number);
  js_set_cstr(item, "relId", rec.rel_id);
  js_set_cstr(item, "partPath", rec.part_path);
  js_set_cstr_or_null(item, "target", rec.target);
  item.set("targetExternal", rec.target_external != 0);
  item.set("kind", rec.kind);
  return item;
}

}  // namespace

double JsWorkbook::getExternalLinkCount() const {
  return js_count(handle_, 0U, &count_external_links);
}

emscripten::val JsWorkbook::getExternalLink(uint32_t externalLinkIndex) const {
  fm_external_link_info_t rec = FM_STRUCT_INIT(fm_external_link_info_t);
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_external_link_at(handle_, externalLinkIndex, &rec)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return external_link_to_val(rec);
}

// ---- Formula text / range enumeration / display text -------------------

namespace {

emscripten::val formula_or_null(fm_status_t rc, int32_t present, const char* formula) {
  const bool has_formula = rc == 0 && present != 0 && formula != nullptr;
  return optional_or_null(rc, has_formula ? 1 : 0,
                          has_formula ? emscripten::val(std::string(formula)) : emscripten::val::null());
}

emscripten::val display_text_val(const char* text, int32_t display_status) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "text", text);
  o.set("displayStatus", display_status);
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

/// Builds the `{ cells, nextCursor }` record of a cell-range page and destroys the page.
emscripten::val cells_page_to_val(fm_cell_range_t* page) {
  std::size_t count = 0;
  uint64_t next = 0;
  fm_cell_range_get_cell_count(page, &count);
  fm_cell_range_next_cursor(page, &next);
  emscripten::val cells = emscripten::val::array();
  for (std::size_t i = 0; i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t v{};
    if (fm_cell_range_get_cell_at(page, i, &row, &col, &formula, &v) != 0) {
      continue;
    }
    emscripten::val cell = emscripten::val::object();
    cell.set("row", row);
    cell.set("col", col);
    cell.set("formula", formula != nullptr ? emscripten::val(std::string(formula)) : emscripten::val::null());
    cell.set("value", value_to_val(v));
    cells.call<void>("push", cell);
  }
  fm_cell_range_destroy(page);
  emscripten::val o = emscripten::val::object();
  o.set("cells", cells);
  o.set("nextCursor", next != UINT64_MAX ? emscripten::val(static_cast<double>(next)) : emscripten::val::null());
  return o;
}

}  // namespace

emscripten::val JsWorkbook::getFormula(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* formula = nullptr;
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_formula(handle_, sheet, row, col, &present, &formula) : kBindingInvalidHandle;
  return formula_or_null(rc, present, formula);
}

emscripten::val JsWorkbook::getFormulaR1C1(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* formula = nullptr;
  int32_t present = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_formula_r1c1(handle_, sheet, row, col, &present, &formula)
                                            : kBindingInvalidHandle;
  return formula_or_null(rc, present, formula);
}

emscripten::val JsWorkbook::getCellsInRange(uint32_t sheet, emscripten::val range, emscripten::val cursor,
                                            emscripten::val limit) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  double cursor_value = 0.0;
  double limit_value = 0.0;
  if (!js_optional_index(cursor, 0.0, kMaxSafeInteger, &cursor_value)) {
    record_range_error("Sheet.getCellsInRange options.cursor must be an integer");
    return emscripten::val::undefined();
  }
  if (!js_optional_index(limit, 0.0, kMaxU32, &limit_value)) {
    record_range_error("Sheet.getCellsInRange options.limit must be an integer");
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.getCellsInRange");
  const fm_cell_rect_t bounds = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  fm_cell_range_t* page = nullptr;
  if (!succeeded(fm_sheet_get_cells_in_range(handle_, sheet, &bounds, static_cast<uint64_t>(cursor_value),
                                             static_cast<uint32_t>(limit_value), &page))) {
    return emscripten::val::undefined();
  }
  return cells_page_to_val(page);
}

emscripten::val JsWorkbook::listInvalidCells(uint32_t sheet, emscripten::val cursor, emscripten::val limit) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  double cursor_value = 0.0;
  double limit_value = 0.0;
  if (!js_optional_index(cursor, 0.0, kMaxSafeInteger, &cursor_value)) {
    record_range_error("Sheet.getInvalidCells options.cursor must be an integer");
    return emscripten::val::undefined();
  }
  if (!js_optional_index(limit, 0.0, kMaxU32, &limit_value)) {
    record_range_error("Sheet.getInvalidCells options.limit must be an integer");
    return emscripten::val::undefined();
  }
  fm_cell_range_t* page = nullptr;
  if (!succeeded(fm_sheet_get_invalid_cells(handle_, sheet, static_cast<uint64_t>(cursor_value),
                                            static_cast<uint32_t>(limit_value), &page))) {
    return emscripten::val::undefined();
  }
  return cells_page_to_val(page);
}

emscripten::val JsWorkbook::validateValue(uint32_t sheet, uint32_t row, uint32_t col, emscripten::val value) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  std::string text;
  fm_value_t proposed{};
  if (!val_to_value(value, &proposed, text)) {
    return emscripten::val::undefined();
  }
  fm_validation_outcome_info_t outcome = FM_STRUCT_INIT(fm_validation_outcome_info_t);
  if (!succeeded(fm_sheet_validate_value(handle_, sheet, row, col, &proposed, &outcome))) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("valid", outcome.valid != 0);
  o.set("ruleId", outcome.has_rule != 0 ? cstr_or_null(outcome.rule_id) : emscripten::val::null());
  js_set_present_or_null(o, "errorStyle", outcome.has_rule, static_cast<uint32_t>(outcome.error_style));
  return o;
}

emscripten::val JsWorkbook::getDisplayText(uint32_t sheet, uint32_t row, uint32_t col) const {
  const char* text = nullptr;
  int32_t display_status = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_display_text(handle_, sheet, row, col, &text, &display_status)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return display_text_val(text, display_status);
}

emscripten::val JsWorkbook::formatValue(emscripten::val value, const std::string& formatCode) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  std::string text;
  fm_value_t v{};
  if (!val_to_value(value, &v, text)) {
    return emscripten::val::undefined();
  }
  const char* out = nullptr;
  int32_t display_status = 0;
  if (!succeeded(fm_workbook_format_value(handle_, &v, formatCode.c_str(), &out, &display_status))) {
    return emscripten::val::undefined();
  }
  return display_text_val(out, display_status);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
