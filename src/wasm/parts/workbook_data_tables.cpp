//
// JsWorkbook What-If data-table surface: creating, reading and removing a
// data table, and the explicit table calculation.

#include <emscripten/val.h>

#include <cstdint>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

// Reads the optional `{row, col}` object `spec[key]`; `*has` stays 0 when it is absent or null.
void pull_input(JsNarrowNumericReader& reader, const emscripten::val& spec, const char* key, int32_t* has,
                uint32_t* row, uint32_t* col) {
  const emscripten::val cell = reader.value(spec, key, key);
  if (!reader.ok() || !js_value_present(cell)) {
    return;
  }
  *has = 1;
  *row = reader.u32(cell, "row", 0U, key);
  *col = reader.u32(cell, "col", 0U, key);
}

emscripten::val input_to_val(int32_t has, uint32_t row, uint32_t col, int32_t deleted) {
  if (has == 0) {
    return emscripten::val::null();
  }
  emscripten::val cell = emscripten::val::object();
  cell.set("row", row);
  cell.set("col", col);
  cell.set("deleted", deleted != 0);
  return cell;
}

}  // namespace

JsStatus JsWorkbook::addDataTable(uint32_t sheet, emscripten::val spec) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("addDataTable");
  fm_data_table table{};
  const fm_merge_range range = js_pull_range(spec, &reader);
  table.first_row = range.first_row;
  table.first_col = range.first_col;
  table.last_row = range.last_row;
  table.last_col = range.last_col;
  if (js_value_present(spec)) {
    pull_input(reader, spec, "rowInput", &table.has_row_input, &table.row_input_row, &table.row_input_col);
    pull_input(reader, spec, "colInput", &table.has_col_input, &table.col_input_row, &table.col_input_col);
  }
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_add_data_table(handle_, sheet, &table));
}

emscripten::val JsWorkbook::getDataTable(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  o.set("dataTable", emscripten::val::null());
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    return o;
  }
  int32_t present = 0;
  fm_data_table t{};
  const fm_status_t rc = fm_sheet_data_table_at(handle_, sheet, row, col, &present, &t);
  o.set("status", status_from_rc(rc));
  if (rc == 0 && present != 0) {
    emscripten::val table = emscripten::val::object();
    table.set("firstRow", t.first_row);
    table.set("lastRow", t.last_row);
    table.set("firstCol", t.first_col);
    table.set("lastCol", t.last_col);
    table.set("twoDimensional", t.two_dimensional != 0);
    table.set("rowOriented", t.row_oriented != 0);
    table.set("rowInput", input_to_val(t.has_row_input, t.row_input_row, t.row_input_col, t.row_input_deleted));
    table.set("colInput", input_to_val(t.has_col_input, t.col_input_row, t.col_input_col, t.col_input_deleted));
    o.set("dataTable", table);
  }
  return o;
}

JsStatus JsWorkbook::removeDataTable(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_data_table(handle_, sheet, row, col));
}

JsStatus JsWorkbook::recalcDataTables() {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_workbook_recalc_data_tables(handle_));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
