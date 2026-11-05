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
  fm_data_table_t table{};
  if (js_value_present(spec)) {
    table.rect = js_pull_range(reader.value(spec, "rect", "rect"), &reader);
    pull_input(reader, spec, "rowInput", &table.has_row_input, &table.row_input_row, &table.row_input_col);
    pull_input(reader, spec, "colInput", &table.has_col_input, &table.col_input_row, &table.col_input_col);
  }
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_sheet_add_data_table(handle_, sheet, &table));
}

emscripten::val JsWorkbook::getDataTable(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t present = 0;
  fm_data_table_t t{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_data_table(handle_, sheet, row, col, &present, &t) : kBindingInvalidHandle;
  if (rc != 0 || present == 0) {
    return optional_result(rc, present, emscripten::val::null());
  }
  emscripten::val table = emscripten::val::object();
  table.set("rect", cell_rect_to_val(t.rect));
  table.set("twoDimensional", t.two_dimensional != 0);
  table.set("rowOriented", t.row_oriented != 0);
  table.set("rowInput", input_to_val(t.has_row_input, t.row_input_row, t.row_input_col, t.row_input_deleted));
  table.set("colInput", input_to_val(t.has_col_input, t.col_input_row, t.col_input_col, t.col_input_deleted));
  return optional_result(rc, present, table);
}

JsStatus JsWorkbook::removeDataTable(uint32_t sheet, uint32_t row, uint32_t col) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_remove_data_table(handle_, sheet, row, col)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::recalcDataTables() {
  return status_from_rc(handle_ != nullptr ? fm_workbook_recalc_data_tables(handle_) : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
