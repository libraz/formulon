// What-If data-table bindings: create, read and remove a data table, and the
// explicit table calculation.

#include <cstdint>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

// Reads the optional `{row, col}` object `spec[key]`; `*has` stays 0 when it is absent or null.
void PullInput(CheckedSpecReader& reader, const Napi::Object& spec, const char* key, int32_t* has, uint32_t* row,
               uint32_t* col) {
  Napi::Object cell;
  if (!reader.Object(spec, key, &cell)) {
    return;
  }
  *has = 1;
  *row = reader.U32(cell, "row", 0U);
  *col = reader.U32(cell, "col", 0U);
}

Napi::Value InputToJs(Napi::Env env, int32_t has, uint32_t row, uint32_t col, int32_t deleted) {
  if (has == 0) {
    return env.Null();
  }
  Napi::Object cell = Napi::Object::New(env);
  cell.Set("row", JsNumber(env, row));
  cell.Set("col", JsNumber(env, col));
  cell.Set("deleted", JsBool(env, deleted));
  return cell;
}

}  // namespace

// Returns the result range's top-left cell, the key `getDataTable` takes.
Napi::Value Workbook::AddDataTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Object spec = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  fm_data_table_t table = FM_STRUCT_INIT(fm_data_table_t);
  Napi::Object rect;
  if (reader.Object(spec, "rect", &rect)) {
    ReadCellRect(reader, rect, &table.rect);
  }
  PullInput(reader, spec, "rowInput", &table.has_row_input, &table.row_input_row, &table.row_input_col);
  PullInput(reader, spec, "columnInput", &table.has_column_input, &table.column_input_row, &table.column_input_col);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_cell_address_t cell{};
  const fm_status_t rc = fm_sheet_add_data_table(handle_, ArgU32(info, 0), &table, &cell);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("row", JsNumber(env, cell.row));
  out.Set("col", JsNumber(env, cell.col));
  return out;
}

Napi::Value Workbook::GetDataTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  fm_data_table_t t = FM_STRUCT_INIT(fm_data_table_t);
  const fm_status_t rc =
      fm_sheet_get_data_table(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &t);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object table = Napi::Object::New(env);
  table.Set("rect", RangeToJs(env, t.rect));
  table.Set("twoDimensional", JsBool(env, t.two_dimensional));
  table.Set("rowOriented", JsBool(env, t.row_oriented));
  table.Set("rowInput", InputToJs(env, t.has_row_input, t.row_input_row, t.row_input_col, t.row_input_deleted));
  table.Set("columnInput",
            InputToJs(env, t.has_column_input, t.column_input_row, t.column_input_col, t.column_input_deleted));
  return OptionalOrNull(env, rc, present, table);
}

Napi::Value Workbook::RemoveDataTable(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_data_table);
}

Napi::Value Workbook::RecalcDataTables(const Napi::CallbackInfo& info) {
  return StatusCall(info, [](fm_workbook_t* handle) { return fm_workbook_recalc_data_tables(handle); });
}

}  // namespace formulon_node
