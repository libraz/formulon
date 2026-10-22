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

Napi::Value Workbook::AddDataTable(const Napi::CallbackInfo& info) {
  fm_data_table table{};
  return InvokeSheetSpec(
      info, "addDataTable expects (sheet:number, spec:object)",
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        fm_merge_range range{};
        ReadMergeRange(reader, spec, &range);
        table.first_row = range.first_row;
        table.first_col = range.first_col;
        table.last_row = range.last_row;
        table.last_col = range.last_col;
        PullInput(reader, spec, "rowInput", &table.has_row_input, &table.row_input_row, &table.row_input_col);
        PullInput(reader, spec, "colInput", &table.has_col_input, &table.col_input_row, &table.col_input_col);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_add_data_table(handle, ArgU32(info, 0), &table); });
}

Napi::Value Workbook::GetDataTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("dataTable", env.Null());
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  int32_t present = 0;
  fm_data_table t{};
  const fm_status_t rc =
      fm_sheet_data_table_at(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &t);
  out.Set("status", MakeStatus(env, rc));
  if (rc != 0 || present == 0) {
    return out;
  }
  Napi::Object table = Napi::Object::New(env);
  table.Set("firstRow", JsNumber(env, t.first_row));
  table.Set("lastRow", JsNumber(env, t.last_row));
  table.Set("firstCol", JsNumber(env, t.first_col));
  table.Set("lastCol", JsNumber(env, t.last_col));
  table.Set("twoDimensional", JsBool(env, t.two_dimensional));
  table.Set("rowOriented", JsBool(env, t.row_oriented));
  table.Set("rowInput", InputToJs(env, t.has_row_input, t.row_input_row, t.row_input_col, t.row_input_deleted));
  table.Set("colInput", InputToJs(env, t.has_col_input, t.col_input_row, t.col_input_col, t.col_input_deleted));
  out.Set("dataTable", table);
  return out;
}

Napi::Value Workbook::RemoveDataTable(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_data_table);
}

Napi::Value Workbook::RecalcDataTables(const Napi::CallbackInfo& info) {
  return StatusCall(info, [](fm_workbook_t* handle) { return fm_workbook_recalc_data_tables(handle); });
}

}  // namespace formulon_node
