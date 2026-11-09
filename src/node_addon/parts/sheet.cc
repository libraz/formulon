// Sheet-level bindings: sheet add/remove/rename/move, row/column
// structural edits, metadata iteration (cells, defined names, tables,
// passthroughs, pivots) and defined-name mutation.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Sheet operations -----------------------------------------------

Napi::Value Workbook::AddSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  // A null name takes the next default sheet name.
  std::string name;
  const bool has_name = !info[0].IsUndefined() && !info[0].IsNull();
  if (has_name) {
    if (!info[0].IsString()) {
      Napi::TypeError::New(env, "Workbook.addSheet name must be a string or null")
          .ThrowAsJavaScriptException();
      return env.Undefined();
    }
    name = ArgString(info, 0);
  }
  const std::size_t insert_before = ArgU32(info, 1);
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    return fm_workbook_add_sheet(handle, has_name ? name.c_str() : nullptr, insert_before, out);
  });
}

Napi::Value Workbook::RemoveSheet(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_remove_sheet);
}

Napi::Value Workbook::RenameSheet(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const uint32_t idx = ArgU32(info, 0);
    const std::string name = ArgString(info, 1);
    return fm_workbook_rename_sheet(handle, idx, name.c_str());
  });
}

Napi::Value Workbook::MoveSheet(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_move_sheet);
}

Napi::Value Workbook::CopySheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::string name;
  const bool has_name = info.Length() > 2 && !info[2].IsUndefined() && !info[2].IsNull();
  if (has_name) {
    if (!info[2].IsString()) {
      Napi::TypeError::New(env, "Workbook.copySheet newName must be a string").ThrowAsJavaScriptException();
      return env.Undefined();
    }
    name = ArgString(info, 2);
  }
  std::size_t out = 0;
  const fm_status_t rc =
      fm_workbook_copy_sheet(handle_, ArgU32(info, 0), ArgU32(info, 1), has_name ? name.c_str() : nullptr, &out);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<uint32_t>(out)));
}

// `Workbook::SheetCount` is now emitted by the binding codegen (see
// `src/node_addon/generated/workbook_counts.cc`).

Napi::Value Workbook::ActiveSheet(const Napi::CallbackInfo& info) {
  return IndexCall(info,
                   [](fm_workbook_t* handle, std::size_t* out) { return fm_workbook_get_active_sheet(handle, out); });
}

Napi::Value Workbook::SetActiveSheet(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_set_active_sheet);
}

Napi::Value Workbook::SheetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_get_sheet_name(handle_, idx, &name);
  return ValueOrThrow(env, rc, JsString(env, name));
}

// ---- Row / column structural edits ----------------------------------

Napi::Value Workbook::InsertRows(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_sheet_insert_rows);
}

Napi::Value Workbook::DeleteRows(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_sheet_delete_rows);
}

Napi::Value Workbook::InsertCols(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_sheet_insert_columns);
}

Napi::Value Workbook::DeleteCols(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_sheet_delete_columns);
}

namespace {

using CellShiftEditFn = fm_status_t (*)(fm_workbook_t*, std::size_t, const fm_cell_rect_t*, int32_t);

// `(sheet, range, shift)`: the guard checks the sheet index and the `CellShift` ordinal's type.
Napi::Value InvokeCellShiftEdit(const Napi::CallbackInfo& info, fm_workbook_t* const& handle, CellShiftEditFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t range{};
  if (!CellRectArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fn(handle, Workbook::ArgU32(info, 0), &range, Workbook::ArgI32(info, 2)), env.Undefined());
}

}  // namespace

Napi::Value Workbook::InsertCells(const Napi::CallbackInfo& info) {
  return InvokeCellShiftEdit(info, handle_, &fm_sheet_insert_cells);
}

Napi::Value Workbook::DeleteCells(const Napi::CallbackInfo& info) {
  return InvokeCellShiftEdit(info, handle_, &fm_sheet_delete_cells);
}

// ---- Iteration / metadata accessors ---------------------------------
//
// `CellCount` / `DefinedNameCount` / `TableCount` / `PassthroughCount`
// / `PivotCount` are now emitted by the binding codegen (see
// `src/node_addon/generated/workbook_counts.cc` and
// `src/node_addon/generated/sheet_counts.cc`).

Napi::Value Workbook::CellAt(const Napi::CallbackInfo& info) {
  uint32_t row = 0;
  uint32_t col = 0;
  const char* formula = nullptr;
  fm_value_t v{};
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
        return fm_sheet_get_cell_at(handle, sheet, idx, &row, &col, &formula, &v);
      },
      [&](Napi::Object& out) {
        Napi::Env env = info.Env();
        out.Set("row", Napi::Number::New(env, row));
        out.Set("col", Napi::Number::New(env, col));
        if (formula != nullptr) {
          out.Set("formula", Napi::String::New(env, formula));
        } else {
          out.Set("formula", env.Null());
        }
        out.Set("value", ValueToJs(env, v));
      });
}

namespace {

Napi::Object DefinedNameToJs(Napi::Env env, const fm_defined_name_t& d) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("name", JsString(env, d.name));
  item.Set("formula", JsString(env, d.formula));
  item.Set("scope", JsIndexOrNull(env, d.local_sheet_id));
  item.Set("comment", JsStringOrNull(env, d.comment));
  item.Set("hidden", Napi::Boolean::New(env, d.hidden != 0));
  return item;
}

}  // namespace

Napi::Value Workbook::ListDefinedNames(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t n = 0;
  fm_status_t rc = fm_workbook_get_defined_name_count(handle_, &n);
  for (std::size_t i = 0; rc == 0 && i < n; ++i) {
    fm_defined_name_t d = FM_STRUCT_INIT(fm_defined_name_t);
    rc = fm_workbook_get_defined_name_at(handle_, i, &d);
    if (rc == 0) {
      arr.Set(static_cast<uint32_t>(i), DefinedNameToJs(env, d));
    }
  }
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetDefinedName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  const int32_t local_sheet_id = ArgI32(info, 1);
  fm_defined_name_t found = FM_STRUCT_INIT(fm_defined_name_t);
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_defined_name(handle_, name.c_str(), local_sheet_id, &present, &found);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  return OptionalOrNull(env, rc, present, DefinedNameToJs(env, found));
}

Napi::Value Workbook::RemoveDefinedName(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 0);
    return fm_workbook_remove_defined_name(handle, name.c_str(), ArgI32(info, 1));
  });
}

Napi::Value Workbook::TableCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_sheet_get_table_count(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &count);
  return ValueOrThrow(env, rc, JsNumber(env, static_cast<double>(count)));
}

Napi::Value Workbook::TableRuntimeId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_sheet_get_table_runtime_id(handle_, static_cast<std::size_t>(ArgU32(info, 0)),
                                                       static_cast<std::size_t>(ArgU32(info, 1)), &out);
  return ValueOrThrow(env, rc, JsNumber(env, out));
}

Napi::Value Workbook::ResolveTableRuntimeId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t sheet_index = 0;
  std::size_t table_index = 0;
  const fm_status_t rc = fm_workbook_resolve_table_runtime_id(handle_, ArgU32(info, 0), &sheet_index, &table_index);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("sheet", JsNumber(env, static_cast<double>(sheet_index)));
  out.Set("index", JsNumber(env, static_cast<double>(table_index)));
  return out;
}

Napi::Value Workbook::FindTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string name = ArgString(info, 0);
  std::size_t sheet_index = 0;
  std::size_t table_index = 0;
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_find_table(handle_, name.c_str(), &present, &sheet_index, &table_index);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  if (present == 0) {
    return env.Null();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("sheet", JsNumber(env, static_cast<double>(sheet_index)));
  out.Set("index", JsNumber(env, static_cast<double>(table_index)));
  return out;
}

Napi::Value Workbook::TableAt(const Napi::CallbackInfo& info) {
  fm_table_info_t table = FM_STRUCT_INIT(fm_table_info_t);
  std::vector<fm_table_column_info_t> columns;
  // A respelled totals formula lives in read scratch, which the next column's read may recycle.
  std::vector<std::string> totals_formulas;
  std::vector<bool> has_totals_formula;
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
        fm_status_t rc = fm_sheet_get_table_at(handle, sheet, idx, &table);
        columns.resize(rc == 0 ? table.column_count : 0U, FM_STRUCT_INIT(fm_table_column_info_t));
        totals_formulas.resize(columns.size());
        has_totals_formula.assign(columns.size(), false);
        for (std::size_t i = 0; rc == 0 && i < columns.size(); ++i) {
          rc = fm_sheet_get_table_column_at(handle, sheet, idx, i, &columns[i]);
          if (rc == 0) {
            has_totals_formula[i] = columns[i].totals_row_formula != nullptr;
            if (has_totals_formula[i]) {
              totals_formulas[i] = columns[i].totals_row_formula;
            }
          }
        }
        return rc;
      },
      [&](Napi::Object& out) {
        Napi::Env env = info.Env();
        out.Set("name", JsString(env, table.name));
        Napi::Object range = Napi::Object::New(env);
        range.Set("sheet", Napi::Number::New(env, static_cast<double>(ArgU32(info, 0))));
        range.Set("rect", RangeToJs(env, table.range));
        out.Set("range", range);
        out.Set("headerRow", JsBool(env, table.header_row_count != 0));
        out.Set("totalsRow", JsBool(env, table.totals_row_count != 0));
        Napi::Array list = Napi::Array::New(env, columns.size());
        for (std::size_t i = 0; i < columns.size(); ++i) {
          Napi::Object c = Napi::Object::New(env);
          c.Set("name", JsString(env, columns[i].name));
          c.Set("totalsRowFunction", JsNumber(env, columns[i].totals_row_function));
          c.Set("totalsRowLabel",
                columns[i].totals_row_label != nullptr ? JsString(env, columns[i].totals_row_label) : env.Null());
          c.Set("totalsRowFormula", has_totals_formula[i] ? Napi::String::New(env, totals_formulas[i]) : env.Null());
          list.Set(static_cast<uint32_t>(i), c);
        }
        out.Set("columns", list);
      });
}

Napi::Value Workbook::GetTableStyleOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_table_style_options_t style = FM_STRUCT_INIT(fm_table_style_options_t);
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_table_style_options(handle_, static_cast<std::size_t>(ArgU32(info, 0)),
                                                          static_cast<std::size_t>(ArgU32(info, 1)), &present, &style);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, 0, env.Null());
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("styleName", JsStringOrNull(env, style.style_name));
  out.Set("showFirstColumn", JsBool(env, style.show_first_column));
  out.Set("showLastColumn", JsBool(env, style.show_last_column));
  out.Set("showRowStripes", JsBool(env, style.show_row_stripes));
  out.Set("showColumnStripes", JsBool(env, style.show_column_stripes));
  return out;
}

Napi::Value Workbook::CreateTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const Napi::Object spec = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  Napi::Object rect_object;
  if (reader.RequiredObject(spec, "range", &rect_object)) {
    ReadCellRect(reader, rect_object, &rect);
  }
  std::string name;
  std::string style_name;
  reader.String(spec, "name", &name);
  // An omitted style takes the workbook default; `styled: false` creates no style element.
  const bool has_style = reader.String(spec, "styleName", &style_name);
  Napi::Array columns;
  const bool has_columns = reader.Array(spec, "columns", &columns);
  std::vector<std::string> names;
  for (uint32_t i = 0; has_columns && reader.ok() && i < columns.Length(); ++i) {
    Napi::Value column;
    reader.ArrayElement(columns, i, &column);
    std::string column_name;
    reader.String(column, "columns[]", &column_name);
    names.push_back(std::move(column_name));
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  if (!has_columns) {
    ThrowArgumentTypeError(env, "input.columns", "must be an array of column names");
    return env.Undefined();
  }
  // `names` has stopped growing, so its `c_str()` views stay put.
  std::vector<const char*> pointers;
  pointers.reserve(names.size());
  for (const std::string& column : names) {
    pointers.push_back(column.c_str());
  }
  fm_table_spec_t table = FM_STRUCT_INIT(fm_table_spec_t);
  table.range = rect;
  table.name = name.c_str();
  table.column_names = pointers.empty() ? nullptr : pointers.data();
  table.column_count = static_cast<uint32_t>(pointers.size());
  table.style_name = has_style ? style_name.c_str() : nullptr;
  table.styled = reader.Tristate(spec, "styled");
  table.header_row = reader.Tristate(spec, "headerRow");
  table.totals_row = reader.Tristate(spec, "totalsRow");
  table.show_first_column = reader.Tristate(spec, "showFirstColumn");
  table.show_last_column = reader.Tristate(spec, "showLastColumn");
  table.show_row_stripes = reader.Tristate(spec, "showRowStripes");
  table.show_column_stripes = reader.Tristate(spec, "showColumnStripes");
  if (!reader.ok()) {
    return env.Undefined();
  }
  std::size_t index = 0;
  const fm_status_t rc = fm_sheet_add_table(handle_, sheet, &table, &index);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<uint32_t>(index)));
}

Napi::Value Workbook::UpdateTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object spec = ArgObject(info, 2);
  CheckedSpecReader reader(env);
  fm_table_patch_t patch = FM_STRUCT_INIT(fm_table_patch_t);
  Napi::Object rect_object;
  if (reader.Object(spec, "rect", &rect_object) && ReadCellRect(reader, rect_object, &patch.range)) {
    patch.range_op = FM_PATCH_OP_SET;
  }
  patch.header_row_op = PatchBool(reader, spec, "headerRow", &patch.header_row);
  patch.totals_row_op = PatchBool(reader, spec, "totalsRow", &patch.totals_row);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_table(handle_, sheet, idx, &patch), env.Undefined());
}

Napi::Value Workbook::RenameTable(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 2);
    return fm_sheet_rename_table(handle, static_cast<std::size_t>(ArgU32(info, 0)),
                                 static_cast<std::size_t>(ArgU32(info, 1)), name.c_str());
  });
}

Napi::Value Workbook::UpdateTableStyleOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object spec = ArgObject(info, 2);
  // An omitted key keeps the stored value; only `styleName` accepts `null`, which removes the style name.
  CheckedSpecReader reader(env);
  fm_table_style_options_patch_t patch = FM_STRUCT_INIT(fm_table_style_options_patch_t);
  std::string style_name;
  patch.style_name_op = PatchString(reader, spec, "styleName", &style_name, &patch.style_name);
  patch.show_first_column_op = PatchBool(reader, spec, "showFirstColumn", &patch.show_first_column);
  patch.show_last_column_op = PatchBool(reader, spec, "showLastColumn", &patch.show_last_column);
  patch.show_row_stripes_op = PatchBool(reader, spec, "showRowStripes", &patch.show_row_stripes);
  patch.show_column_stripes_op = PatchBool(reader, spec, "showColumnStripes", &patch.show_column_stripes);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_table_style_options(handle_, sheet, idx, &patch), env.Undefined());
}

Napi::Value Workbook::RemoveTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(env,
                      fm_sheet_remove_table(handle_, static_cast<std::size_t>(ArgU32(info, 0)),
                                            static_cast<std::size_t>(ArgU32(info, 1))),
                      env.Undefined());
}

Napi::Value Workbook::SetTableColumnTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t column = static_cast<std::size_t>(ArgU32(info, 2));
  const Napi::Object totals = ArgObject(info, 3);
  CheckedSpecReader reader(env);
  std::string label;
  std::string formula;
  const int32_t function = reader.I32(totals, "totalsRowFunction", FM_TABLE_TOTALS_ROW_FUNCTION_NONE);
  const bool has_label = reader.String(totals, "totalsRowLabel", &label);
  const bool has_formula = reader.String(totals, "totalsRowFormula", &formula);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc =
      fm_sheet_set_table_column_totals(handle_, sheet, idx, column, function, has_label ? label.c_str() : nullptr,
                                       has_formula ? formula.c_str() : nullptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::RenameTableColumn(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t column = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string name = ArgString(info, 3);
  const char* final_name = nullptr;
  const fm_status_t rc = fm_sheet_rename_table_column(handle_, sheet, idx, column, name.c_str(), &final_name);
  return ValueOrThrow(env, rc, JsString(env, final_name));
}

Napi::Value Workbook::ListPassthroughs(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const fm_status_t rc = CollectList<const char*>(
      arr, [&](size_t* n) { return fm_workbook_get_passthrough_count(handle_, n); },
      [&](size_t i, const char** path) { return fm_workbook_get_passthrough_at(handle_, i, path); },
      [&](const char* path) { return JsString(env, path); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::HasPassthrough(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string path = ArgString(info, 0);
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_has_passthrough(handle_, path.c_str(), &present);
  return ValueOrThrow(env, rc, Napi::Boolean::New(env, present != 0));
}

Napi::Value Workbook::PivotLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }

  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_report_t* cells = nullptr;
  fm_status_t rc = fm_sheet_get_pivot_layout(handle_, sheet, pivot_index, &cells);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }

  int32_t present = 0;
  fm_cell_rect_t bounds{};
  rc = fm_pivot_report_get_bounds(cells, &present, &bounds);
  if (rc != 0) {
    fm_pivot_report_destroy(cells);
    return ThrowStatus(env, rc);
  }

  std::size_t count = 0;
  rc = fm_pivot_report_get_cell_count(cells, &count);
  if (rc != 0) {
    fm_pivot_report_destroy(cells);
    return ThrowStatus(env, rc);
  }
  // Pre-sized `Array::New(env, count)` would leave trailing `undefined`
  // holes whenever `fm_pivot_report_get_cell_at` skips an index below, so grow
  // the array dynamically and track a separate `emitted` counter —
  // matching the dense-array contract the other list getters use.
  Napi::Array arr = Napi::Array::New(env);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_info_t cell = FM_STRUCT_INIT(fm_pivot_cell_info_t);
    if (fm_pivot_report_get_cell_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(emitted), TranslatePivotCell(env, cell));
    ++emitted;
  }
  fm_pivot_report_destroy(cells);

  Napi::Object out = Napi::Object::New(env);
  out.Set("rect", present != 0 ? Napi::Value(RangeToJs(env, bounds)) : env.Null());
  out.Set("cells", arr);
  return out;
}

// ---- Defined names --------------------------------------------------

Napi::Value Workbook::SetDefinedName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // Creates or wholly replaces the name; an omitted or `null` scope is the workbook.
  Napi::Object input = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  const std::string name = reader.RequiredString(input, "name");
  const std::string formula = reader.RequiredString(input, "formula");
  bool scoped = false;
  const uint32_t sheet = reader.U32(input, "scope", 0U, &scoped);
  if (reader.ok() && scoped && sheet > static_cast<uint32_t>(std::numeric_limits<int32_t>::max())) {
    reader.FailInteger("scope", static_cast<double>(sheet), 0.0);
  }
  std::string comment;
  const bool has_comment = reader.String(input, "comment", &comment);
  const bool hidden = reader.Bool(input, "hidden", false);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_defined_name_t in = FM_STRUCT_INIT(fm_defined_name_t);
  in.name = name.c_str();
  in.formula = formula.c_str();
  in.local_sheet_id = scoped ? static_cast<int32_t>(sheet) : FM_DEFINED_NAME_SCOPE_WORKBOOK;
  in.comment = has_comment ? comment.c_str() : nullptr;
  in.hidden = hidden ? 1 : 0;
  return ValueOrThrow(env, fm_workbook_set_defined_name(handle_, &in), env.Undefined());
}

}  // namespace formulon_node
