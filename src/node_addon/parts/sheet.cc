// Copyright 2026 libraz. Licensed under the MIT License.
//
// Sheet-level bindings: sheet add/remove/rename/move, row/column
// structural edits, metadata iteration (cells, defined names, tables,
// passthroughs, pivots) and defined-name mutation.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Sheet operations -----------------------------------------------

Napi::Value Workbook::AddSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  fm_status_t rc = fm_workbook_add_sheet(handle_, name.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t idx = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_remove_sheet(handle_, idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RenameSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t idx = ArgU32(info, 0);
  const std::string name = ArgString(info, 1);
  fm_status_t rc = fm_workbook_rename_sheet(handle_, idx, name.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::MoveSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t from_idx = ArgU32(info, 0);
  const uint32_t to_idx = ArgU32(info, 1);
  fm_status_t rc = fm_workbook_move_sheet(handle_, from_idx, to_idx);
  return MakeStatus(env, rc);
}

// `Workbook::SheetCount` is now emitted by the binding codegen (see
// `src/node_addon/generated/workbook_counts.cc`).

Napi::Value Workbook::SheetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_sheet_name(handle_, idx, &name);
  if (rc != 0) {
    return MakeStringFieldResult(env, MakeErrorStatus(env, rc), "value", "");
  }
  return MakeStringFieldResult(env, MakeOkStatus(env), "value", name);
}

// ---- Row / column structural edits ----------------------------------

Napi::Value Workbook::InsertRows(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_workbook_insert_rows);
}

Napi::Value Workbook::DeleteRows(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_workbook_delete_rows);
}

Napi::Value Workbook::InsertCols(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_workbook_insert_cols);
}

Napi::Value Workbook::DeleteCols(const Napi::CallbackInfo& info) {
  return InvokeRowColEdit(info, fm_workbook_delete_cols);
}

// ---- Iteration / metadata accessors ---------------------------------
//
// `CellCount` / `DefinedNameCount` / `TableCount` / `PassthroughCount`
// / `PivotCount` are now emitted by the binding codegen (see
// `src/node_addon/generated/workbook_counts.cc` and
// `src/node_addon/generated/sheet_counts.cc`).

Napi::Value Workbook::CellAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
  uint32_t row = 0;
  uint32_t col = 0;
  const char* formula = nullptr;
  fm_value_t v{};
  fm_status_t rc = fm_workbook_cell_at(handle_, sheet, idx, &row, &col, &formula, &v);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("row", Napi::Number::New(env, row));
  out.Set("col", Napi::Number::New(env, col));
  if (formula != nullptr) {
    out.Set("formula", Napi::String::New(env, formula));
  } else {
    out.Set("formula", env.Null());
  }
  out.Set("value", TranslateValue(env, v));
  return out;
}

Napi::Value Workbook::DefinedNameAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  const char* formula = nullptr;
  int32_t local_sheet_id = -1;
  fm_status_t rc = fm_workbook_defined_name_at_ex(handle_, idx, &name, &formula, &local_sheet_id);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, name != nullptr ? name : ""));
  out.Set("formula", Napi::String::New(env, formula != nullptr ? formula : ""));
  out.Set("localSheetId", Napi::Number::New(env, local_sheet_id));
  return out;
}

Napi::Value Workbook::TableAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  const char* display = nullptr;
  const char* ref = nullptr;
  std::size_t sheet_index = 0;
  fm_status_t rc = fm_workbook_table_at(handle_, idx, &name, &display, &ref, &sheet_index);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, name != nullptr ? name : ""));
  out.Set("displayName", Napi::String::New(env, display != nullptr ? display : ""));
  out.Set("ref", Napi::String::New(env, ref != nullptr ? ref : ""));
  out.Set("sheetIndex", Napi::Number::New(env, static_cast<double>(sheet_index)));
  return out;
}

Napi::Value Workbook::PassthroughAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* path = nullptr;
  fm_status_t rc = fm_workbook_passthrough_at(handle_, idx, &path);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("path", Napi::String::New(env, path != nullptr ? path : ""));
  return out;
}

Napi::Value Workbook::PivotLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return EmptyPivotLayoutResult(env, NullHandleError(env));
  }

  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_cells_t* cells = nullptr;
  fm_status_t rc = fm_workbook_pivot_layout(handle_, sheet, pivot_index, &cells);
  if (rc != 0) {
    return EmptyPivotLayoutResult(env, MakeErrorStatus(env, rc));
  }

  uint32_t top = 0;
  uint32_t left = 0;
  uint32_t rows = 0;
  uint32_t cols = 0;
  rc = fm_pivot_cells_bounds(cells, &top, &left, &rows, &cols);
  if (rc != 0) {
    fm_pivot_cells_destroy(cells);
    return EmptyPivotLayoutResult(env, MakeErrorStatus(env, rc));
  }

  const std::size_t count = fm_pivot_cells_count(cells);
  Napi::Array arr = Napi::Array::New(env, count);
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_t cell{};
    if (fm_pivot_cells_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(i), TranslatePivotCell(env, cell));
  }
  fm_pivot_cells_destroy(cells);

  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeOkStatus(env));
  out.Set("top", Napi::Number::New(env, top));
  out.Set("left", Napi::Number::New(env, left));
  out.Set("rows", Napi::Number::New(env, rows));
  out.Set("cols", Napi::Number::New(env, cols));
  out.Set("cells", arr);
  return out;
}

// ---- Defined names --------------------------------------------------

Napi::Value Workbook::SetDefinedName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  const std::string formula = ArgString(info, 1);
  fm_status_t rc = fm_workbook_set_defined_name(handle_, name.c_str(), formula.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetDefinedNameScoped(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  const std::string formula = ArgString(info, 1);
  const int32_t local_sheet_id = info.Length() > 2 ? info[2].As<Napi::Number>().Int32Value() : -1;
  fm_status_t rc = fm_workbook_set_defined_name_scoped(handle_, name.c_str(), formula.c_str(), local_sheet_id);
  return MakeStatus(env, rc);
}

}  // namespace formulon_node
