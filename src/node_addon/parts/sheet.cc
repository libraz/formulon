// Sheet-level bindings: sheet add/remove/rename/move, row/column
// structural edits, metadata iteration (cells, defined names, tables,
// passthroughs, pivots) and defined-name mutation.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Sheet operations -----------------------------------------------

Napi::Value Workbook::AddSheet(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 0);
    return fm_workbook_add_sheet(handle, name.c_str());
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

// `Workbook::SheetCount` is now emitted by the binding codegen (see
// `src/node_addon/generated/workbook_counts.cc`).

Napi::Value Workbook::ActiveSheet(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_get_active_sheet(handle_, &out);
  return MakeIndexResult(env, rc, out);
}

Napi::Value Workbook::SetActiveSheet(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_set_active_sheet);
}

Napi::Value Workbook::SheetName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_sheet_name(handle_, idx, &name);
  return MakeStringResult(env, rc, name);
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
  uint32_t row = 0;
  uint32_t col = 0;
  const char* formula = nullptr;
  fm_value_t v{};
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 1));
        return fm_workbook_cell_at(handle, sheet, idx, &row, &col, &formula, &v);
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
        out.Set("value", TranslateValue(env, v));
      });
}

Napi::Value Workbook::DefinedNameAt(const Napi::CallbackInfo& info) {
  const char* name = nullptr;
  const char* formula = nullptr;
  int32_t local_sheet_id = -1;
  const char* comment = "";
  int32_t hidden = 0;
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
        const fm_status_t rc = fm_workbook_defined_name_at(handle, idx, &name, &formula, &local_sheet_id);
        return rc != 0 ? rc : fm_workbook_defined_name_attrs_at(handle, idx, &comment, &hidden);
      },
      [&](Napi::Object& out) {
        Napi::Env env = info.Env();
        out.Set("name", JsString(env, name));
        out.Set("formula", JsString(env, formula));
        out.Set("localSheetId", Napi::Number::New(env, local_sheet_id));
        out.Set("comment", JsString(env, comment));
        out.Set("hidden", Napi::Boolean::New(env, hidden != 0));
      });
}

Napi::Value Workbook::TableAt(const Napi::CallbackInfo& info) {
  const char* name = nullptr;
  const char* display = nullptr;
  const char* ref = nullptr;
  std::size_t sheet_index = 0;
  fm_table_info_t table{};
  std::vector<fm_table_column_info_t> columns;
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
        fm_status_t rc = fm_workbook_table_at(handle, idx, &name, &display, &ref, &sheet_index);
        if (rc == 0) {
          rc = fm_workbook_table_get_info(handle, idx, &table);
        }
        columns.resize(rc == 0 ? table.column_count : 0U);
        for (std::size_t i = 0; rc == 0 && i < columns.size(); ++i) {
          rc = fm_workbook_table_get_column(handle, idx, i, &columns[i]);
        }
        return rc;
      },
      [&](Napi::Object& out) {
        Napi::Env env = info.Env();
        out.Set("name", JsString(env, name));
        out.Set("displayName", JsString(env, display));
        out.Set("ref", JsString(env, ref));
        out.Set("sheetIndex", Napi::Number::New(env, static_cast<double>(sheet_index)));
        out.Set("styleName", JsString(env, table.style_name));
        out.Set("hasStyleInfo", JsBool(env, table.style_present));
        out.Set("showFirstColumn", JsBool(env, table.show_first_column));
        out.Set("showLastColumn", JsBool(env, table.show_last_column));
        out.Set("showRowStripes", JsBool(env, table.show_row_stripes));
        out.Set("showColumnStripes", JsBool(env, table.show_column_stripes));
        out.Set("headerRowCount", Napi::Number::New(env, table.header_row_count));
        out.Set("totalsRowCount", Napi::Number::New(env, table.totals_row_count));
        Napi::Array list = Napi::Array::New(env, columns.size());
        for (std::size_t i = 0; i < columns.size(); ++i) {
          Napi::Object c = Napi::Object::New(env);
          c.Set("name", JsString(env, columns[i].name));
          c.Set("totalsRowFunction", JsString(env, columns[i].totals_row_function));
          c.Set("totalsRowLabel", JsString(env, columns[i].totals_row_label));
          c.Set("totalsRowFormula", JsString(env, columns[i].totals_row_formula));
          list.Set(static_cast<uint32_t>(i), c);
        }
        out.Set("columns", list);
      });
}

namespace {

/// The four `<tableStyleInfo>` toggles of a table input, each `-1` when the
/// input omits it.
struct TableStyleToggles {
  int32_t first_column = -1;
  int32_t last_column = -1;
  int32_t row_stripes = -1;
  int32_t column_stripes = -1;

  bool any() const { return first_column >= 0 || last_column >= 0 || row_stripes >= 0 || column_stripes >= 0; }
};

int32_t OptionalBool(CheckedSpecReader& reader, const Napi::Object& spec, const char* key) {
  bool present = false;
  const bool value = reader.Bool(spec, key, false, &present);
  return present ? (value ? 1 : 0) : -1;
}

TableStyleToggles ReadTableStyleToggles(CheckedSpecReader& reader, const Napi::Object& spec) {
  TableStyleToggles toggles;
  toggles.first_column = OptionalBool(reader, spec, "showFirstColumn");
  toggles.last_column = OptionalBool(reader, spec, "showLastColumn");
  toggles.row_stripes = OptionalBool(reader, spec, "showRowStripes");
  toggles.column_stripes = OptionalBool(reader, spec, "showColumnStripes");
  return toggles;
}

/// Applies the supplied toggles over the table's current ones, keeping its
/// style name; a table without the element takes the creation defaults.
fm_status_t ApplyTableStyleToggles(fm_workbook_t* handle, std::size_t idx, const TableStyleToggles& toggles) {
  fm_table_info_t current{};
  if (const fm_status_t rc = fm_workbook_table_get_info(handle, idx, &current); rc != 0) {
    return rc;
  }
  const bool present = current.style_present != 0;
  const std::string name = current.style_name != nullptr ? current.style_name : "";
  const auto pick = [](int32_t supplied, int32_t existing) { return supplied >= 0 ? supplied : existing; };
  return fm_workbook_table_set_style_info(handle, idx, name.c_str(),
                                          pick(toggles.first_column, current.show_first_column),
                                          pick(toggles.last_column, current.show_last_column),
                                          pick(toggles.row_stripes, present ? current.show_row_stripes : 1),
                                          pick(toggles.column_stripes, current.show_column_stripes));
}

}  // namespace

Napi::Value Workbook::CreateTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0U);
  }
  if (info.Length() < 1 || !info[0].IsObject()) {
    return MakeIndexResult(env, MakeBindingArgumentError(env, "createTable: `input` must be an object"), 0U);
  }
  const Napi::Object spec = info[0].As<Napi::Object>();
  CheckedSpecReader reader(env);
  const uint32_t sheet = reader.U32(spec, "sheetIndex", 0U);
  std::string ref;
  std::string name;
  std::string display_name;
  std::string style_name;
  reader.String(spec, "ref", &ref);
  reader.String(spec, "name", &name);
  reader.String(spec, "displayName", &display_name);
  if (display_name.empty()) {
    display_name = name;
  }
  reader.String(spec, "styleName", &style_name);
  const bool header_row = reader.Bool(spec, "headerRow", true);
  const bool totals_row = reader.Bool(spec, "totalsRow", false);
  const TableStyleToggles toggles = ReadTableStyleToggles(reader, spec);
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
    return MakeIndexResult(
        env, MakeBindingArgumentError(env, "createTable: `columns` must be an array of column names"), 0U);
  }
  // `names` has stopped growing, so its `c_str()` views stay put.
  std::vector<const char*> pointers;
  pointers.reserve(names.size());
  for (const std::string& column : names) {
    pointers.push_back(column.c_str());
  }
  std::size_t index = 0;
  fm_status_t rc =
      fm_workbook_table_create(handle_, sheet, ref.c_str(), name.c_str(), display_name.c_str(), pointers.data(),
                               pointers.size(), style_name.c_str(), header_row ? 1 : 0, totals_row ? 1 : 0, &index);
  if (rc == 0 && toggles.any()) {
    rc = ApplyTableStyleToggles(handle_, index, toggles);
    if (rc != 0) {
      fm_workbook_table_remove(handle_, index);
    }
  }
  return MakeIndexResult(env, rc, static_cast<uint32_t>(index));
}

Napi::Value Workbook::UpdateTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const Napi::Object spec = ArgObjectOrEmpty(info, 1);
  CheckedSpecReader reader(env);
  std::string ref;
  std::string style_name;
  std::string name;
  std::string display_name;
  const bool has_ref = reader.String(spec, "ref", &ref);
  const bool has_style = reader.String(spec, "styleName", &style_name);
  const bool has_name = reader.String(spec, "name", &name);
  const bool has_display_name = reader.String(spec, "displayName", &display_name);
  const int32_t header_row = OptionalBool(reader, spec, "headerRow");
  const int32_t totals_row = OptionalBool(reader, spec, "totalsRow");
  const TableStyleToggles toggles = ReadTableStyleToggles(reader, spec);
  if (!reader.ok()) {
    return env.Undefined();
  }

  const char* current_name = nullptr;
  const char* current_display_name = nullptr;
  const char* current_ref = nullptr;
  std::size_t ignored_sheet = 0;
  if (const fm_status_t rc =
          fm_workbook_table_at(handle_, idx, &current_name, &current_display_name, &current_ref, &ignored_sheet);
      rc != 0) {
    return MakeStatus(env, rc);
  }
  const std::string old_name = current_name != nullptr ? current_name : "";
  const std::string old_display_name = current_display_name != nullptr ? current_display_name : "";
  if (!has_ref) {
    ref = current_ref != nullptr ? current_ref : "";
  }

  // The rename goes first because it is the step a name clash can refuse;
  // a later failure renames back so the update stays all-or-nothing.
  const bool renaming = has_name || has_display_name;
  if (renaming) {
    const std::string next_name = has_name ? name : old_name;
    const std::string next_display_name = has_display_name ? display_name : (has_name ? name : old_display_name);
    if (const fm_status_t rc = fm_workbook_table_rename(handle_, idx, next_name.c_str(), next_display_name.c_str());
        rc != 0) {
      return MakeStatus(env, rc);
    }
  }
  fm_status_t rc = fm_workbook_table_update(handle_, idx, ref.c_str(), has_style ? style_name.c_str() : nullptr,
                                            header_row, totals_row);
  if (rc == 0 && toggles.any()) {
    rc = ApplyTableStyleToggles(handle_, idx, toggles);
  }
  if (rc != 0 && renaming) {
    fm_workbook_table_rename(handle_, idx, old_name.c_str(), old_display_name.c_str());
  }
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_workbook_table_remove(handle_, static_cast<std::size_t>(ArgU32(info, 0))));
}

Napi::Value Workbook::SetTableColumnTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t column = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object totals = ArgObjectOrEmpty(info, 2);
  CheckedSpecReader reader(env);
  std::string function;
  std::string label;
  std::string formula;
  const bool has_function = reader.String(totals, "function", &function);
  const bool has_label = reader.String(totals, "label", &label);
  const bool has_formula = reader.String(totals, "formula", &formula);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_workbook_table_set_column_totals(
                             handle_, idx, column, has_function ? function.c_str() : nullptr,
                             has_label ? label.c_str() : nullptr, has_formula ? formula.c_str() : nullptr));
}

Napi::Value Workbook::PassthroughAt(const Napi::CallbackInfo& info) {
  const char* path = nullptr;
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
        return fm_workbook_passthrough_at(handle, idx, &path);
      },
      [&](Napi::Object& out) { out.Set("path", JsString(info.Env(), path)); });
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
  // Pre-sized `Array::New(env, count)` would leave trailing `undefined`
  // holes whenever `fm_pivot_cells_at` skips an index below, so grow
  // the array dynamically and track a separate `emitted` counter —
  // matching the dense-array contract the other list getters use.
  Napi::Array arr = Napi::Array::New(env);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_t cell{};
    if (fm_pivot_cells_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(emitted), TranslatePivotCell(env, cell));
    ++emitted;
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
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 0);
    const std::string formula = ArgString(info, 1);
    return fm_workbook_set_defined_name(handle, name.c_str(), formula.c_str());
  });
}

Napi::Value Workbook::SetDefinedNameScoped(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 0);
    const std::string formula = ArgString(info, 1);
    const int32_t local_sheet_id = info.Length() > 2 ? info[2].As<Napi::Number>().Int32Value() : -1;
    return fm_workbook_set_defined_name_scoped(handle, name.c_str(), formula.c_str(), local_sheet_id);
  });
}

Napi::Value Workbook::SetDefinedNameAttrs(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  const int32_t local_sheet_id = ArgI32(info, 1);
  Napi::Object attrs = ArgObjectOrEmpty(info, 2);
  CheckedSpecReader reader(env);
  std::string comment;
  bool has_hidden = false;
  const bool has_comment = reader.String(attrs, "comment", &comment);
  const bool hidden = reader.Bool(attrs, "hidden", false, &has_hidden);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_workbook_set_defined_name_attrs(handle_, name.c_str(), local_sheet_id,
                                                            has_comment ? comment.c_str() : nullptr,
                                                            has_hidden ? (hidden ? 1 : 0) : -1));
}

}  // namespace formulon_node
