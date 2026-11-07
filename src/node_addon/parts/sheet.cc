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
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    const std::string name = ArgString(info, 0);
    return fm_workbook_add_sheet(handle, name.c_str(), out);
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
      Napi::TypeError::New(env, "positional argument must be a string").ThrowAsJavaScriptException();
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

Napi::Object DefinedNameToJs(Napi::Env env, const char* name, const char* formula, int32_t local_sheet_id,
                             const char* comment, int32_t hidden) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("name", JsString(env, name));
  item.Set("formula", JsString(env, formula));
  item.Set("localSheetId", JsIndexOrNull(env, local_sheet_id));
  item.Set("comment", JsStringOrNull(env, comment));
  item.Set("hidden", Napi::Boolean::New(env, hidden != 0));
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
      arr.Set(static_cast<uint32_t>(i), DefinedNameToJs(env, d.name, d.formula, d.local_sheet_id, d.comment, d.hidden));
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
  const char* stored = nullptr;
  const char* formula = nullptr;
  const char* comment = nullptr;
  int32_t hidden = 0;
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_defined_name(handle_, name.c_str(), local_sheet_id, &present, &stored,
                                                      &formula, &comment, &hidden);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  return OptionalOrNull(env, rc, present, DefinedNameToJs(env, stored, formula, local_sheet_id, comment, hidden));
}

Napi::Value Workbook::RemoveDefinedName(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 0);
    return fm_workbook_remove_defined_name(handle, name.c_str(), ArgI32(info, 1));
  });
}

Napi::Value Workbook::TableRuntimeId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_get_table_runtime_id(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &out);
  return ValueOrThrow(env, rc, JsNumber(env, out));
}

Napi::Value Workbook::ResolveTableRuntimeId(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t out = 0;
  const fm_status_t rc = fm_workbook_resolve_table_runtime_id(handle_, ArgU32(info, 0), &out);
  return ValueOrThrow(env, rc, JsNumber(env, static_cast<double>(out)));
}

Napi::Value Workbook::TableAt(const Napi::CallbackInfo& info) {
  const char* name = nullptr;
  const char* display = nullptr;
  fm_cell_rect_t rect{};
  std::size_t sheet_index = 0;
  fm_table_info_t table = FM_STRUCT_INIT(fm_table_info_t);
  std::vector<fm_table_column_info_t> columns;
  // A respelled totals formula lives in read scratch, which the next column's read may recycle.
  std::vector<std::string> totals_formulas;
  return StatusObject(
      info,
      [&](fm_workbook_t* handle) {
        const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
        fm_status_t rc = fm_workbook_get_table(handle, idx, &name, &display, &rect, &sheet_index);
        if (rc == 0) {
          rc = fm_workbook_get_table_info(handle, idx, &table);
        }
        columns.resize(rc == 0 ? table.column_count : 0U, FM_STRUCT_INIT(fm_table_column_info_t));
        totals_formulas.resize(columns.size());
        for (std::size_t i = 0; rc == 0 && i < columns.size(); ++i) {
          rc = fm_workbook_get_table_column(handle, idx, i, &columns[i]);
          if (rc == 0) {
            totals_formulas[i] = columns[i].totals_row_formula;
          }
        }
        return rc;
      },
      [&](Napi::Object& out) {
        Napi::Env env = info.Env();
        out.Set("name", JsString(env, name));
        out.Set("displayName", JsString(env, display));
        Napi::Object range = Napi::Object::New(env);
        range.Set("sheet", Napi::Number::New(env, static_cast<double>(sheet_index)));
        range.Set("rect", RangeToJs(env, rect));
        out.Set("range", range);
        out.Set("styleName", JsStringOrNull(env, table.style_name));
        out.Set("showFirstColumn", JsBool(env, table.show_first_column));
        out.Set("showLastColumn", JsBool(env, table.show_last_column));
        out.Set("showRowStripes", JsBool(env, table.show_row_stripes));
        out.Set("showColumnStripes", JsBool(env, table.show_column_stripes));
        out.Set("headerRow", JsBool(env, table.header_row_count != 0));
        out.Set("totalsRow", JsBool(env, table.totals_row_count != 0));
        Napi::Array list = Napi::Array::New(env, columns.size());
        for (std::size_t i = 0; i < columns.size(); ++i) {
          Napi::Object c = Napi::Object::New(env);
          c.Set("name", JsString(env, columns[i].name));
          c.Set("totalsRowFunction", JsString(env, columns[i].totals_row_function));
          c.Set("totalsRowLabel", JsString(env, columns[i].totals_row_label));
          c.Set("totalsRowFormula", Napi::String::New(env, totals_formulas[i]));
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
  fm_table_info_t current = FM_STRUCT_INIT(fm_table_info_t);
  if (const fm_status_t rc = fm_workbook_get_table_info(handle, idx, &current); rc != 0) {
    return rc;
  }
  const bool present = current.style_present != 0;
  const std::string name = current.style_name != nullptr ? current.style_name : "";
  const auto pick = [](int32_t supplied, int32_t existing) { return supplied >= 0 ? supplied : existing; };
  return fm_workbook_set_table_style_info(handle, idx, name.c_str(),
                                          pick(toggles.first_column, current.show_first_column),
                                          pick(toggles.last_column, current.show_last_column),
                                          pick(toggles.row_stripes, present ? current.show_row_stripes : 1),
                                          pick(toggles.column_stripes, current.show_column_stripes));
}

}  // namespace

Napi::Value Workbook::CreateTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Object spec = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  uint32_t sheet = 0U;
  fm_cell_rect_t rect{};
  Napi::Object range_object;
  if (reader.RequiredObject(spec, "range", &range_object)) {
    sheet = reader.RequiredU32(range_object, "sheet");
    Napi::Object rect_object;
    if (reader.RequiredObject(range_object, "rect", &rect_object)) {
      ReadCellRect(reader, rect_object, &rect);
    }
  }
  std::string name;
  std::string display_name;
  std::string style_name;
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
    Napi::TypeError::New(env, "addTable: `columns` must be an array of column names").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  // `names` has stopped growing, so its `c_str()` views stay put.
  std::vector<const char*> pointers;
  pointers.reserve(names.size());
  for (const std::string& column : names) {
    pointers.push_back(column.c_str());
  }
  std::size_t index = 0;
  fm_status_t rc =
      fm_workbook_add_table(handle_, sheet, &rect, name.c_str(), display_name.c_str(), pointers.data(), pointers.size(),
                            style_name.c_str(), header_row ? 1 : 0, totals_row ? 1 : 0, &index);
  if (rc == 0 && toggles.any()) {
    rc = ApplyTableStyleToggles(handle_, index, toggles);
    if (rc != 0) {
      fm_workbook_remove_table(handle_, index);
    }
  }
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<uint32_t>(index)));
}

Napi::Value Workbook::UpdateTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const Napi::Object spec = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  Napi::Object rect_object;
  const bool has_rect = reader.Object(spec, "rect", &rect_object);
  if (has_rect) {
    ReadCellRect(reader, rect_object, &rect);
  }
  std::string style_name;
  std::string name;
  std::string display_name;
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
  fm_cell_rect_t current_rect{};
  std::size_t ignored_sheet = 0;
  if (const fm_status_t rc =
          fm_workbook_get_table(handle_, idx, &current_name, &current_display_name, &current_rect, &ignored_sheet);
      rc != 0) {
    return ValueOrThrow(env, rc, env.Undefined());
  }
  const std::string old_name = current_name != nullptr ? current_name : "";
  const std::string old_display_name = current_display_name != nullptr ? current_display_name : "";
  if (!has_rect) {
    rect = current_rect;
  }

  // The rename goes first because it is the step a name clash can refuse;
  // a later failure renames back so the update stays all-or-nothing.
  const bool renaming = has_name || has_display_name;
  if (renaming) {
    const std::string next_name = has_name ? name : old_name;
    const std::string next_display_name = has_display_name ? display_name : (has_name ? name : old_display_name);
    if (const fm_status_t rc = fm_workbook_rename_table(handle_, idx, next_name.c_str(), next_display_name.c_str());
        rc != 0) {
      return ValueOrThrow(env, rc, env.Undefined());
    }
  }
  fm_status_t rc =
      fm_workbook_update_table(handle_, idx, &rect, has_style ? style_name.c_str() : nullptr, header_row, totals_row);
  if (rc == 0 && toggles.any()) {
    rc = ApplyTableStyleToggles(handle_, idx, toggles);
  }
  if (rc != 0 && renaming) {
    fm_workbook_rename_table(handle_, idx, old_name.c_str(), old_display_name.c_str());
  }
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::RemoveTable(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(env, fm_workbook_remove_table(handle_, static_cast<std::size_t>(ArgU32(info, 0))),
                      env.Undefined());
}

Napi::Value Workbook::SetTableColumnTotals(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t column = static_cast<std::size_t>(ArgU32(info, 1));
  const Napi::Object totals = ArgObject(info, 2);
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
  const fm_status_t rc =
      fm_workbook_set_table_column_totals(handle_, idx, column, has_function ? function.c_str() : nullptr,
                                          has_label ? label.c_str() : nullptr, has_formula ? formula.c_str() : nullptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::RenameTableColumn(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t column = static_cast<std::size_t>(ArgU32(info, 1));
  const std::string name = ArgString(info, 2);
  const char* final_name = nullptr;
  const fm_status_t rc = fm_workbook_rename_table_column(handle_, idx, column, name.c_str(), &final_name);
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

Napi::Value Workbook::GetPassthrough(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string path = ArgString(info, 0);
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_passthrough(handle_, path.c_str(), &present);
  return OptionalOrNull(env, rc, present, JsString(env, path.c_str()));
}

Napi::Value Workbook::PivotLayout(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }

  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t pivot_index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_pivot_cells_t* cells = nullptr;
  fm_status_t rc = fm_sheet_get_pivot_layout(handle_, sheet, pivot_index, &cells);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }

  int32_t present = 0;
  fm_cell_rect_t bounds{};
  rc = fm_pivot_cells_get_bounds(cells, &present, &bounds);
  if (rc != 0) {
    fm_pivot_cells_destroy(cells);
    return ThrowStatus(env, rc);
  }

  std::size_t count = 0;
  rc = fm_pivot_cells_get_cell_count(cells, &count);
  if (rc != 0) {
    fm_pivot_cells_destroy(cells);
    return ThrowStatus(env, rc);
  }
  // Pre-sized `Array::New(env, count)` would leave trailing `undefined`
  // holes whenever `fm_pivot_cells_get_cell_at` skips an index below, so grow
  // the array dynamically and track a separate `emitted` counter —
  // matching the dense-array contract the other list getters use.
  Napi::Array arr = Napi::Array::New(env);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_pivot_cell_t cell = FM_STRUCT_INIT(fm_pivot_cell_t);
    if (fm_pivot_cells_get_cell_at(cells, i, &cell) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(emitted), TranslatePivotCell(env, cell));
    ++emitted;
  }
  fm_pivot_cells_destroy(cells);

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
  const uint32_t sheet = reader.U32(input, "localSheetId", 0U, &scoped);
  if (reader.ok() && scoped && sheet > static_cast<uint32_t>(std::numeric_limits<int32_t>::max())) {
    reader.FailRange("localSheetId", "int32");
  }
  std::string comment;
  const bool has_comment = reader.String(input, "comment", &comment);
  const bool hidden = reader.Bool(input, "hidden", false);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_defined_name_input_t in = FM_STRUCT_INIT(fm_defined_name_input_t);
  in.name = name.c_str();
  in.formula = formula.c_str();
  in.local_sheet_id = scoped ? static_cast<int32_t>(sheet) : -1;
  in.comment = has_comment ? comment.c_str() : nullptr;
  in.hidden = hidden ? 1 : 0;
  return ValueOrThrow(env, fm_workbook_set_defined_name(handle_, &in), env.Undefined());
}

}  // namespace formulon_node
