// Sheet view / layout bindings: zoom, freeze, tab visibility, protection,
// and column / row layout.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Sheet view / layout --------------------------------------------

namespace {

// Reads a geometry mode ordinal; an absent argument is the display mode.
int32_t ArgMode(const Napi::CallbackInfo& info, size_t idx) {
  return info[idx].IsUndefined() ? static_cast<int32_t>(FM_GEOMETRY_MODE_DISPLAY) : Workbook::ArgI32(info, idx);
}

// Enumerates a per-sheet layout list into an array, skipping entries the C
// getter cannot read.
template <typename Entry>
Napi::Value SheetLayoutList(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                            fm_status_t (*count_fn)(const fm_workbook_t*, size_t, size_t*),
                            fm_status_t (*at_fn)(const fm_workbook_t*, size_t, size_t, Entry*),
                            Napi::Object (*to_js)(Napi::Env, const Entry&)) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = count_fn(handle, sheet, &count);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array arr = Napi::Array::New(env, count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    Entry entry = InitOut<Entry>();
    if (at_fn(handle, sheet, i, &entry) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(emitted), to_js(env, entry));
    ++emitted;
  }
  return arr;
}

Napi::Object ColumnLayoutToJs(Napi::Env env, const fm_column_layout_info_t& entry) {
  Napi::Object col = Napi::Object::New(env);
  col.Set("first", Napi::Number::New(env, entry.first));
  col.Set("last", Napi::Number::New(env, entry.last));
  col.Set("widthChars", JsNumberOrNull(env, entry.has_width_chars, entry.width_chars));
  col.Set("hidden", JsBool(env, entry.hidden));
  col.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
  col.Set("collapsed", JsBool(env, entry.collapsed));
  col.Set("xfIndex", JsNumberOrNull(env, entry.has_xf_index, entry.xf_index));
  return col;
}

Napi::Object RowLayoutToJs(Napi::Env env, const fm_row_layout_info_t& entry) {
  Napi::Object row = Napi::Object::New(env);
  row.Set("row", Napi::Number::New(env, entry.row));
  row.Set("heightPt", JsNumberOrNull(env, entry.has_height_pt, entry.height_pt));
  row.Set("hidden", JsBool(env, entry.hidden));
  row.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
  row.Set("collapsed", JsBool(env, entry.collapsed));
  row.Set("xfIndex", JsNumberOrNull(env, entry.has_xf_index, entry.xf_index));
  row.Set("customHeight", JsBool(env, entry.custom_height));
  return row;
}

// Shared body of the column width unit converters (chars <-> points).
using ColumnUnitFn = fm_status_t (*)(const fm_workbook_t*, size_t, double, int32_t, double*);

Napi::Value ConvertColumnUnit(const Napi::CallbackInfo& info, const fm_workbook_t* handle, ColumnUnitFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  double out = 0.0;
  const fm_status_t rc = fn(handle, Workbook::ArgU32(info, 0), Workbook::ArgDouble(info, 2), ArgMode(info, 1), &out);
  return ValueOrThrow(env, rc, Napi::Number::New(env, out));
}

}  // namespace

Napi::Value Workbook::GetSheetView(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_view_info_t v = FM_STRUCT_INIT(fm_sheet_view_info_t);
  fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object view = Napi::Object::New(env);
  view.Set("zoomPct", Napi::Number::New(env, v.zoom_pct));
  view.Set("freezeRows", Napi::Number::New(env, v.freeze_rows));
  view.Set("freezeCols", Napi::Number::New(env, v.freeze_cols));
  view.Set("showGridLines", JsBool(env, v.show_grid_lines));
  view.Set("showRowColumnHeaders", JsBool(env, v.show_row_column_headers));
  view.Set("showZeros", JsBool(env, v.show_zeros));
  view.Set("rightToLeft", JsBool(env, v.right_to_left));
  view.Set("tabSelected", JsBool(env, v.tab_selected));
  view.Set("viewMode", Napi::Number::New(env, v.view_mode));
  view.Set("topLeftRow", Napi::Number::New(env, v.top_left_row));
  view.Set("topLeftCol", Napi::Number::New(env, v.top_left_col));
  view.Set("paneTopLeftRow", Napi::Number::New(env, v.pane_top_left_row));
  view.Set("paneTopLeftCol", Napi::Number::New(env, v.pane_top_left_col));
  view.Set("splitXTwips", Napi::Number::New(env, v.split_x_twips));
  view.Set("splitYTwips", Napi::Number::New(env, v.split_y_twips));
  view.Set("showFormulas", JsBool(env, v.show_formulas));
  view.Set("showOutlineSymbols", JsBool(env, v.show_outline_symbols));
  return view;
}

Napi::Value Workbook::GetVisibility(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t visibility = 0;
  const fm_status_t rc = fm_sheet_get_visibility(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &visibility);
  return ValueOrThrow(env, rc, Napi::Number::New(env, visibility));
}

Napi::Value Workbook::SetVisibility(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_visibility);
}

Napi::Value Workbook::GetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  int32_t present = 0;
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  const fm_status_t rc = fm_sheet_get_sheet_protection(handle_, sheet, &present, &p);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  if (present == 0) {
    return env.Null();
  }
  Napi::Object pr = Napi::Object::New(env);
  pr.Set("algorithmName", JsStringOrNull(env, p.algorithm_name));
  pr.Set("hashValue", JsStringOrNull(env, p.hash_value));
  pr.Set("saltValue", JsStringOrNull(env, p.salt_value));
  pr.Set("spinCount", Napi::Number::New(env, p.spin_count));
  pr.Set("legacyPasswordHash", JsStringOrNull(env, p.legacy_password_hash));
  pr.Set("allowEditObjects", JsBool(env, p.allow_edit_objects));
  pr.Set("allowEditScenarios", JsBool(env, p.allow_edit_scenarios));
  pr.Set("allowFormatCells", JsBool(env, p.allow_format_cells));
  pr.Set("allowFormatColumns", JsBool(env, p.allow_format_columns));
  pr.Set("allowFormatRows", JsBool(env, p.allow_format_rows));
  pr.Set("allowInsertColumns", JsBool(env, p.allow_insert_columns));
  pr.Set("allowInsertRows", JsBool(env, p.allow_insert_rows));
  pr.Set("allowInsertHyperlinks", JsBool(env, p.allow_insert_hyperlinks));
  pr.Set("allowDeleteColumns", JsBool(env, p.allow_delete_columns));
  pr.Set("allowDeleteRows", JsBool(env, p.allow_delete_rows));
  pr.Set("allowSelectLockedCells", JsTristate(env, p.allow_select_locked_cells));
  pr.Set("allowSelectUnlockedCells", JsTristate(env, p.allow_select_unlocked_cells));
  pr.Set("allowSort", JsBool(env, p.allow_sort));
  pr.Set("allowAutoFilter", JsBool(env, p.allow_auto_filter));
  pr.Set("allowPivotTables", JsBool(env, p.allow_pivot_tables));
  return pr;
}

Napi::Value Workbook::SetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  // `null` unprotects the sheet.
  if (info[1].IsNull()) {
    return ValueOrThrow(env, fm_sheet_set_sheet_protection(handle_, sheet, nullptr), env.Undefined());
  }
  Napi::Object in = ArgObject(info, 1);

  // A missing or `null` string is no stored value. The std::string buffers
  // outlive the C ABI call, which borrows them.
  CheckedSpecReader reader(env);
  std::string algorithm_name;
  std::string hash_value;
  std::string salt_value;
  std::string legacy_password_hash;
  const auto text = [&](const char* key, std::string* store) {
    return reader.String(in, key, store) ? store->c_str() : nullptr;
  };
  const auto flag = [&](const char* key) { return reader.Bool(in, key, false) ? 1 : 0; };
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  p.algorithm_name = text("algorithmName", &algorithm_name);
  p.hash_value = text("hashValue", &hash_value);
  p.salt_value = text("saltValue", &salt_value);
  p.spin_count = reader.U32(in, "spinCount", 0U);
  p.legacy_password_hash = text("legacyPasswordHash", &legacy_password_hash);
  p.allow_edit_objects = flag("allowEditObjects");
  p.allow_edit_scenarios = flag("allowEditScenarios");
  p.allow_format_cells = flag("allowFormatCells");
  p.allow_format_columns = flag("allowFormatColumns");
  p.allow_format_rows = flag("allowFormatRows");
  p.allow_insert_columns = flag("allowInsertColumns");
  p.allow_insert_rows = flag("allowInsertRows");
  p.allow_insert_hyperlinks = flag("allowInsertHyperlinks");
  p.allow_delete_columns = flag("allowDeleteColumns");
  p.allow_delete_rows = flag("allowDeleteRows");
  p.allow_select_locked_cells = reader.Tristate(in, "allowSelectLockedCells");
  p.allow_select_unlocked_cells = reader.Tristate(in, "allowSelectUnlockedCells");
  p.allow_sort = flag("allowSort");
  p.allow_auto_filter = flag("allowAutoFilter");
  p.allow_pivot_tables = flag("allowPivotTables");
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_sheet_protection(handle_, sheet, &p);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::UpdateSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object in = ArgObject(info, 1);

  // An omitted key keeps the stored value, `null` clears a string or `spinCount`, and a flag has no cleared form.
  CheckedSpecReader reader(env);
  std::string text[4];
  fm_sheet_protection_patch_t p = FM_STRUCT_INIT(fm_sheet_protection_patch_t);
  p.algorithm_name_op = PatchString(reader, in, "algorithmName", &text[0], &p.algorithm_name);
  p.hash_value_op = PatchString(reader, in, "hashValue", &text[1], &p.hash_value);
  p.salt_value_op = PatchString(reader, in, "saltValue", &text[2], &p.salt_value);
  p.legacy_password_hash_op = PatchString(reader, in, "legacyPasswordHash", &text[3], &p.legacy_password_hash);
  p.spin_count_op = PatchNullableU32(reader, in, "spinCount", &p.spin_count);
  p.allow_edit_objects_op = PatchBool(reader, in, "allowEditObjects", &p.allow_edit_objects);
  p.allow_edit_scenarios_op = PatchBool(reader, in, "allowEditScenarios", &p.allow_edit_scenarios);
  p.allow_format_cells_op = PatchBool(reader, in, "allowFormatCells", &p.allow_format_cells);
  p.allow_format_columns_op = PatchBool(reader, in, "allowFormatColumns", &p.allow_format_columns);
  p.allow_format_rows_op = PatchBool(reader, in, "allowFormatRows", &p.allow_format_rows);
  p.allow_insert_columns_op = PatchBool(reader, in, "allowInsertColumns", &p.allow_insert_columns);
  p.allow_insert_rows_op = PatchBool(reader, in, "allowInsertRows", &p.allow_insert_rows);
  p.allow_insert_hyperlinks_op = PatchBool(reader, in, "allowInsertHyperlinks", &p.allow_insert_hyperlinks);
  p.allow_delete_columns_op = PatchBool(reader, in, "allowDeleteColumns", &p.allow_delete_columns);
  p.allow_delete_rows_op = PatchBool(reader, in, "allowDeleteRows", &p.allow_delete_rows);
  p.allow_select_locked_cells_op = PatchBool(reader, in, "allowSelectLockedCells", &p.allow_select_locked_cells);
  p.allow_select_unlocked_cells_op = PatchBool(reader, in, "allowSelectUnlockedCells", &p.allow_select_unlocked_cells);
  p.allow_sort_op = PatchBool(reader, in, "allowSort", &p.allow_sort);
  p.allow_auto_filter_op = PatchBool(reader, in, "allowAutoFilter", &p.allow_auto_filter);
  p.allow_pivot_tables_op = PatchBool(reader, in, "allowPivotTables", &p.allow_pivot_tables);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_sheet_protection(handle_, sheet, &p), env.Undefined());
}

// Every member is optional: an omitted key keeps the stored value, and no member can be cleared.
Napi::Value Workbook::UpdateSheetView(const Napi::CallbackInfo& info) {
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  return InvokeSheetSpec(
      info,
      [&](CheckedSpecReader& reader, const Napi::Object& spec) {
        patch.zoom_pct_op = PatchU32(reader, spec, "zoomPct", &patch.zoom_pct);
        patch.freeze_rows_op = PatchU32(reader, spec, "freezeRows", &patch.freeze_rows);
        patch.freeze_cols_op = PatchU32(reader, spec, "freezeCols", &patch.freeze_cols);
        patch.show_grid_lines_op = PatchBool(reader, spec, "showGridLines", &patch.show_grid_lines);
        patch.show_row_column_headers_op =
            PatchBool(reader, spec, "showRowColumnHeaders", &patch.show_row_column_headers);
        patch.show_zeros_op = PatchBool(reader, spec, "showZeros", &patch.show_zeros);
        patch.right_to_left_op = PatchBool(reader, spec, "rightToLeft", &patch.right_to_left);
        patch.tab_selected_op = PatchBool(reader, spec, "tabSelected", &patch.tab_selected);
        patch.view_mode_op = PatchI32(reader, spec, "viewMode", &patch.view_mode);
        patch.top_left_row_op = PatchU32(reader, spec, "topLeftRow", &patch.top_left_row);
        patch.top_left_col_op = PatchU32(reader, spec, "topLeftCol", &patch.top_left_col);
        patch.pane_top_left_row_op = PatchU32(reader, spec, "paneTopLeftRow", &patch.pane_top_left_row);
        patch.pane_top_left_col_op = PatchU32(reader, spec, "paneTopLeftCol", &patch.pane_top_left_col);
        patch.split_x_twips_op = PatchDouble(reader, spec, "splitXTwips", &patch.split_x_twips);
        patch.split_y_twips_op = PatchDouble(reader, spec, "splitYTwips", &patch.split_y_twips);
        patch.show_formulas_op = PatchBool(reader, spec, "showFormulas", &patch.show_formulas);
        patch.show_outline_symbols_op = PatchBool(reader, spec, "showOutlineSymbols", &patch.show_outline_symbols);
      },
      [&](fm_workbook_t* handle) { return fm_sheet_update_view(handle, ArgU32(info, 0), &patch); });
}

Napi::Value Workbook::GetSheetSelection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  uint32_t row = 0;
  uint32_t col = 0;
  std::size_t total = 0;
  fm_status_t rc = fm_sheet_get_selection(handle_, sheet, &row, &col, nullptr, 0, &total);
  std::vector<fm_cell_rect_t> found(total);
  if (rc == 0) {
    rc = fm_sheet_get_selection(handle_, sheet, &row, &col, found.data(), found.size(), &total);
  }
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array ranges = Napi::Array::New(env);
  for (std::size_t i = 0; i < found.size(); ++i) {
    ranges.Set(static_cast<uint32_t>(i), RangeToJs(env, found[i]));
  }
  Napi::Object out = Napi::Object::New(env);
  Napi::Object active = Napi::Object::New(env);
  active.Set("row", Napi::Number::New(env, row));
  active.Set("col", Napi::Number::New(env, col));
  out.Set("active", active);
  out.Set("ranges", ranges);
  return out;
}

Napi::Value Workbook::SetSheetSelection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const Napi::Object selection = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  uint32_t row = 0;
  uint32_t col = 0;
  Napi::Object active;
  if (reader.RequiredObject(selection, "active", &active)) {
    row = reader.RequiredU32(active, "row");
    col = reader.RequiredU32(active, "col");
  }
  std::vector<fm_cell_rect_t> ranges;
  Napi::Array array;
  if (reader.Array(selection, "ranges", &array)) {
    const uint32_t n = array.Length();
    for (uint32_t i = 0; i < n && reader.ok(); ++i) {
      Napi::Value value;
      Napi::Object object;
      fm_cell_rect_t range{};
      if (reader.ArrayElement(array, i, &value) && reader.Object(value, "ranges[]", &object) &&
          ReadCellRect(reader, object, &range)) {
        ranges.push_back(range);
      }
    }
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_selection(handle_, sheet, row, col, ranges.data(), ranges.size()),
                      env.Undefined());
}

Napi::Value Workbook::GetSheetColumns(const Napi::CallbackInfo& info) {
  return SheetLayoutList(info, handle_, &fm_sheet_get_column_layout_count, &fm_sheet_get_column_layout_at,
                         &ColumnLayoutToJs);
}

Napi::Value Workbook::SetColumnWidth(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_column_width);
}

Napi::Value Workbook::SetColumnHidden(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t first = ArgU32(info, 1);
    const uint32_t last = ArgU32(info, 2);
    const int32_t value = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_column_hidden(handle, sheet, first, last, value);
  });
}

Napi::Value Workbook::SetColumnOutline(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_column_outline);
}

Napi::Value Workbook::SetColumnCollapsed(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t first = ArgU32(info, 1);
    const uint32_t last = ArgU32(info, 2);
    const int32_t value = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_column_collapsed(handle, sheet, first, last, value);
  });
}

Napi::Value Workbook::SetColumnStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_column_xf_index);
}

Napi::Value Workbook::GetSheetRowOverrides(const Napi::CallbackInfo& info) {
  return SheetLayoutList(info, handle_, &fm_sheet_get_row_layout_count, &fm_sheet_get_row_layout_at, &RowLayoutToJs);
}

Napi::Value Workbook::SetRowHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_height);
}

Napi::Value Workbook::SetRowAutoHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_auto_height);
}

Napi::Value Workbook::ClearRowHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_reset_row_height);
}

Napi::Value Workbook::ClearColumnWidth(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_reset_column_width);
}

Napi::Value Workbook::ClearColumnStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_reset_column_xf_index);
}

Napi::Value Workbook::GetSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  const fm_status_t rc = fm_sheet_get_format_defaults(handle_, ArgU32(info, 0), &d);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("defaultColumnWidthChars",
          JsNumberOrNull(env, d.has_default_column_width_chars, d.default_column_width_chars));
  out.Set("defaultRowHeightPt", JsNumberOrNull(env, d.has_default_row_height_pt, d.default_row_height_pt));
  out.Set("baseColumnWidthChars", Napi::Number::New(env, d.base_column_width_chars));
  out.Set("customHeight", Napi::Boolean::New(env, d.custom_height != 0));
  out.Set("zeroHeight", Napi::Boolean::New(env, d.zero_height != 0));
  out.Set("effectiveDefaultColumnWidthChars", Napi::Number::New(env, d.effective_default_column_width_chars));
  out.Set("effectiveDefaultRowHeightPt", Napi::Number::New(env, d.effective_default_row_height_pt));
  return out;
}

Napi::Value Workbook::UpdateSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // A patch: an omitted key keeps the stored value, `null` clears a nullable one.
  const Napi::Object spec = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  fm_sheet_format_defaults_patch_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  d.default_column_width_chars_op =
      PatchNullableDouble(reader, spec, "defaultColumnWidthChars", &d.default_column_width_chars);
  d.default_row_height_pt_op = PatchNullableDouble(reader, spec, "defaultRowHeightPt", &d.default_row_height_pt);
  d.base_column_width_chars_op = PatchDouble(reader, spec, "baseColumnWidthChars", &d.base_column_width_chars);
  d.custom_height_op = PatchBool(reader, spec, "customHeight", &d.custom_height);
  d.zero_height_op = PatchBool(reader, spec, "zeroHeight", &d.zero_height);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_format_defaults(handle_, ArgU32(info, 0), &d), env.Undefined());
}

// ---- Geometry -------------------------------------------------------

Napi::Value Workbook::GetCellRectPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Object range = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  fm_cell_rect_t cells{};
  if (!ReadCellRect(reader, range, &cells)) {
    return env.Undefined();
  }
  fm_rect_pt_info_t rect = FM_STRUCT_INIT(fm_rect_pt_info_t);
  const fm_status_t rc = fm_sheet_get_cell_rect_pt(handle_, ArgU32(info, 0), &cells, ArgMode(info, 2), &rect);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("x", Napi::Number::New(env, rect.x));
  out.Set("y", Napi::Number::New(env, rect.y));
  out.Set("widthPt", Napi::Number::New(env, rect.width_pt));
  out.Set("heightPt", Napi::Number::New(env, rect.height_pt));
  return out;
}

Napi::Value Workbook::GetColumnWidthPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  double pt = 0.0;
  const fm_status_t rc = fm_sheet_get_column_width_pt(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgMode(info, 2), &pt);
  return ValueOrThrow(env, rc, Napi::Number::New(env, pt));
}

Napi::Value Workbook::GetRowHeightPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  double pt = 0.0;
  const fm_status_t rc = fm_sheet_get_row_height_pt(handle_, ArgU32(info, 0), ArgU32(info, 1), &pt);
  return ValueOrThrow(env, rc, Napi::Number::New(env, pt));
}

Napi::Value Workbook::GetWidthModel(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_width_model_info_t model = FM_STRUCT_INIT(fm_width_model_info_t);
  const fm_status_t rc = fm_sheet_get_width_model(handle_, ArgU32(info, 0), ArgMode(info, 1), &model);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("pointsPerChar", Napi::Number::New(env, model.points_per_char));
  out.Set("paddingPt", Napi::Number::New(env, model.padding_pt));
  out.Set("normalFontSizePt", Napi::Number::New(env, model.normal_font_size_pt));
  out.Set("calibrated", Napi::Boolean::New(env, model.calibrated != 0));
  out.Set("normalFontName", JsString(env, model.normal_font_name));
  out.Set("platform", Napi::Number::New(env, model.platform));
  return out;
}

Napi::Value Workbook::ColumnCharsToPt(const Napi::CallbackInfo& info) {
  return ConvertColumnUnit(info, handle_, &fm_sheet_convert_column_chars_to_pt);
}

Napi::Value Workbook::ColumnPtToChars(const Napi::CallbackInfo& info) {
  return ConvertColumnUnit(info, handle_, &fm_sheet_convert_column_pt_to_chars);
}

Napi::Value Workbook::SetRowHidden(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t first = ArgU32(info, 1);
    const uint32_t last = ArgU32(info, 2);
    const int32_t value = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_row_hidden(handle, sheet, first, last, value);
  });
}

Napi::Value Workbook::SetRowOutline(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_outline);
}

Napi::Value Workbook::SetRowCollapsed(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t first = ArgU32(info, 1);
    const uint32_t last = ArgU32(info, 2);
    const int32_t value = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_row_collapsed(handle, sheet, first, last, value);
  });
}

Napi::Value Workbook::SetRowStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_xf_index);
}

Napi::Value Workbook::ClearRowStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_reset_row_xf_index);
}

namespace {

Napi::Value RangeStyleBasesResult(Napi::Env env, fm_workbook_t* handle, uint32_t sheet, const fm_cell_rect_t& rect,
                                  const std::vector<fm_cell_address_t>& excluded) {
  size_t len = 0;
  fm_status_t rc =
      fm_sheet_get_range_xf_bases_excluding(handle, sheet, &rect, excluded.data(), excluded.size(), nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_range_xf_bases_excluding(handle, sheet, &rect, excluded.data(), excluded.size(), xfs.data(),
                                               xfs.size(), &len);
  }
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array list = Napi::Array::New(env, xfs.size());
  for (size_t i = 0; i < xfs.size(); ++i) {
    list.Set(static_cast<uint32_t>(i), JsNumber(env, xfs[i]));
  }
  return list;
}

// Reads two equal-length arrays of unsigned integers element by element; `key` names them in diagnostics.
void PullParallelU32(CheckedSpecReader& reader, const Napi::Array& first_arr, const Napi::Array& second_arr,
                     const char* key, std::vector<uint32_t>* first, std::vector<uint32_t>* second) {
  const uint32_t count = first_arr.Length();
  first->assign(count, 0U);
  second->assign(count, 0U);
  for (uint32_t i = 0; i < count; ++i) {
    Napi::Value v;
    if (reader.ArrayElement(first_arr, i, &v)) {
      (*first)[i] = reader.U32(v, key, 0U);
    }
    if (reader.ArrayElement(second_arr, i, &v)) {
      (*second)[i] = reader.U32(v, key, 0U);
    }
  }
}

}  // namespace

Napi::Value Workbook::RangeStyleBases(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  return RangeStyleBasesResult(env, handle_, ArgU32(info, 0), rect, {});
}

Napi::Value Workbook::RangeStyleBasesExcluding(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const auto refused = [&env](const char* path, const char* constraint) {
    ThrowArgumentTypeError(env, path, constraint);
    return env.Undefined();
  };
  Napi::Array rows_arr;
  Napi::Array cols_arr;
  if (!ArrayArg(info, 2, &rows_arr) || !ArrayArg(info, 3, &cols_arr)) {
    return refused("excludedRows and excludedCols", "must be an array");
  }
  if (rows_arr.Length() != cols_arr.Length()) {
    return refused("excludedRows and excludedCols", "must have the same length");
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  std::vector<uint32_t> rows;
  std::vector<uint32_t> cols;
  PullParallelU32(reader, rows_arr, cols_arr, "excludedRows/excludedCols", &rows, &cols);
  if (!reader.ok()) {
    return env.Undefined();
  }
  std::vector<fm_cell_address_t> excluded(rows.size());
  for (std::size_t i = 0; i < rows.size(); ++i) {
    excluded[i] = fm_cell_address_t{rows[i], cols[i]};
  }
  return RangeStyleBasesResult(env, handle_, ArgU32(info, 0), rect, excluded);
}

Napi::Value Workbook::RemapRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array from_arr;
  Napi::Array to_arr;
  if (!ArrayArg(info, 2, &from_arr) || !ArrayArg(info, 3, &to_arr)) {
    ThrowArgumentTypeError(env, "fromXfs and toXfs", "must be an array");
    return env.Undefined();
  }
  if (from_arr.Length() != to_arr.Length()) {
    ThrowArgumentTypeError(env, "fromXfs and toXfs", "must have the same length");
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  std::vector<uint32_t> from;
  std::vector<uint32_t> to;
  PullParallelU32(reader, from_arr, to_arr, "fromXfs/toXfs", &from, &to);
  if (!reader.ok()) {
    return env.Undefined();
  }
  std::vector<fm_xf_remap_t> pairs(from.size());
  for (std::size_t i = 0; i < from.size(); ++i) {
    pairs[i] = fm_xf_remap_t{from[i], to[i]};
  }
  const fm_status_t rc =
      fm_sheet_remap_range_xfs(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &rect, pairs.data(), pairs.size());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::SnapshotRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  uint8_t* bytes = nullptr;
  size_t len = 0;
  const fm_status_t rc =
      fm_sheet_snapshot_range_xfs(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &rect, &bytes, &len);
  if (rc != 0) {
    const Napi::Value failure = ThrowStatus(env, rc);
    fm_buffer_destroy(bytes);
    return failure;
  }
  return TakeSavedBytes(env, bytes, len);
}

Napi::Value Workbook::RestoreRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Uint8Array u8;
  if (!BytesArg(info, 1, &u8)) {
    ThrowArgumentTypeError(env, "bytes", "must be a Uint8Array");
    return env.Undefined();
  }
  const fm_status_t rc =
      fm_sheet_restore_range_xfs(handle_, static_cast<std::size_t>(ArgU32(info, 0)), u8.Data(), u8.ElementLength());
  return ValueOrThrow(env, rc, env.Undefined());
}

}  // namespace formulon_node
