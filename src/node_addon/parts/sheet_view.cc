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
  return info[idx].IsUndefined() ? static_cast<int32_t>(FM_GEOMETRY_DISPLAY) : Workbook::ArgI32(info, idx);
}

Napi::Object DefaultSheetView(Napi::Env env) {
  Napi::Object view = Napi::Object::New(env);
  view.Set("zoomScale", Napi::Number::New(env, 100));
  view.Set("freezeRows", Napi::Number::New(env, 0));
  view.Set("freezeCols", Napi::Number::New(env, 0));
  view.Set("tabHidden", Napi::Number::New(env, 0));
  view.Set("visibility", Napi::Number::New(env, 0));
  view.Set("showGridLines", Napi::Number::New(env, 1));
  view.Set("showRowColHeaders", Napi::Number::New(env, 1));
  view.Set("showZeros", Napi::Number::New(env, 1));
  view.Set("rightToLeft", Napi::Number::New(env, 0));
  view.Set("tabSelected", Napi::Number::New(env, 0));
  view.Set("viewMode", Napi::String::New(env, ""));
  view.Set("topLeftRow", Napi::Number::New(env, 0));
  view.Set("topLeftCol", Napi::Number::New(env, 0));
  view.Set("paneTopLeftRow", Napi::Number::New(env, 0));
  view.Set("paneTopLeftCol", Napi::Number::New(env, 0));
  view.Set("splitX", Napi::Number::New(env, 0));
  view.Set("splitY", Napi::Number::New(env, 0));
  view.Set("showFormulas", Napi::Number::New(env, 0));
  view.Set("showOutlineSymbols", Napi::Number::New(env, 1));
  return view;
}

// Enumerates a per-sheet layout list into `{ status, <field>: [...] }`,
// skipping entries the C getter cannot read.
template <typename Entry>
Napi::Value SheetLayoutList(const Napi::CallbackInfo& info, const fm_workbook_t* handle, const char* field,
                            fm_status_t (*count_fn)(const fm_workbook_t*, size_t, size_t*),
                            fm_status_t (*at_fn)(const fm_workbook_t*, size_t, size_t, Entry*),
                            Napi::Object (*to_js)(Napi::Env, const Entry&)) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeFieldResult(env, MakeErrorStatus(env, kBindingInvalidHandle), field, Napi::Array::New(env));
  }
  const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = count_fn(handle, sheet, &count);
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), field, Napi::Array::New(env));
  }
  Napi::Array arr = Napi::Array::New(env, count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    Entry entry{};
    if (at_fn(handle, sheet, i, &entry) != 0) {
      continue;
    }
    arr.Set(static_cast<uint32_t>(emitted), to_js(env, entry));
    ++emitted;
  }
  return MakeFieldResult(env, MakeOkStatus(env), field, arr);
}

Napi::Object ColumnLayoutToJs(Napi::Env env, const fm_column_layout_t& entry) {
  Napi::Object col = Napi::Object::New(env);
  col.Set("first", Napi::Number::New(env, entry.first));
  col.Set("last", Napi::Number::New(env, entry.last));
  col.Set("width", Napi::Number::New(env, entry.width));
  col.Set("hidden", Napi::Number::New(env, entry.hidden));
  col.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
  col.Set("collapsed", Napi::Number::New(env, static_cast<int32_t>(entry.collapsed)));
  // The C getter reports logical presence, including legacy aggregate
  // layouts with a non-zero width but a clear raw presence bit. Keep the
  // binding defensive in case an older ABI implementation is loaded.
  col.Set("hasWidth", Napi::Number::New(env, entry.has_width || entry.width != 0.0));
  col.Set("hasStyle", Napi::Number::New(env, entry.has_style));
  col.Set("styleXf", Napi::Number::New(env, entry.style_xf));
  return col;
}

Napi::Object RowLayoutToJs(Napi::Env env, const fm_row_layout_t& entry) {
  Napi::Object row = Napi::Object::New(env);
  row.Set("row", Napi::Number::New(env, entry.row));
  row.Set("height", Napi::Number::New(env, entry.height));
  row.Set("hasHeight", Napi::Number::New(env, entry.has_height));
  row.Set("customHeight", Napi::Number::New(env, entry.custom_height));
  row.Set("hidden", Napi::Number::New(env, entry.hidden));
  row.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
  row.Set("collapsed", Napi::Number::New(env, static_cast<int32_t>(entry.collapsed)));
  row.Set("hasStyle", Napi::Number::New(env, entry.has_style));
  row.Set("styleXf", Napi::Number::New(env, entry.style_xf));
  return row;
}

// Shared body of the per-sheet boolean view setters.
using SheetFlagFn = fm_status_t (*)(fm_workbook_t*, size_t, int32_t);

Napi::Value InvokeSheetFlag(Workbook& wb, const Napi::CallbackInfo& info, SheetFlagFn fn) {
  return wb.StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(Workbook::ArgU32(info, 0));
    return fn(handle, sheet, Workbook::ArgFlag(info, 1));
  });
}

// Shared body of the column width unit converters (chars <-> points).
using ColumnUnitFn = fm_status_t (*)(const fm_workbook_t*, size_t, int32_t, double, double*);

Napi::Value ConvertColumnUnit(const Napi::CallbackInfo& info, const fm_workbook_t* handle, ColumnUnitFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  double out = 0.0;
  const fm_status_t rc = fn(handle, Workbook::ArgU32(info, 0), ArgMode(info, 1), Workbook::ArgDouble(info, 2), &out);
  return MakeNumberResult(env, rc, out);
}

}  // namespace

Napi::Value Workbook::GetSheetView(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "view", DefaultSheetView(env));
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_view_t v{};
  fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "view", DefaultSheetView(env));
  }
  Napi::Object view = Napi::Object::New(env);
  view.Set("zoomScale", Napi::Number::New(env, v.zoom_scale));
  view.Set("freezeRows", Napi::Number::New(env, v.freeze_rows));
  view.Set("freezeCols", Napi::Number::New(env, v.freeze_cols));
  view.Set("tabHidden", Napi::Number::New(env, v.tab_hidden));
  view.Set("visibility", Napi::Number::New(env, v.visibility));
  view.Set("showGridLines", Napi::Number::New(env, v.show_grid_lines));
  view.Set("showRowColHeaders", Napi::Number::New(env, v.show_row_col_headers));
  view.Set("showZeros", Napi::Number::New(env, v.show_zeros));
  view.Set("rightToLeft", Napi::Number::New(env, v.right_to_left));
  view.Set("tabSelected", Napi::Number::New(env, v.tab_selected));
  view.Set("viewMode", JsString(env, v.view_mode));
  view.Set("topLeftRow", Napi::Number::New(env, v.top_left_row));
  view.Set("topLeftCol", Napi::Number::New(env, v.top_left_col));
  view.Set("paneTopLeftRow", Napi::Number::New(env, v.pane_top_left_row));
  view.Set("paneTopLeftCol", Napi::Number::New(env, v.pane_top_left_col));
  view.Set("splitX", Napi::Number::New(env, v.split_x));
  view.Set("splitY", Napi::Number::New(env, v.split_y));
  view.Set("showFormulas", Napi::Number::New(env, v.show_formulas));
  view.Set("showOutlineSymbols", Napi::Number::New(env, v.show_outline_symbols));
  return MakeFieldResult(env, MakeOkStatus(env), "view", view);
}

Napi::Value Workbook::GetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  // `protection` is emitted on every exit path, as `view` already is above
  // and as the WASM binding does for both -- a value object there cannot
  // omit a field. A caller that skips the status check then reads a
  // defaulted record rather than tripping over a missing key.
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_protection_t p{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_protection(handle_, sheet, &p) : kBindingInvalidHandle;
  if (rc != 0) {
    p = fm_sheet_protection_t{};
  }
  Napi::Object pr = Napi::Object::New(env);
  pr.Set("enabled", Napi::Number::New(env, p.enabled));
  pr.Set("algorithmName", JsString(env, p.algorithm_name));
  pr.Set("hashValue", JsString(env, p.hash_value));
  pr.Set("saltValue", JsString(env, p.salt_value));
  pr.Set("spinCount", Napi::Number::New(env, p.spin_count));
  pr.Set("legacyPassword", JsString(env, p.legacy_password));
  pr.Set("sheet", Napi::Number::New(env, p.sheet));
  pr.Set("objects", Napi::Number::New(env, p.objects));
  pr.Set("scenarios", Napi::Number::New(env, p.scenarios));
  pr.Set("formatCells", Napi::Number::New(env, p.format_cells));
  pr.Set("formatColumns", Napi::Number::New(env, p.format_columns));
  pr.Set("formatRows", Napi::Number::New(env, p.format_rows));
  pr.Set("insertColumns", Napi::Number::New(env, p.insert_columns));
  pr.Set("insertRows", Napi::Number::New(env, p.insert_rows));
  pr.Set("insertHyperlinks", Napi::Number::New(env, p.insert_hyperlinks));
  pr.Set("deleteColumns", Napi::Number::New(env, p.delete_columns));
  pr.Set("deleteRows", Napi::Number::New(env, p.delete_rows));
  pr.Set("selectLockedCells", Napi::Number::New(env, p.select_locked_cells));
  pr.Set("selectUnlockedCells", Napi::Number::New(env, p.select_unlocked_cells));
  pr.Set("sort", Napi::Number::New(env, p.sort));
  pr.Set("autoFilter", Napi::Number::New(env, p.auto_filter));
  pr.Set("pivotTables", Napi::Number::New(env, p.pivot_tables));
  return MakeFieldResult(env, MakeStatus(env, rc), "protection", pr);
}

Napi::Value Workbook::SetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object in = ArgObject(info, 1);

  // Keep the std::string buffers alive until after the C ABI call so the
  // borrowed `const char*` fields stay valid.
  CheckedSpecReader reader(env);
  std::string algorithm_name;
  std::string hash_value;
  std::string salt_value;
  std::string legacy_password;
  (void)reader.String(in, "algorithmName", &algorithm_name);
  (void)reader.String(in, "hashValue", &hash_value);
  (void)reader.String(in, "saltValue", &salt_value);
  (void)reader.String(in, "legacyPassword", &legacy_password);
  if (!reader.ok()) {
    return env.Undefined();
  }

  fm_sheet_protection_t p{};
  p.enabled = reader.I32(in, "enabled", 0);
  p.algorithm_name = algorithm_name.c_str();
  p.hash_value = hash_value.c_str();
  p.salt_value = salt_value.c_str();
  p.spin_count = reader.U32(in, "spinCount", 0U);
  p.legacy_password = legacy_password.c_str();
  p.sheet = reader.I32(in, "sheet", 0);
  p.objects = reader.I32(in, "objects", 0);
  p.scenarios = reader.I32(in, "scenarios", 0);
  p.format_cells = reader.I32(in, "formatCells", 0);
  p.format_columns = reader.I32(in, "formatColumns", 0);
  p.format_rows = reader.I32(in, "formatRows", 0);
  p.insert_columns = reader.I32(in, "insertColumns", 0);
  p.insert_rows = reader.I32(in, "insertRows", 0);
  p.insert_hyperlinks = reader.I32(in, "insertHyperlinks", 0);
  p.delete_columns = reader.I32(in, "deleteColumns", 0);
  p.delete_rows = reader.I32(in, "deleteRows", 0);
  p.select_locked_cells = reader.I32(in, "selectLockedCells", 0);
  p.select_unlocked_cells = reader.I32(in, "selectUnlockedCells", 0);
  p.sort = reader.I32(in, "sort", 0);
  p.auto_filter = reader.I32(in, "autoFilter", 0);
  p.pivot_tables = reader.I32(in, "pivotTables", 0);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_protection(handle_, sheet, &p);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetZoom(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_zoom);
}

Napi::Value Workbook::SetSheetFreeze(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_freeze);
}

Napi::Value Workbook::SetSheetTabHidden(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_tab_hidden);
}

Napi::Value Workbook::SetSheetVisibility(const Napi::CallbackInfo& info) {
  // Raw ordinal: the C ABI rejects an unknown value, so coercing it to a
  // narrower type here would turn a caller's mistake into a silent state.
  return InvokeSheetCall(info, &fm_sheet_set_visibility);
}

Napi::Value Workbook::SetSheetShowGridLines(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_show_grid_lines);
}

Napi::Value Workbook::SetSheetShowRowColHeaders(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_show_row_col_headers);
}

Napi::Value Workbook::SetSheetShowZeros(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_show_zeros);
}

Napi::Value Workbook::SetSheetShowFormulas(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_show_formulas);
}

Napi::Value Workbook::SetSheetShowOutlineSymbols(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_show_outline_symbols);
}

Napi::Value Workbook::SetSheetRightToLeft(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_right_to_left);
}

Napi::Value Workbook::SetSheetTopLeft(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_top_left);
}

Napi::Value Workbook::SetSheetPaneTopLeft(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_pane_top_left);
}

Napi::Value Workbook::SetSheetSplit(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_split);
}

Napi::Value Workbook::GetSheetSelection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  uint32_t row = 0;
  uint32_t col = 0;
  std::size_t total = 0;
  fm_status_t rc = handle_ != nullptr ? fm_sheet_get_selection(handle_, sheet, &row, &col, nullptr, 0, &total)
                                      : kBindingInvalidHandle;
  std::vector<fm_merge_range> found(total);
  if (rc == 0) {
    rc = fm_sheet_get_selection(handle_, sheet, &row, &col, found.data(), found.size(), &total);
  }
  Napi::Array ranges = Napi::Array::New(env);
  for (std::size_t i = 0; rc == 0 && i < found.size(); ++i) {
    ranges.Set(static_cast<uint32_t>(i), RangeToJs(env, found[i]));
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("activeRow", Napi::Number::New(env, rc == 0 ? row : 0U));
  out.Set("activeCol", Napi::Number::New(env, rc == 0 ? col : 0U));
  out.Set("ranges", ranges);
  return out;
}

Napi::Value Workbook::SetSheetSelection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const Napi::Object selection = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  const uint32_t row = reader.U32(selection, "activeRow", 0U);
  const uint32_t col = reader.U32(selection, "activeCol", 0U);
  std::vector<fm_merge_range> ranges;
  Napi::Array array;
  if (reader.Array(selection, "ranges", &array)) {
    const uint32_t n = array.Length();
    for (uint32_t i = 0; i < n && reader.ok(); ++i) {
      Napi::Value value;
      Napi::Object object;
      fm_merge_range range{};
      if (reader.ArrayElement(array, i, &value) && reader.Object(value, "ranges[]", &object) &&
          ReadMergeRange(reader, object, &range)) {
        ranges.push_back(range);
      }
    }
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_selection(handle_, sheet, row, col, ranges.data(), ranges.size()));
}

Napi::Value Workbook::SetSheetTabSelected(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_tab_selected);
}

Napi::Value Workbook::SetSheetViewMode(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const std::string mode = ArgString(info, 1);
    return fm_sheet_set_view_mode(handle, sheet, mode.c_str());
  });
}

Napi::Value Workbook::GetSheetColumns(const Napi::CallbackInfo& info) {
  return SheetLayoutList(info, handle_, "columns", &fm_sheet_get_column_count, &fm_sheet_get_column, &ColumnLayoutToJs);
}

Napi::Value Workbook::SetColumnWidth(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_column_width);
}

Napi::Value Workbook::SetColumnHidden(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t first = ArgU32(info, 1);
    const uint32_t last = ArgU32(info, 2);
    return fm_sheet_set_column_hidden(handle, sheet, first, last, ArgFlag(info, 3));
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
    return fm_sheet_set_column_collapsed(handle, sheet, first, last, ArgFlag(info, 3));
  });
}

Napi::Value Workbook::SetColumnStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_column_style);
}

Napi::Value Workbook::GetSheetRowOverrides(const Napi::CallbackInfo& info) {
  return SheetLayoutList(info, handle_, "rows", &fm_sheet_get_row_override_count, &fm_sheet_get_row_override,
                         &RowLayoutToJs);
}

Napi::Value Workbook::SetRowHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_height);
}

Napi::Value Workbook::ClearRowHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_row_height);
}

Napi::Value Workbook::ClearColumnWidth(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_column_width);
}

Napi::Value Workbook::ClearColumnStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_column_style);
}

Napi::Value Workbook::GetSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_sheet_format_defaults d{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_format_defaults(handle_, ArgU32(info, 0), &d) : kBindingInvalidHandle;
  if (rc != 0) {
    d = fm_sheet_format_defaults{};
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("defaultColWidth", Napi::Number::New(env, d.default_col_width));
  out.Set("defaultRowHeight", Napi::Number::New(env, d.default_row_height));
  out.Set("baseColWidth", Napi::Number::New(env, d.base_col_width));
  out.Set("hasDefaultColWidth", Napi::Boolean::New(env, d.has_default_col_width != 0));
  out.Set("hasDefaultRowHeight", Napi::Boolean::New(env, d.has_default_row_height != 0));
  out.Set("customHeight", Napi::Boolean::New(env, d.custom_height != 0));
  out.Set("zeroHeight", Napi::Boolean::New(env, d.zero_height != 0));
  out.Set("effectiveDefaultColWidth", Napi::Number::New(env, d.effective_default_col_width));
  out.Set("effectiveDefaultRowHeight", Napi::Number::New(env, d.effective_default_row_height));
  return out;
}

Napi::Value Workbook::SetSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const Napi::Object spec = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  bool has_default_col_width = false;
  bool has_default_row_height = false;
  // An omitted flag keeps the stored one.
  fm_sheet_format_defaults d{};
  if (fm_sheet_get_format_defaults(handle_, ArgU32(info, 0), &d) != 0) {
    d = fm_sheet_format_defaults{};
  }
  d.default_col_width = reader.Double(spec, "defaultColWidth", 0.0, &has_default_col_width);
  d.default_row_height = reader.Double(spec, "defaultRowHeight", 0.0, &has_default_row_height);
  d.base_col_width = reader.Double(spec, "baseColWidth", 8.0);
  d.has_default_col_width = reader.Bool(spec, "hasDefaultColWidth", has_default_col_width) ? 1 : 0;
  d.has_default_row_height = reader.Bool(spec, "hasDefaultRowHeight", has_default_row_height) ? 1 : 0;
  d.custom_height = reader.Bool(spec, "customHeight", d.custom_height != 0) ? 1 : 0;
  d.zero_height = reader.Bool(spec, "zeroHeight", d.zero_height != 0) ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_format_defaults(handle_, ArgU32(info, 0), &d));
}

// ---- Geometry -------------------------------------------------------

Napi::Value Workbook::GetCellRectPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_rect_pt rect{};
  Napi::Object status = Napi::Object::New(env);
  if (handle_ == nullptr) {
    status = NullHandleError(env);
  } else {
    const Napi::Object range = ArgObject(info, 1);
    CheckedSpecReader reader(env);
    const uint32_t first_row = reader.U32(range, "firstRow", 0U);
    const uint32_t first_col = reader.U32(range, "firstCol", 0U);
    const uint32_t last_row = reader.U32(range, "lastRow", 0U);
    const uint32_t last_col = reader.U32(range, "lastCol", 0U);
    if (!reader.ok()) {
      return env.Undefined();
    }
    const fm_status_t rc = fm_sheet_cell_rect_pt(handle_, ArgU32(info, 0), first_row, first_col, last_row, last_col,
                                                 ArgMode(info, 2), &rect);
    if (rc != 0) {
      rect = fm_rect_pt{};
    }
    status = MakeStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("x", Napi::Number::New(env, rect.x));
  out.Set("y", Napi::Number::New(env, rect.y));
  out.Set("width", Napi::Number::New(env, rect.width));
  out.Set("height", Napi::Number::New(env, rect.height));
  return out;
}

Napi::Value Workbook::GetColumnWidthPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  double pt = 0.0;
  const fm_status_t rc = fm_sheet_column_width_pt(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgMode(info, 2), &pt);
  return MakeNumberResult(env, rc, pt);
}

Napi::Value Workbook::GetRowHeightPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  double pt = 0.0;
  const fm_status_t rc = fm_sheet_row_height_pt(handle_, ArgU32(info, 0), ArgU32(info, 1), &pt);
  return MakeNumberResult(env, rc, pt);
}

Napi::Value Workbook::GetWidthModel(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_width_model model{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_width_model(handle_, ArgU32(info, 0), ArgMode(info, 1), &model)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    model = fm_width_model{};
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("pointsPerChar", Napi::Number::New(env, model.points_per_char));
  out.Set("paddingPt", Napi::Number::New(env, model.padding_pt));
  out.Set("normalFontSize", Napi::Number::New(env, model.normal_font_size));
  out.Set("calibrated", Napi::Boolean::New(env, model.calibrated != 0));
  out.Set("normalFontName", JsString(env, model.normal_font_name));
  out.Set("platform", JsString(env, model.platform));
  return out;
}

Napi::Value Workbook::ColumnCharsToPt(const Napi::CallbackInfo& info) {
  return ConvertColumnUnit(info, handle_, &fm_sheet_column_chars_to_pt);
}

Napi::Value Workbook::ColumnPtToChars(const Napi::CallbackInfo& info) {
  return ConvertColumnUnit(info, handle_, &fm_sheet_column_pt_to_chars);
}

Napi::Value Workbook::SetRowHidden(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t row = ArgU32(info, 1);
    return fm_sheet_set_row_hidden(handle, sheet, row, ArgFlag(info, 2));
  });
}

Napi::Value Workbook::SetRowOutline(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_outline);
}

Napi::Value Workbook::SetRowCollapsed(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    const uint32_t row = ArgU32(info, 1);
    return fm_sheet_set_row_collapsed(handle, sheet, row, ArgFlag(info, 2));
  });
}

Napi::Value Workbook::SetRowStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_style);
}

Napi::Value Workbook::ClearRowStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_row_style);
}

Napi::Value Workbook::RangeStyleBases(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("xfIndices", Napi::Array::New(env));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  size_t len = 0;
  fm_status_t rc = fm_sheet_range_style_bases(handle_, sheet, ArgU32(info, 1), ArgU32(info, 2), ArgU32(info, 3),
                                              ArgU32(info, 4), nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_range_style_bases(handle_, sheet, ArgU32(info, 1), ArgU32(info, 2), ArgU32(info, 3), ArgU32(info, 4),
                                    xfs.data(), xfs.size(), &len);
  }
  if (rc != 0) {
    out.Set("status", MakeStatus(env, rc));
    return out;
  }
  Napi::Array list = Napi::Array::New(env, xfs.size());
  for (size_t i = 0; i < xfs.size(); ++i) {
    list.Set(static_cast<uint32_t>(i), JsNumber(env, xfs[i]));
  }
  out.Set("xfIndices", list);
  out.Set("status", MakeOkStatus(env));
  return out;
}

Napi::Value Workbook::RemapRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Array from_arr;
  Napi::Array to_arr;
  if (!ArrayArg(info, 5, &from_arr) || !ArrayArg(info, 6, &to_arr)) {
    return MakeBindingArgumentError(
        env, "remapRangeStyle expects (sheet, firstRow, firstCol, lastRow, lastCol, fromXfs:number[], toXfs:number[])");
  }
  if (from_arr.Length() != to_arr.Length()) {
    return MakeBindingArgumentError(env, "remapRangeStyle: fromXfs and toXfs must have the same length");
  }
  const uint32_t count = from_arr.Length();
  CheckedSpecReader reader(env);
  std::vector<uint32_t> from(count);
  std::vector<uint32_t> to(count);
  for (uint32_t i = 0; i < count; ++i) {
    const std::string key = std::to_string(i);
    from[i] = reader.U32(from_arr, key.c_str(), 0U);
    to[i] = reader.U32(to_arr, key.c_str(), 0U);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_remap_range_style(handle_, static_cast<std::size_t>(ArgU32(info, 0)), ArgU32(info, 1),
                                                    ArgU32(info, 2), ArgU32(info, 3), ArgU32(info, 4), from.data(),
                                                    to.data(), count));
}

Napi::Value Workbook::SnapshotRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("bytes", Napi::Uint8Array::New(env, 0));
    out.Set("status", NullHandleError(env));
    return out;
  }
  uint8_t* bytes = nullptr;
  size_t len = 0;
  const fm_status_t rc =
      fm_sheet_snapshot_range_style(handle_, static_cast<std::size_t>(ArgU32(info, 0)), ArgU32(info, 1),
                                    ArgU32(info, 2), ArgU32(info, 3), ArgU32(info, 4), &bytes, &len);
  if (rc == 0) {
    out.Set("bytes", TakeSavedBytes(env, bytes, len));
  } else {
    fm_buffer_free(bytes);
    out.Set("bytes", Napi::Uint8Array::New(env, 0));
  }
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::RestoreRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Uint8Array u8;
  if (!BytesArg(info, 1, &u8)) {
    return MakeBindingArgumentError(env, "restoreRangeStyle expects (sheet:number, bytes:Uint8Array)");
  }
  return MakeStatus(env, fm_sheet_restore_range_style(handle_, static_cast<std::size_t>(ArgU32(info, 0)), u8.Data(),
                                                      u8.ElementLength()));
}

}  // namespace formulon_node
