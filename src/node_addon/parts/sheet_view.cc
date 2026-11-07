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

Napi::Object ColumnLayoutToJs(Napi::Env env, const fm_column_layout_t& entry) {
  Napi::Object col = Napi::Object::New(env);
  col.Set("first", Napi::Number::New(env, entry.first));
  col.Set("last", Napi::Number::New(env, entry.last));
  col.Set("width", JsNumberOrNull(env, entry.has_width, entry.width));
  col.Set("hidden", JsBool(env, entry.hidden));
  col.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
  col.Set("collapsed", JsBool(env, entry.collapsed));
  col.Set("styleXf", JsNumberOrNull(env, entry.has_style, entry.style_xf));
  return col;
}

Napi::Object RowLayoutToJs(Napi::Env env, const fm_row_layout_t& entry) {
  Napi::Object row = Napi::Object::New(env);
  row.Set("row", Napi::Number::New(env, entry.row));
  row.Set("height", JsNumberOrNull(env, entry.has_height, entry.height));
  row.Set("hidden", JsBool(env, entry.hidden));
  row.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
  row.Set("collapsed", JsBool(env, entry.collapsed));
  row.Set("styleXf", JsNumberOrNull(env, entry.has_style, entry.style_xf));
  row.Set("customHeight", JsBool(env, entry.custom_height));
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
    return ThrowInvalidHandle(env);
  }
  double out = 0.0;
  const fm_status_t rc = fn(handle, Workbook::ArgU32(info, 0), ArgMode(info, 1), Workbook::ArgDouble(info, 2), &out);
  return ValueOrThrow(env, rc, Napi::Number::New(env, out));
}

}  // namespace

Napi::Value Workbook::GetSheetView(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_view_t v = FM_STRUCT_INIT(fm_sheet_view_t);
  fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object view = Napi::Object::New(env);
  view.Set("zoomScale", Napi::Number::New(env, v.zoom_scale));
  view.Set("freezeRows", Napi::Number::New(env, v.freeze_rows));
  view.Set("freezeCols", Napi::Number::New(env, v.freeze_cols));
  view.Set("visibility", Napi::Number::New(env, v.visibility));
  view.Set("showGridLines", JsBool(env, v.show_grid_lines));
  view.Set("showRowColHeaders", JsBool(env, v.show_row_col_headers));
  view.Set("showZeros", JsBool(env, v.show_zeros));
  view.Set("rightToLeft", JsBool(env, v.right_to_left));
  view.Set("tabSelected", JsBool(env, v.tab_selected));
  view.Set("viewMode", Napi::Number::New(env, v.view_mode));
  view.Set("topLeftRow", Napi::Number::New(env, v.top_left_row));
  view.Set("topLeftCol", Napi::Number::New(env, v.top_left_col));
  view.Set("paneTopLeftRow", Napi::Number::New(env, v.pane_top_left_row));
  view.Set("paneTopLeftCol", Napi::Number::New(env, v.pane_top_left_col));
  view.Set("splitX", Napi::Number::New(env, v.split_x));
  view.Set("splitY", Napi::Number::New(env, v.split_y));
  view.Set("showFormulas", JsBool(env, v.show_formulas));
  view.Set("showOutlineSymbols", JsBool(env, v.show_outline_symbols));
  return view;
}

Napi::Value Workbook::GetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  const fm_status_t rc = fm_sheet_get_sheet_protection(handle_, sheet, &p);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object pr = Napi::Object::New(env);
  pr.Set("enabled", JsBool(env, p.enabled));
  pr.Set("algorithmName", JsStringOrNull(env, p.algorithm_name));
  pr.Set("hashValue", JsStringOrNull(env, p.hash_value));
  pr.Set("saltValue", JsStringOrNull(env, p.salt_value));
  pr.Set("spinCount", Napi::Number::New(env, p.spin_count));
  pr.Set("legacyPassword", JsStringOrNull(env, p.legacy_password));
  pr.Set("sheet", JsBool(env, p.sheet));
  pr.Set("objects", JsBool(env, p.objects));
  pr.Set("scenarios", JsBool(env, p.scenarios));
  pr.Set("formatCells", JsBool(env, p.format_cells));
  pr.Set("formatColumns", JsBool(env, p.format_columns));
  pr.Set("formatRows", JsBool(env, p.format_rows));
  pr.Set("insertColumns", JsBool(env, p.insert_columns));
  pr.Set("insertRows", JsBool(env, p.insert_rows));
  pr.Set("insertHyperlinks", JsBool(env, p.insert_hyperlinks));
  pr.Set("deleteColumns", JsBool(env, p.delete_columns));
  pr.Set("deleteRows", JsBool(env, p.delete_rows));
  pr.Set("selectLockedCells", JsBool(env, p.select_locked_cells));
  pr.Set("selectUnlockedCells", JsBool(env, p.select_unlocked_cells));
  pr.Set("sort", JsBool(env, p.sort));
  pr.Set("autoFilter", JsBool(env, p.auto_filter));
  pr.Set("pivotTables", JsBool(env, p.pivot_tables));
  return pr;
}

Napi::Value Workbook::SetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object in = ArgObject(info, 1);

  // A missing or `null` string is no stored value. The std::string buffers
  // outlive the C ABI call, which borrows them.
  CheckedSpecReader reader(env);
  std::string algorithm_name;
  std::string hash_value;
  std::string salt_value;
  std::string legacy_password;
  const auto text = [&](const char* key, std::string* store) {
    return reader.String(in, key, store) ? store->c_str() : nullptr;
  };
  const auto flag = [&](const char* key) { return reader.Bool(in, key, false) ? 1 : 0; };
  fm_sheet_protection_t p = FM_STRUCT_INIT(fm_sheet_protection_t);
  p.algorithm_name = text("algorithmName", &algorithm_name);
  p.hash_value = text("hashValue", &hash_value);
  p.salt_value = text("saltValue", &salt_value);
  p.spin_count = reader.U32(in, "spinCount", 0U);
  p.legacy_password = text("legacyPassword", &legacy_password);
  p.enabled = flag("enabled");
  p.sheet = flag("sheet");
  p.objects = flag("objects");
  p.scenarios = flag("scenarios");
  p.format_cells = flag("formatCells");
  p.format_columns = flag("formatColumns");
  p.format_rows = flag("formatRows");
  p.insert_columns = flag("insertColumns");
  p.insert_rows = flag("insertRows");
  p.insert_hyperlinks = flag("insertHyperlinks");
  p.delete_columns = flag("deleteColumns");
  p.delete_rows = flag("deleteRows");
  p.select_locked_cells = flag("selectLockedCells");
  p.select_unlocked_cells = flag("selectUnlockedCells");
  p.sort = flag("sort");
  p.auto_filter = flag("autoFilter");
  p.pivot_tables = flag("pivotTables");
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_sheet_protection(handle_, sheet, &p);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::SetSheetZoom(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_zoom);
}

Napi::Value Workbook::SetSheetFreeze(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_freeze);
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
  out.Set("activeRow", Napi::Number::New(env, row));
  out.Set("activeCol", Napi::Number::New(env, col));
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
  const uint32_t row = reader.U32(selection, "activeRow", 0U);
  const uint32_t col = reader.U32(selection, "activeCol", 0U);
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

Napi::Value Workbook::SetSheetTabSelected(const Napi::CallbackInfo& info) {
  return InvokeSheetFlag(*this, info, &fm_sheet_set_tab_selected);
}

Napi::Value Workbook::SetSheetViewMode(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
    // An ordinal past int32 reaches the C ABI as -1, which it rejects like any unknown mode.
    const uint32_t mode = ArgU32(info, 1);
    return fm_sheet_set_view_mode(handle, sheet,
                                  mode > static_cast<uint32_t>(INT32_MAX) ? -1 : static_cast<int32_t>(mode));
  });
}

Napi::Value Workbook::GetSheetColumns(const Napi::CallbackInfo& info) {
  return SheetLayoutList(info, handle_, &fm_sheet_get_column_override_count, &fm_sheet_get_column_override_at,
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
  return SheetLayoutList(info, handle_, &fm_sheet_get_row_override_count, &fm_sheet_get_row_override_at,
                         &RowLayoutToJs);
}

Napi::Value Workbook::SetRowHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_height);
}

Napi::Value Workbook::SetRowAutoHeight(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_auto_height);
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
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  const fm_status_t rc = fm_sheet_get_format_defaults(handle_, ArgU32(info, 0), &d);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("defaultColWidth", JsNumberOrNull(env, d.has_default_col_width, d.default_col_width));
  out.Set("defaultRowHeight", JsNumberOrNull(env, d.has_default_row_height, d.default_row_height));
  out.Set("baseColWidth", Napi::Number::New(env, d.base_col_width));
  out.Set("customHeight", Napi::Boolean::New(env, d.custom_height != 0));
  out.Set("zeroHeight", Napi::Boolean::New(env, d.zero_height != 0));
  out.Set("effectiveDefaultColWidth", Napi::Number::New(env, d.effective_default_col_width));
  out.Set("effectiveDefaultRowHeight", Napi::Number::New(env, d.effective_default_row_height));
  return out;
}

Napi::Value Workbook::SetSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // A patch: an omitted key keeps the stored value, `null` clears a nullable one.
  const Napi::Object spec = ArgObject(info, 1);
  fm_sheet_format_defaults_t d = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  const fm_status_t read_rc = fm_sheet_get_format_defaults(handle_, ArgU32(info, 0), &d);
  if (read_rc != 0) {
    return ThrowStatus(env, read_rc);
  }
  CheckedSpecReader reader(env);
  const auto optional_number = [&](const char* key, int32_t* has, double* value) {
    const Napi::Value v = reader.Get(spec, key);
    switch (ClassifyPatchKey(reader, v, key, true)) {
      case PatchKey::kClear:
        *has = 0;
        *value = 0.0;
        break;
      case PatchKey::kSet:
        *has = 1;
        *value = reader.Double(v, key, 0.0);
        break;
      case PatchKey::kKeep:
        break;
    }
  };
  optional_number("defaultColWidth", &d.has_default_col_width, &d.default_col_width);
  optional_number("defaultRowHeight", &d.has_default_row_height, &d.default_row_height);
  if (ClassifyPatchKey(reader, reader.Get(spec, "baseColWidth"), "baseColWidth", false) == PatchKey::kSet) {
    d.base_col_width = reader.Double(spec, "baseColWidth", d.base_col_width);
  }
  if (ClassifyPatchKey(reader, reader.Get(spec, "customHeight"), "customHeight", false) == PatchKey::kSet) {
    d.custom_height = reader.Bool(spec, "customHeight", false) ? 1 : 0;
  }
  if (ClassifyPatchKey(reader, reader.Get(spec, "zeroHeight"), "zeroHeight", false) == PatchKey::kSet) {
    d.zero_height = reader.Bool(spec, "zeroHeight", false) ? 1 : 0;
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_format_defaults(handle_, ArgU32(info, 0), &d), env.Undefined());
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
  fm_rect_pt_t rect = FM_STRUCT_INIT(fm_rect_pt_t);
  const fm_status_t rc = fm_sheet_get_cell_rect_pt(handle_, ArgU32(info, 0), &cells, ArgMode(info, 2), &rect);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("x", Napi::Number::New(env, rect.x));
  out.Set("y", Napi::Number::New(env, rect.y));
  out.Set("width", Napi::Number::New(env, rect.width));
  out.Set("height", Napi::Number::New(env, rect.height));
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
  fm_width_model_t model = FM_STRUCT_INIT(fm_width_model_t);
  const fm_status_t rc = fm_sheet_get_width_model(handle_, ArgU32(info, 0), ArgMode(info, 1), &model);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("pointsPerChar", Napi::Number::New(env, model.points_per_char));
  out.Set("paddingPt", Napi::Number::New(env, model.padding_pt));
  out.Set("normalFontSize", Napi::Number::New(env, model.normal_font_size));
  out.Set("calibrated", Napi::Boolean::New(env, model.calibrated != 0));
  out.Set("normalFontName", JsString(env, model.normal_font_name));
  out.Set("platform", JsString(env, model.platform));
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
    return fm_sheet_set_row_hidden(handle, sheet, first, last, ArgFlag(info, 3));
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
    return fm_sheet_set_row_collapsed(handle, sheet, first, last, ArgFlag(info, 3));
  });
}

Napi::Value Workbook::SetRowStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_row_style);
}

Napi::Value Workbook::ClearRowStyle(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_row_style);
}

namespace {

Napi::Value RangeStyleBasesResult(Napi::Env env, fm_workbook_t* handle, uint32_t sheet, const fm_cell_rect_t& rect,
                                  const std::vector<uint32_t>& rows, const std::vector<uint32_t>& cols) {
  size_t len = 0;
  fm_status_t rc = fm_sheet_get_range_style_bases_excluding(handle, sheet, &rect, rows.data(), cols.data(), rows.size(),
                                                            nullptr, 0, &len);
  std::vector<uint32_t> xfs(len);
  if (rc == 0 && len > 0) {
    rc = fm_sheet_get_range_style_bases_excluding(handle, sheet, &rect, rows.data(), cols.data(), rows.size(),
                                                  xfs.data(), xfs.size(), &len);
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
  return RangeStyleBasesResult(env, handle_, ArgU32(info, 0), rect, {}, {});
}

Napi::Value Workbook::RangeStyleBasesExcluding(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const auto refused = [&env](const char* message) {
    Napi::TypeError::New(env, message).ThrowAsJavaScriptException();
    return env.Undefined();
  };
  Napi::Array rows_arr;
  Napi::Array cols_arr;
  if (!ArrayArg(info, 2, &rows_arr) || !ArrayArg(info, 3, &cols_arr)) {
    return refused(
        "getRangeStyleBasesExcluding expects (sheet, range:CellRect, excludedRows:number[], excludedCols:number[])");
  }
  if (rows_arr.Length() != cols_arr.Length()) {
    return refused("getRangeStyleBasesExcluding: excludedRows and excludedCols must have the same length");
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
  return RangeStyleBasesResult(env, handle_, ArgU32(info, 0), rect, rows, cols);
}

Napi::Value Workbook::RemapRangeStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array from_arr;
  Napi::Array to_arr;
  if (!ArrayArg(info, 2, &from_arr) || !ArrayArg(info, 3, &to_arr)) {
    Napi::TypeError::New(env, "remapRangeStyle expects (sheet, range:CellRect, fromXfs:number[], toXfs:number[])")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  if (from_arr.Length() != to_arr.Length()) {
    Napi::TypeError::New(env, "remapRangeStyle: fromXfs and toXfs must have the same length")
        .ThrowAsJavaScriptException();
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
  const fm_status_t rc = fm_sheet_remap_range_style(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &rect,
                                                    from.data(), to.data(), from.size());
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
      fm_sheet_snapshot_range_style(handle_, static_cast<std::size_t>(ArgU32(info, 0)), &rect, &bytes, &len);
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
    Napi::TypeError::New(env, "restoreRangeStyle expects (sheet:number, bytes:Uint8Array)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const fm_status_t rc =
      fm_sheet_restore_range_style(handle_, static_cast<std::size_t>(ArgU32(info, 0)), u8.Data(), u8.ElementLength());
  return ValueOrThrow(env, rc, env.Undefined());
}

}  // namespace formulon_node
