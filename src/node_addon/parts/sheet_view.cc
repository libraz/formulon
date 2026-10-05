// Sheet view / layout bindings: zoom, freeze, tab visibility, protection,
// and column / row layout.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Sheet view / layout --------------------------------------------

namespace {

// Reads a geometry mode ordinal; an absent argument is the display mode.
int32_t ArgMode(const Napi::CallbackInfo& info, size_t idx) {
  return idx < info.Length() ? info[idx].ToNumber().Int32Value() : static_cast<int32_t>(FM_GEOMETRY_DISPLAY);
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
  return view;
}

}  // namespace

Napi::Value Workbook::GetSheetView(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("view", DefaultSheetView(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_sheet_view_t v{};
  fm_status_t rc = fm_sheet_get_view(handle_, sheet, &v);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("view", DefaultSheetView(env));
    return out;
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
  view.Set("viewMode", Napi::String::New(env, v.view_mode != nullptr ? v.view_mode : ""));
  out.Set("status", MakeOkStatus(env));
  out.Set("view", view);
  return out;
}

Napi::Value Workbook::GetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
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
  pr.Set("algorithmName", Napi::String::New(env, p.algorithm_name != nullptr ? p.algorithm_name : ""));
  pr.Set("hashValue", Napi::String::New(env, p.hash_value != nullptr ? p.hash_value : ""));
  pr.Set("saltValue", Napi::String::New(env, p.salt_value != nullptr ? p.salt_value : ""));
  pr.Set("spinCount", Napi::Number::New(env, p.spin_count));
  pr.Set("legacyPassword", Napi::String::New(env, p.legacy_password != nullptr ? p.legacy_password : ""));
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
  out.Set("status", MakeStatus(env, rc));
  out.Set("protection", pr);
  return out;
}

Napi::Value Workbook::SetSheetProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsObject()) {
    Napi::TypeError::New(env, "setSheetProtection expects (sheet:number, protection:object)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object in = info[1].As<Napi::Object>();

  // Keep the std::string buffers alive until after the C ABI call so the
  // borrowed `const char*` fields stay valid.
  auto pull_string = [&](const char* key) -> std::string {
    if (!SpecHas(in, key)) {
      return std::string();
    }
    return in.Get(key).ToString().Utf8Value();
  };
  const std::string algorithm_name = pull_string("algorithmName");
  const std::string hash_value = pull_string("hashValue");
  const std::string salt_value = pull_string("saltValue");
  const std::string legacy_password = pull_string("legacyPassword");

  fm_sheet_protection_t p{};
  p.enabled = SpecPullInt32(in, "enabled", 0);
  p.algorithm_name = algorithm_name.c_str();
  p.hash_value = hash_value.c_str();
  p.salt_value = salt_value.c_str();
  p.spin_count = SpecPullU32(in, "spinCount", 0U);
  p.legacy_password = legacy_password.c_str();
  p.sheet = SpecPullInt32(in, "sheet", 0);
  p.objects = SpecPullInt32(in, "objects", 0);
  p.scenarios = SpecPullInt32(in, "scenarios", 0);
  p.format_cells = SpecPullInt32(in, "formatCells", 0);
  p.format_columns = SpecPullInt32(in, "formatColumns", 0);
  p.format_rows = SpecPullInt32(in, "formatRows", 0);
  p.insert_columns = SpecPullInt32(in, "insertColumns", 0);
  p.insert_rows = SpecPullInt32(in, "insertRows", 0);
  p.insert_hyperlinks = SpecPullInt32(in, "insertHyperlinks", 0);
  p.delete_columns = SpecPullInt32(in, "deleteColumns", 0);
  p.delete_rows = SpecPullInt32(in, "deleteRows", 0);
  p.select_locked_cells = SpecPullInt32(in, "selectLockedCells", 0);
  p.select_unlocked_cells = SpecPullInt32(in, "selectUnlockedCells", 0);
  p.sort = SpecPullInt32(in, "sort", 0);
  p.auto_filter = SpecPullInt32(in, "autoFilter", 0);
  p.pivot_tables = SpecPullInt32(in, "pivotTables", 0);
  if (env.IsExceptionPending()) {
    // A malformed field left a pending JS exception (see SpecPullInt32 /
    // SpecPullU32): stop before the C ABI call commits a default value
    // for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_protection(handle_, sheet, &p);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetZoom(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t zoom = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_set_zoom(handle_, sheet, zoom);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetFreeze(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t freeze_rows = ArgU32(info, 1);
  const uint32_t freeze_cols = ArgU32(info, 2);
  fm_status_t rc = fm_sheet_set_freeze(handle_, sheet, freeze_rows, freeze_cols);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetTabHidden(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool hidden = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_tab_hidden(handle_, sheet, hidden ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetVisibility(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  // Raw ordinal: the C ABI rejects an unknown value, so coercing it to a
  // narrower type here would turn a caller's mistake into a silent state.
  const std::int32_t visibility = info.Length() > 1 ? info[1].ToNumber().Int32Value() : 0;
  fm_status_t rc = fm_sheet_set_visibility(handle_, sheet, visibility);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetShowGridLines(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool show = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_show_grid_lines(handle_, sheet, show ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetShowRowColHeaders(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool show = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_show_row_col_headers(handle_, sheet, show ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetShowZeros(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool show = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_show_zeros(handle_, sheet, show ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetRightToLeft(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool right_to_left = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_right_to_left(handle_, sheet, right_to_left ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetTabSelected(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const bool selected = ArgBool(info, 1);
  fm_status_t rc = fm_sheet_set_tab_selected(handle_, sheet, selected ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetSheetViewMode(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::string mode = ArgString(info, 1);
  fm_status_t rc = fm_sheet_set_view_mode(handle_, sheet, mode.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetSheetColumns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("columns", Napi::Array::New(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_column_count(handle_, sheet, &count);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("columns", Napi::Array::New(env));
    return out;
  }
  Napi::Array arr = Napi::Array::New(env, count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_column_layout_t entry{};
    if (fm_sheet_get_column(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    Napi::Object col = Napi::Object::New(env);
    col.Set("first", Napi::Number::New(env, entry.first));
    col.Set("last", Napi::Number::New(env, entry.last));
    col.Set("width", Napi::Number::New(env, entry.width));
    col.Set("hidden", Napi::Number::New(env, entry.hidden));
    col.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
    // The C getter reports logical presence, including legacy aggregate
    // layouts with a non-zero width but a clear raw presence bit. Keep the
    // binding defensive in case an older ABI implementation is loaded.
    col.Set("hasWidth", Napi::Number::New(env, entry.has_width || entry.width != 0.0));
    col.Set("hasStyle", Napi::Number::New(env, entry.has_style));
    col.Set("styleXf", Napi::Number::New(env, entry.style_xf));
    arr.Set(static_cast<uint32_t>(emitted), col);
    ++emitted;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("columns", arr);
  return out;
}

Napi::Value Workbook::SetColumnWidth(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  const double width = ArgDouble(info, 3);
  fm_status_t rc = fm_sheet_set_column_width(handle_, sheet, first, last, width);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetColumnHidden(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  const bool hidden = ArgBool(info, 3);
  fm_status_t rc = fm_sheet_set_column_hidden(handle_, sheet, first, last, hidden ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetColumnOutline(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  uint32_t level = ArgU32(info, 3);
  if (level > 255U) {
    level = 255U;
  }
  fm_status_t rc = fm_sheet_set_column_outline(handle_, sheet, first, last, static_cast<uint8_t>(level));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetSheetRowOverrides(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("rows", Napi::Array::New(env));
    return out;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  fm_status_t rc = fm_sheet_get_row_override_count(handle_, sheet, &count);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("rows", Napi::Array::New(env));
    return out;
  }
  Napi::Array arr = Napi::Array::New(env, count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_row_layout_t entry{};
    if (fm_sheet_get_row_override(handle_, sheet, i, &entry) != 0) {
      continue;
    }
    Napi::Object row = Napi::Object::New(env);
    row.Set("row", Napi::Number::New(env, entry.row));
    row.Set("height", Napi::Number::New(env, entry.height));
    row.Set("hasHeight", Napi::Number::New(env, entry.has_height));
    row.Set("customHeight", Napi::Number::New(env, entry.custom_height));
    row.Set("hidden", Napi::Number::New(env, entry.hidden));
    row.Set("outlineLevel", Napi::Number::New(env, static_cast<int32_t>(entry.outline_level)));
    row.Set("hasStyle", Napi::Number::New(env, entry.has_style));
    row.Set("styleXf", Napi::Number::New(env, entry.style_xf));
    arr.Set(static_cast<uint32_t>(emitted), row);
    ++emitted;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("rows", arr);
  return out;
}

Napi::Value Workbook::SetRowHeight(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const double height = ArgDouble(info, 2);
  fm_status_t rc = fm_sheet_set_row_height(handle_, sheet, row, height);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearRowHeight(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  return MakeStatus(env, fm_sheet_clear_row_height(handle_, sheet, ArgU32(info, 1)));
}

Napi::Value Workbook::GetSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_sheet_format_defaults d{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_format_defaults(handle_, ArgU32(info, 0), &d)
                                            : kBindingInvalidHandle;
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
  return out;
}

Napi::Value Workbook::SetSheetFormatDefaults(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    return MakeBindingArgumentError(env, "setSheetFormatDefaults expects (sheet:number, defaults:object)");
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  fm_sheet_format_defaults d{};
  d.default_col_width = SpecPullDouble(spec, "defaultColWidth", 0.0);
  d.default_row_height = SpecPullDouble(spec, "defaultRowHeight", 0.0);
  d.base_col_width = SpecPullDouble(spec, "baseColWidth", 8.0);
  d.has_default_col_width = SpecPullBool(spec, "hasDefaultColWidth", SpecHas(spec, "defaultColWidth")) ? 1 : 0;
  d.has_default_row_height = SpecPullBool(spec, "hasDefaultRowHeight", SpecHas(spec, "defaultRowHeight")) ? 1 : 0;
  return MakeStatus(env, fm_sheet_set_format_defaults(handle_, ArgU32(info, 0), &d));
}

// ---- Geometry -------------------------------------------------------

Napi::Value Workbook::GetCellRectPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_rect_pt rect{};
  Napi::Object status = Napi::Object::New(env);
  if (handle_ == nullptr) {
    status = NullHandleError(env);
  } else if (info.Length() < 2 || !info[1].IsObject()) {
    status = MakeBindingArgumentError(env, "getCellRectPt expects (sheet:number, range:object, mode:number)");
  } else {
    const Napi::Object range = info[1].As<Napi::Object>();
    const fm_status_t rc = fm_sheet_cell_rect_pt(
        handle_, ArgU32(info, 0), SpecPullU32(range, "firstRow", 0U), SpecPullU32(range, "firstCol", 0U),
        SpecPullU32(range, "lastRow", 0U), SpecPullU32(range, "lastCol", 0U), ArgMode(info, 2), &rect);
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
  out.Set("normalFontName", Napi::String::New(env, model.normal_font_name != nullptr ? model.normal_font_name : ""));
  out.Set("platform", Napi::String::New(env, model.platform != nullptr ? model.platform : ""));
  return out;
}

Napi::Value Workbook::ColumnCharsToPt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  double pt = 0.0;
  const fm_status_t rc = fm_sheet_column_chars_to_pt(handle_, ArgU32(info, 0), ArgMode(info, 1), ArgDouble(info, 2), &pt);
  return MakeNumberResult(env, rc, pt);
}

Napi::Value Workbook::ColumnPtToChars(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  double chars = 0.0;
  const fm_status_t rc =
      fm_sheet_column_pt_to_chars(handle_, ArgU32(info, 0), ArgMode(info, 1), ArgDouble(info, 2), &chars);
  return MakeNumberResult(env, rc, chars);
}

Napi::Value Workbook::SetRowHidden(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  const bool hidden = ArgBool(info, 2);
  fm_status_t rc = fm_sheet_set_row_hidden(handle_, sheet, row, hidden ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetRowOutline(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const uint32_t row = ArgU32(info, 1);
  uint32_t level = ArgU32(info, 2);
  if (level > 255U) {
    level = 255U;
  }
  fm_status_t rc = fm_sheet_set_row_outline(handle_, sheet, row, static_cast<uint8_t>(level));
  return MakeStatus(env, rc);
}

}  // namespace formulon_node
