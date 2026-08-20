// Sheet view / layout bindings (zoom, freeze, tab visibility, column /
// row layout) and the per-sheet UI feature surfaces: merges,
// comments, hyperlinks, and data validations.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Sheet view / layout --------------------------------------------

namespace {

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

// ---- Merges ---------------------------------------------------------

Napi::Value Workbook::AddMerge(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_merge_range m{};
  if (info.Length() > 1 && info[1].IsObject()) {
    Napi::Object range = info[1].As<Napi::Object>();
    m.first_row = range.Get("firstRow").ToNumber().Uint32Value();
    m.last_row = range.Get("lastRow").ToNumber().Uint32Value();
    m.first_col = range.Get("firstCol").ToNumber().Uint32Value();
    m.last_col = range.Get("lastCol").ToNumber().Uint32Value();
  }
  fm_status_t rc = fm_sheet_add_merge(handle_, sheet, m);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveMerge(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_merge_range m{};
  if (info.Length() > 1 && info[1].IsObject()) {
    Napi::Object range = info[1].As<Napi::Object>();
    m.first_row = range.Get("firstRow").ToNumber().Uint32Value();
    m.last_row = range.Get("lastRow").ToNumber().Uint32Value();
    m.first_col = range.Get("firstCol").ToNumber().Uint32Value();
    m.last_col = range.Get("lastCol").ToNumber().Uint32Value();
  }
  fm_status_t rc = fm_sheet_remove_merge(handle_, sheet, m);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveMergeAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_remove_merge_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearMerges(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_clear_merges(handle_, sheet);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetMerges(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_merge_count(handle_, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_merge_range m{};
    rc = fm_sheet_get_merge_at(handle_, sheet, i, &m);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("firstRow", Napi::Number::New(env, m.first_row));
    item.Set("lastRow", Napi::Number::New(env, m.last_row));
    item.Set("firstCol", Napi::Number::New(env, m.first_col));
    item.Set("lastCol", Napi::Number::New(env, m.last_col));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return FinishListResult(env, arr, 0);
}

// ---- Comments -------------------------------------------------------

Napi::Value Workbook::GetComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return env.Null();
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_comment c{};
  if (fm_sheet_get_comment_at(handle_, sheet, row, col, &c) != 0) {
    return env.Null();
  }
  Napi::Object o = Napi::Object::New(env);
  o.Set("author", Napi::String::New(env, c.author != nullptr ? c.author : ""));
  o.Set("text", Napi::String::New(env, c.text != nullptr ? c.text : ""));
  return o;
}

Napi::Value Workbook::GetCommentResult(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("comment", env.Null());
    return out;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_comment c{};
  const fm_status_t rc = fm_sheet_get_comment_at(handle_, sheet, row, col, &c);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("comment", env.Null());
    return out;
  }
  Napi::Object comment = Napi::Object::New(env);
  comment.Set("author", Napi::String::New(env, c.author != nullptr ? c.author : ""));
  comment.Set("text", Napi::String::New(env, c.text != nullptr ? c.text : ""));
  out.Set("status", MakeOkStatus(env));
  out.Set("comment", comment);
  return out;
}

Napi::Value Workbook::GetComments(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_comment_count(handle_, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_comment c{};
    rc = fm_sheet_get_comment_at_index(handle_, sheet, i, &c);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("row", Napi::Number::New(env, c.row));
    item.Set("col", Napi::Number::New(env, c.col));
    item.Set("author", Napi::String::New(env, c.author != nullptr ? c.author : ""));
    item.Set("text", Napi::String::New(env, c.text != nullptr ? c.text : ""));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return FinishListResult(env, arr, 0);
}

Napi::Value Workbook::SetComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const std::string author = ArgString(info, 3);
  const std::string text = ArgString(info, 4);
  // Empty strings turn into NULL in the C ABI to remove an entry,
  // matching the embind binding's contract.
  const char* author_c = author.empty() ? nullptr : author.c_str();
  const char* text_c = text.empty() ? nullptr : text.c_str();
  fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
  return MakeStatus(env, rc);
}

// ---- Hyperlinks -----------------------------------------------------

Napi::Value Workbook::AddHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 7 || !info[0].IsNumber() || !info[1].IsNumber() || !info[2].IsNumber() || !info[3].IsString() ||
      !info[4].IsString() || !info[5].IsString() || !info[6].IsString()) {
    Napi::TypeError::New(env,
                         "addHyperlink expects (sheet:number, row:number, col:number, "
                         "target:string, display:string, tooltip:string, location:string)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  // Keep the std::string buffers alive until after the C ABI call so the
  // borrowed `const char*` pointers in `fm_hyperlink` stay valid.
  const std::string target = ArgString(info, 3);
  const std::string display = ArgString(info, 4);
  const std::string tooltip = ArgString(info, 5);
  const std::string location = ArgString(info, 6);
  fm_hyperlink hl{};
  hl.row = row;
  hl.col = col;
  hl.last_row = row;
  hl.last_col = col;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  fm_status_t rc = fm_sheet_add_hyperlink(handle_, sheet, hl);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::AddHyperlinkRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 9 || !info[0].IsNumber() || !info[1].IsNumber() || !info[2].IsNumber() || !info[3].IsNumber() ||
      !info[4].IsNumber() || !info[5].IsString() || !info[6].IsString() || !info[7].IsString() || !info[8].IsString()) {
    Napi::TypeError::New(env,
                         "addHyperlinkRange expects (sheet:number, row:number, col:number, "
                         "lastRow:number, lastCol:number, target:string, display:string, "
                         "tooltip:string, location:string)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t last_row = ArgU32(info, 3);
  const uint32_t last_col = ArgU32(info, 4);
  const std::string target = ArgString(info, 5);
  const std::string display = ArgString(info, 6);
  const std::string tooltip = ArgString(info, 7);
  const std::string location = ArgString(info, 8);
  fm_hyperlink hl{};
  hl.row = row;
  hl.col = col;
  hl.last_row = last_row;
  hl.last_col = last_col;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  fm_status_t rc = fm_sheet_add_hyperlink(handle_, sheet, hl);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::GetHyperlinks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_hyperlink_count(handle_, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_hyperlink h{};
    rc = fm_sheet_get_hyperlink_at(handle_, sheet, i, &h);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("row", Napi::Number::New(env, h.row));
    item.Set("col", Napi::Number::New(env, h.col));
    item.Set("lastRow", Napi::Number::New(env, h.last_row));
    item.Set("lastCol", Napi::Number::New(env, h.last_col));
    item.Set("target", Napi::String::New(env, h.target != nullptr ? h.target : ""));
    item.Set("location", Napi::String::New(env, h.location != nullptr ? h.location : ""));
    item.Set("display", Napi::String::New(env, h.display != nullptr ? h.display : ""));
    item.Set("tooltip", Napi::String::New(env, h.tooltip != nullptr ? h.tooltip : ""));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return FinishListResult(env, arr, 0);
}

Napi::Value Workbook::RemoveHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_status_t rc = fm_sheet_remove_hyperlink(handle_, sheet, row, col);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveHyperlinkAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_remove_hyperlink_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearHyperlinks(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_clear_hyperlinks(handle_, sheet);
  return MakeStatus(env, rc);
}

// ---- Validations ----------------------------------------------------

Napi::Value Workbook::GetValidations(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_validation_count(handle_, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::size_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_data_validation v{};
    rc = fm_sheet_get_validation_at(handle_, sheet, i, &v);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    Napi::Object item = Napi::Object::New(env);
    Napi::Array ranges = Napi::Array::New(env);
    for (uint32_t r = 0; r < v.range_count; ++r) {
      Napi::Object rng = Napi::Object::New(env);
      rng.Set("firstRow", Napi::Number::New(env, v.ranges[r].first_row));
      rng.Set("lastRow", Napi::Number::New(env, v.ranges[r].last_row));
      rng.Set("firstCol", Napi::Number::New(env, v.ranges[r].first_col));
      rng.Set("lastCol", Napi::Number::New(env, v.ranges[r].last_col));
      ranges.Set(r, rng);
    }
    item.Set("ranges", ranges);
    item.Set("type", Napi::Number::New(env, v.type));
    item.Set("op", Napi::Number::New(env, v.op));
    item.Set("errorStyle", Napi::Number::New(env, v.error_style));
    item.Set("allowBlank", Napi::Boolean::New(env, v.allow_blank != 0));
    item.Set("showInputMessage", Napi::Boolean::New(env, v.show_input_message != 0));
    item.Set("showErrorMessage", Napi::Boolean::New(env, v.show_error_message != 0));
    item.Set("showDropDown", Napi::Boolean::New(env, v.show_dropdown != 0));
    item.Set("formula1", Napi::String::New(env, v.formula1 != nullptr ? v.formula1 : ""));
    item.Set("formula2", Napi::String::New(env, v.formula2 != nullptr ? v.formula2 : ""));
    item.Set("errorTitle", Napi::String::New(env, v.error_title != nullptr ? v.error_title : ""));
    item.Set("errorMessage", Napi::String::New(env, v.error_message != nullptr ? v.error_message : ""));
    item.Set("promptTitle", Napi::String::New(env, v.prompt_title != nullptr ? v.prompt_title : ""));
    item.Set("promptMessage", Napi::String::New(env, v.prompt_message != nullptr ? v.prompt_message : ""));
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return FinishListResult(env, arr, 0);
}

Napi::Value Workbook::AddValidation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsObject()) {
    Napi::TypeError::New(env, "addValidation expects (sheet:number, validation:object)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  Napi::Object v = info[1].As<Napi::Object>();

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  std::vector<fm_merge_range> ranges_buf;
  if (v.Has("ranges")) {
    Napi::Value ranges_js = v.Get("ranges");
    if (ranges_js.IsArray()) {
      Napi::Array ranges_arr = ranges_js.As<Napi::Array>();
      const uint32_t n = ranges_arr.Length();
      ranges_buf.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        Napi::Value rng_v = ranges_arr.Get(i);
        if (!rng_v.IsObject()) {
          continue;
        }
        Napi::Object rng = rng_v.As<Napi::Object>();
        fm_merge_range m{};
        m.first_row = rng.Get("firstRow").ToNumber().Uint32Value();
        m.last_row = rng.Get("lastRow").ToNumber().Uint32Value();
        m.first_col = rng.Get("firstCol").ToNumber().Uint32Value();
        m.last_col = rng.Get("lastCol").ToNumber().Uint32Value();
        ranges_buf.push_back(m);
      }
    }
  }
  auto pull_string = [&](const char* key) -> std::string {
    if (!v.Has(key)) {
      return std::string();
    }
    Napi::Value f = v.Get(key);
    if (f.IsUndefined() || f.IsNull()) {
      return std::string();
    }
    return f.ToString().Utf8Value();
  };
  auto pull_u8 = [&](const char* key) -> uint8_t {
    if (!v.Has(key)) {
      return 0;
    }
    Napi::Value f = v.Get(key);
    if (f.IsUndefined() || f.IsNull()) {
      return 0;
    }
    return static_cast<uint8_t>(f.ToNumber().Uint32Value() & 0xFFU);
  };
  auto pull_bool = [&](const char* key, bool dflt) -> bool {
    if (!v.Has(key)) {
      return dflt;
    }
    Napi::Value f = v.Get(key);
    if (f.IsUndefined() || f.IsNull()) {
      return dflt;
    }
    return f.ToBoolean().Value();
  };
  const std::string formula1 = pull_string("formula1");
  const std::string formula2 = pull_string("formula2");
  const std::string error_title = pull_string("errorTitle");
  const std::string error_message = pull_string("errorMessage");
  const std::string prompt_title = pull_string("promptTitle");
  const std::string prompt_message = pull_string("promptMessage");

  fm_data_validation dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = static_cast<uint32_t>(ranges_buf.size());
  dv.type = pull_u8("type");
  dv.op = pull_u8("op");
  dv.error_style = pull_u8("errorStyle");
  dv.allow_blank = pull_bool("allowBlank", false) ? 1 : 0;
  dv.show_input_message = pull_bool("showInputMessage", false) ? 1 : 0;
  dv.show_error_message = pull_bool("showErrorMessage", false) ? 1 : 0;
  dv.show_dropdown = pull_bool("showDropDown", true) ? 1 : 0;
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, dv);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveValidationAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  fm_status_t rc = fm_sheet_remove_validation_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearValidations(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_clear_validations(handle_, sheet);
  return MakeStatus(env, rc);
}

}  // namespace formulon_node
