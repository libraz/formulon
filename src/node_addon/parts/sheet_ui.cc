// Per-sheet UI feature bindings: merges, comments, hyperlinks, and data
// validations.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Merges ---------------------------------------------------------

namespace {

// Reads a `{ firstRow, lastRow, firstCol, lastCol }` range object.
fm_merge_range ReadMergeRange(const Napi::Object& range) {
  fm_merge_range m{};
  m.first_row = range.Get("firstRow").ToNumber().Uint32Value();
  m.last_row = range.Get("lastRow").ToNumber().Uint32Value();
  m.first_col = range.Get("firstCol").ToNumber().Uint32Value();
  m.last_col = range.Get("lastCol").ToNumber().Uint32Value();
  return m;
}

// Reads the range at `info[idx]`; a missing or non-object argument reads as all zeros.
fm_merge_range MergeRangeArg(const Napi::CallbackInfo& info, size_t idx) {
  return info.Length() > idx && info[idx].IsObject() ? ReadMergeRange(info[idx].As<Napi::Object>()) : fm_merge_range{};
}

}  // namespace

Napi::Value Workbook::AddMerge(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_add_merge(handle_, sheet, MergeRangeArg(info, 1));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveMerge(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  fm_status_t rc = fm_sheet_remove_merge(handle_, sheet, MergeRangeArg(info, 1));
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

Napi::Value Workbook::GetMergesInRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const fm_merge_range range = MergeRangeArg(info, 1);
  uint32_t total = 0;
  fm_status_t rc = fm_sheet_merges_in_range(handle_, sheet, range, nullptr, 0, &total);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::vector<fm_merge_range> merges(total);
  uint32_t again = 0;
  rc = fm_sheet_merges_in_range(handle_, sheet, range, merges.data(), total, &again);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  for (uint32_t i = 0; i < total; ++i) {
    Napi::Object item = Napi::Object::New(env);
    item.Set("firstRow", Napi::Number::New(env, merges[i].first_row));
    item.Set("lastRow", Napi::Number::New(env, merges[i].last_row));
    item.Set("firstCol", Napi::Number::New(env, merges[i].first_col));
    item.Set("lastCol", Napi::Number::New(env, merges[i].last_col));
    arr.Set(i, item);
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

namespace {

// Adds the hyperlink described by the already type-checked arguments
// `(sheet, row, col[, lastRow, lastCol], target, display, tooltip, location)`.
// Empty strings are forwarded as NULL.
fm_status_t AddHyperlinkFromArgs(fm_workbook_t* wb, const Napi::CallbackInfo& info, bool ranged) {
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  const uint32_t row = Workbook::ArgU32(info, 1);
  const uint32_t col = Workbook::ArgU32(info, 2);
  const uint32_t last_row = ranged ? Workbook::ArgU32(info, 3) : row;
  const uint32_t last_col = ranged ? Workbook::ArgU32(info, 4) : col;
  const std::size_t text_idx = ranged ? 5 : 3;
  // Keep the std::string buffers alive until after the C ABI call so the
  // borrowed `const char*` pointers in `fm_hyperlink` stay valid.
  const std::string target = Workbook::ArgString(info, text_idx);
  const std::string display = Workbook::ArgString(info, text_idx + 1);
  const std::string tooltip = Workbook::ArgString(info, text_idx + 2);
  const std::string location = Workbook::ArgString(info, text_idx + 3);
  fm_hyperlink hl{};
  hl.row = row;
  hl.col = col;
  hl.last_row = last_row;
  hl.last_col = last_col;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  return fm_sheet_add_hyperlink(wb, sheet, hl);
}

}  // namespace

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
  return MakeStatus(env, AddHyperlinkFromArgs(handle_, info, /*ranged=*/false));
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
  return MakeStatus(env, AddHyperlinkFromArgs(handle_, info, /*ranged=*/true));
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
        ranges_buf.push_back(ReadMergeRange(rng_v.As<Napi::Object>()));
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
