// Per-sheet UI feature bindings: merges, comments, hyperlinks, and data
// validations.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
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

// Enumerates a per-sheet list through its count / at-index pair into a
// `ListResult` array, stopping at the first failed element read.
template <typename T>
Napi::Value SheetListResult(const Napi::CallbackInfo& info, fm_workbook_t* handle,
                            fm_status_t (*count_fn)(fm_workbook_t*, uint32_t, uint32_t*),
                            fm_status_t (*at_fn)(fm_workbook_t*, uint32_t, uint32_t, T*),
                            Napi::Object (*to_js)(Napi::Env, const T&)) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  uint32_t count = 0;
  fm_status_t rc = count_fn(handle, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  for (uint32_t i = 0; i < count; ++i) {
    T entry{};
    rc = at_fn(handle, sheet, i, &entry);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    arr.Set(i, to_js(env, entry));
  }
  return FinishListResult(env, arr, 0);
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

namespace {

Napi::Object MergeToJs(Napi::Env env, const fm_merge_range& m) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("firstRow", Napi::Number::New(env, m.first_row));
  item.Set("lastRow", Napi::Number::New(env, m.last_row));
  item.Set("firstCol", Napi::Number::New(env, m.first_col));
  item.Set("lastCol", Napi::Number::New(env, m.last_col));
  return item;
}

}  // namespace

Napi::Value Workbook::GetMerges(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_merge_count, &fm_sheet_get_merge_at, &MergeToJs);
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

namespace {

Napi::Object CommentToJs(Napi::Env env, const fm_comment& c) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("row", Napi::Number::New(env, c.row));
  item.Set("col", Napi::Number::New(env, c.col));
  item.Set("author", Napi::String::New(env, c.author != nullptr ? c.author : ""));
  item.Set("text", Napi::String::New(env, c.text != nullptr ? c.text : ""));
  return item;
}

}  // namespace

Napi::Value Workbook::GetComments(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_comment_count, &fm_sheet_get_comment_at_index, &CommentToJs);
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

namespace {

Napi::Object HyperlinkToJs(Napi::Env env, const fm_hyperlink& h) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("row", Napi::Number::New(env, h.row));
  item.Set("col", Napi::Number::New(env, h.col));
  item.Set("lastRow", Napi::Number::New(env, h.last_row));
  item.Set("lastCol", Napi::Number::New(env, h.last_col));
  item.Set("target", Napi::String::New(env, h.target != nullptr ? h.target : ""));
  item.Set("location", Napi::String::New(env, h.location != nullptr ? h.location : ""));
  item.Set("display", Napi::String::New(env, h.display != nullptr ? h.display : ""));
  item.Set("tooltip", Napi::String::New(env, h.tooltip != nullptr ? h.tooltip : ""));
  return item;
}

}  // namespace

Napi::Value Workbook::GetHyperlinks(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_hyperlink_count, &fm_sheet_get_hyperlink_at, &HyperlinkToJs);
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

namespace {

Napi::Object ValidationToJs(Napi::Env env, const fm_data_validation& v) {
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
  return item;
}

}  // namespace

Napi::Value Workbook::GetValidations(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_validation_count, &fm_sheet_get_validation_at, &ValidationToJs);
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

// ---- Typed AutoFilter -----------------------------------------------

namespace {

constexpr uint32_t kAbsentDxfId = UINT32_MAX;

std::string PullString(const Napi::Object& spec, const char* key) {
  return SpecHas(spec, key) ? spec.Get(key).ToString().Utf8Value() : std::string();
}

Napi::Value MergeRangeToJs(Napi::Env env, const fm_merge_range& r) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("firstRow", Napi::Number::New(env, r.first_row));
  o.Set("lastRow", Napi::Number::New(env, r.last_row));
  o.Set("firstCol", Napi::Number::New(env, r.first_col));
  o.Set("lastCol", Napi::Number::New(env, r.last_col));
  return o;
}

fm_merge_range PullRange(const Napi::Object& spec, const char* key) {
  fm_merge_range m{};
  if (SpecHas(spec, key) && spec.Get(key).IsObject()) {
    m = ReadMergeRange(spec.Get(key).As<Napi::Object>());
  }
  return m;
}

Napi::Value CStr(Napi::Env env, const char* s) {
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value Num(Napi::Env env, double v) {
  return Napi::Number::New(env, v);
}

Napi::Value Bool(Napi::Env env, int32_t v) {
  return Napi::Boolean::New(env, v != 0);
}

Napi::Object FilterColumnToJs(Napi::Env env, const fm_filter_column& c) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("colId", Num(env, c.col_id));
  o.Set("hiddenButton", Bool(env, c.hidden_button));
  o.Set("showButton", Bool(env, c.show_button));
  o.Set("kind", Num(env, c.kind));
  o.Set("filterBlank", Bool(env, c.filter_blank));
  Napi::Array values = Napi::Array::New(env, c.value_count);
  for (uint32_t i = 0; i < c.value_count; ++i) {
    values.Set(i, CStr(env, c.values[i]));
  }
  o.Set("values", values);
  Napi::Array groups = Napi::Array::New(env, c.date_group_count);
  for (uint32_t i = 0; i < c.date_group_count; ++i) {
    const fm_date_group_item& g = c.date_groups[i];
    Napi::Object item = Napi::Object::New(env);
    item.Set("year", Num(env, g.year));
    item.Set("month", Num(env, g.month));
    item.Set("day", Num(env, g.day));
    item.Set("hour", Num(env, g.hour));
    item.Set("minute", Num(env, g.minute));
    item.Set("second", Num(env, g.second));
    item.Set("grouping", Num(env, g.grouping));
    groups.Set(i, item);
  }
  o.Set("dateGroups", groups);
  o.Set("customAnd", Bool(env, c.custom_and));
  o.Set("customCount", Num(env, c.custom_count));
  o.Set("op1", Num(env, c.op1));
  o.Set("val1", CStr(env, c.val1));
  o.Set("op2", Num(env, c.op2));
  o.Set("val2", CStr(env, c.val2));
  o.Set("top", Bool(env, c.top));
  o.Set("percent", Bool(env, c.percent));
  o.Set("hasFilterVal", Bool(env, c.has_filter_val));
  o.Set("topVal", Num(env, c.top_val));
  o.Set("filterVal", Num(env, c.filter_val));
  o.Set("dynamicType", Num(env, c.dynamic_type));
  o.Set("hasDynVal", Bool(env, c.has_dyn_val));
  o.Set("hasDynMaxVal", Bool(env, c.has_dyn_max_val));
  o.Set("dynVal", Num(env, c.dyn_val));
  o.Set("dynMaxVal", Num(env, c.dyn_max_val));
  o.Set("valIso", CStr(env, c.val_iso));
  o.Set("maxValIso", CStr(env, c.max_val_iso));
  o.Set("dxfId", Num(env, c.dxf_id));
  o.Set("cellColor", Bool(env, c.cell_color));
  o.Set("iconSet", Num(env, c.icon_set));
  o.Set("iconId", Num(env, c.icon_id));
  o.Set("hasIconId", Bool(env, c.has_icon_id));
  return o;
}

Napi::Object AutoFilterToJs(Napi::Env env, const fm_auto_filter& f) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("range", MergeRangeToJs(env, f.range));
  Napi::Array columns = Napi::Array::New(env, f.column_count);
  for (uint32_t i = 0; i < f.column_count; ++i) {
    columns.Set(i, FilterColumnToJs(env, f.columns[i]));
  }
  o.Set("columns", columns);
  if (f.has_sort == 0) {
    o.Set("sort", env.Null());
    return o;
  }
  Napi::Object sort = Napi::Object::New(env);
  sort.Set("ref", MergeRangeToJs(env, f.sort_ref));
  sort.Set("columnSort", Bool(env, f.column_sort));
  sort.Set("caseSensitive", Bool(env, f.case_sensitive));
  sort.Set("sortMethod", Num(env, f.sort_method));
  Napi::Array conditions = Napi::Array::New(env, f.condition_count);
  for (uint32_t i = 0; i < f.condition_count; ++i) {
    const fm_sort_condition& c = f.conditions[i];
    Napi::Object item = Napi::Object::New(env);
    item.Set("ref", MergeRangeToJs(env, c.ref));
    item.Set("descending", Bool(env, c.descending));
    item.Set("sortBy", Num(env, c.sort_by));
    item.Set("customList", CStr(env, c.custom_list));
    item.Set("dxfId", Num(env, c.dxf_id));
    item.Set("hasDxfId", Bool(env, c.has_dxf_id));
    item.Set("iconSet", Num(env, c.icon_set));
    item.Set("iconId", Num(env, c.icon_id));
    item.Set("hasIconId", Bool(env, c.has_icon_id));
    conditions.Set(i, item);
  }
  sort.Set("conditions", conditions);
  o.Set("sort", sort);
  return o;
}

// Owns the strings and arrays an `fm_auto_filter` borrows from.
struct AutoFilterInput {
  fm_auto_filter filter{};
  std::vector<fm_filter_column> columns;
  std::vector<fm_sort_condition> conditions;
  // A deque never relocates its elements, so the `c_str()` views stay valid.
  std::deque<std::string> strings;
  std::vector<std::vector<const char*>> value_ptrs;
  std::vector<std::vector<fm_date_group_item>> groups;

  const char* Keep(std::string s) {
    strings.push_back(std::move(s));
    return strings.back().c_str();
  }
};

void ReadFilterColumn(const Napi::Object& spec, AutoFilterInput& in, fm_filter_column& c) {
  c.col_id = SpecPullU32(spec, "colId", 0U);
  c.hidden_button = SpecPullBool(spec, "hiddenButton", false) ? 1 : 0;
  c.show_button = SpecPullBool(spec, "showButton", true) ? 1 : 0;
  c.kind = SpecPullInt32(spec, "kind", 0);
  c.filter_blank = SpecPullBool(spec, "filterBlank", false) ? 1 : 0;
  in.value_ptrs.emplace_back();
  std::vector<const char*>& ptrs = in.value_ptrs.back();
  if (spec.Has("values") && spec.Get("values").IsArray()) {
    const Napi::Array values = spec.Get("values").As<Napi::Array>();
    for (uint32_t i = 0; i < values.Length(); ++i) {
      ptrs.push_back(in.Keep(values.Get(i).ToString().Utf8Value()));
    }
  }
  c.values = ptrs.empty() ? nullptr : ptrs.data();
  c.value_count = static_cast<uint32_t>(ptrs.size());
  in.groups.emplace_back();
  std::vector<fm_date_group_item>& groups = in.groups.back();
  if (spec.Has("dateGroups") && spec.Get("dateGroups").IsArray()) {
    const Napi::Array arr = spec.Get("dateGroups").As<Napi::Array>();
    for (uint32_t i = 0; i < arr.Length(); ++i) {
      if (!arr.Get(i).IsObject()) {
        continue;
      }
      const Napi::Object g = arr.Get(i).As<Napi::Object>();
      fm_date_group_item item{};
      item.year = static_cast<uint16_t>(SpecPullU32(g, "year", 0U));
      item.month = static_cast<uint8_t>(SpecPullU32(g, "month", 0U));
      item.day = static_cast<uint8_t>(SpecPullU32(g, "day", 0U));
      item.hour = static_cast<uint8_t>(SpecPullU32(g, "hour", 0U));
      item.minute = static_cast<uint8_t>(SpecPullU32(g, "minute", 0U));
      item.second = static_cast<uint8_t>(SpecPullU32(g, "second", 0U));
      item.grouping = static_cast<uint8_t>(SpecPullU32(g, "grouping", 0U));
      groups.push_back(item);
    }
  }
  c.date_groups = groups.empty() ? nullptr : groups.data();
  c.date_group_count = static_cast<uint32_t>(groups.size());
  c.custom_and = SpecPullBool(spec, "customAnd", false) ? 1 : 0;
  c.custom_count = SpecPullInt32(spec, "customCount", 0);
  c.op1 = SpecPullInt32(spec, "op1", 0);
  c.val1 = in.Keep(PullString(spec, "val1"));
  c.op2 = SpecPullInt32(spec, "op2", 0);
  c.val2 = in.Keep(PullString(spec, "val2"));
  c.top = SpecPullBool(spec, "top", false) ? 1 : 0;
  c.percent = SpecPullBool(spec, "percent", false) ? 1 : 0;
  c.has_filter_val = SpecPullBool(spec, "hasFilterVal", false) ? 1 : 0;
  c.top_val = SpecPullDouble(spec, "topVal", 0.0);
  c.filter_val = SpecPullDouble(spec, "filterVal", 0.0);
  c.dynamic_type = SpecPullInt32(spec, "dynamicType", 0);
  c.has_dyn_val = SpecPullBool(spec, "hasDynVal", false) ? 1 : 0;
  c.has_dyn_max_val = SpecPullBool(spec, "hasDynMaxVal", false) ? 1 : 0;
  c.dyn_val = SpecPullDouble(spec, "dynVal", 0.0);
  c.dyn_max_val = SpecPullDouble(spec, "dynMaxVal", 0.0);
  c.val_iso = in.Keep(PullString(spec, "valIso"));
  c.max_val_iso = in.Keep(PullString(spec, "maxValIso"));
  c.dxf_id = SpecPullU32(spec, "dxfId", kAbsentDxfId);
  c.cell_color = SpecPullBool(spec, "cellColor", false) ? 1 : 0;
  c.icon_set = SpecPullInt32(spec, "iconSet", 0);
  c.icon_id = SpecPullInt32(spec, "iconId", 0);
  c.has_icon_id = SpecPullBool(spec, "hasIconId", false) ? 1 : 0;
}

void ReadSortCondition(const Napi::Object& spec, AutoFilterInput& in, fm_sort_condition& c) {
  c.ref = PullRange(spec, "ref");
  c.descending = SpecPullBool(spec, "descending", false) ? 1 : 0;
  c.sort_by = SpecPullInt32(spec, "sortBy", 0);
  c.custom_list = in.Keep(PullString(spec, "customList"));
  c.dxf_id = SpecPullU32(spec, "dxfId", kAbsentDxfId);
  c.has_dxf_id = SpecPullBool(spec, "hasDxfId", false) ? 1 : 0;
  c.icon_set = SpecPullInt32(spec, "iconSet", 0);
  c.icon_id = SpecPullInt32(spec, "iconId", 0);
  c.has_icon_id = SpecPullBool(spec, "hasIconId", false) ? 1 : 0;
}

// Reads the `AutoFilter` record at `info[idx]`; false when it is not an object.
bool ReadAutoFilter(const Napi::CallbackInfo& info, size_t idx, AutoFilterInput& in) {
  if (info.Length() <= idx || !info[idx].IsObject()) {
    return false;
  }
  const Napi::Object spec = info[idx].As<Napi::Object>();
  fm_auto_filter& f = in.filter;
  f.range = PullRange(spec, "range");
  const Napi::Array columns = spec.Has("columns") && spec.Get("columns").IsArray()
                                  ? spec.Get("columns").As<Napi::Array>()
                                  : Napi::Array::New(info.Env());
  const Napi::Object sort = SpecHas(spec, "sort") && spec.Get("sort").IsObject() ? spec.Get("sort").As<Napi::Object>()
                                                                                 : Napi::Object::New(info.Env());
  const bool has_sort = SpecHas(spec, "sort") && spec.Get("sort").IsObject();
  const Napi::Array conditions = sort.Has("conditions") && sort.Get("conditions").IsArray()
                                     ? sort.Get("conditions").As<Napi::Array>()
                                     : Napi::Array::New(info.Env());
  for (uint32_t i = 0; i < columns.Length(); ++i) {
    if (!columns.Get(i).IsObject()) {
      continue;
    }
    in.columns.emplace_back();
    ReadFilterColumn(columns.Get(i).As<Napi::Object>(), in, in.columns.back());
  }
  for (uint32_t i = 0; i < conditions.Length(); ++i) {
    if (!conditions.Get(i).IsObject()) {
      continue;
    }
    in.conditions.emplace_back();
    ReadSortCondition(conditions.Get(i).As<Napi::Object>(), in, in.conditions.back());
  }
  f.columns = in.columns.empty() ? nullptr : in.columns.data();
  f.column_count = static_cast<uint32_t>(in.columns.size());
  f.has_sort = has_sort ? 1 : 0;
  f.sort_ref = PullRange(sort, "ref");
  f.column_sort = SpecPullBool(sort, "columnSort", false) ? 1 : 0;
  f.case_sensitive = SpecPullBool(sort, "caseSensitive", false) ? 1 : 0;
  f.sort_method = SpecPullInt32(sort, "sortMethod", 0);
  f.conditions = in.conditions.empty() ? nullptr : in.conditions.data();
  f.condition_count = static_cast<uint32_t>(in.conditions.size());
  return true;
}

// Shared by the sheet and table entry points: `table` selects which C
// function family a call goes to, `idx` being the sheet or table index.
Napi::Value GetAutoFilterImpl(Napi::Env env, const fm_workbook_t* wb, bool table, size_t idx) {
  fm_auto_filter af{};
  int32_t present = 0;
  const fm_status_t rc =
      table ? fm_table_get_auto_filter(wb, idx, &af, &present) : fm_sheet_get_auto_filter(wb, idx, &af, &present);
  const bool has = rc == 0 && present != 0;
  return MakeFieldResult(env, MakeStatus(env, rc), "autoFilter",
                         has ? static_cast<Napi::Value>(AutoFilterToJs(env, af)) : env.Null());
}

Napi::Value EvaluateAutoFilterImpl(Napi::Env env, const fm_workbook_t* wb, bool table, size_t idx) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("firstRow", Num(env, 0));
  out.Set("match", Napi::Array::New(env));
  auto call = [&](uint8_t* buf, size_t cap, size_t* len, uint32_t* first) {
    return table ? fm_table_evaluate_auto_filter(wb, idx, buf, cap, len, first)
                 : fm_sheet_evaluate_auto_filter(wb, idx, buf, cap, len, first);
  };
  size_t len = 0;
  uint32_t first = 0;
  fm_status_t rc = call(nullptr, 0, &len, &first);
  std::vector<uint8_t> flags(len);
  if (rc == 0 && len > 0) {
    rc = call(flags.data(), flags.size(), &len, &first);
  }
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  Napi::Array match = Napi::Array::New(env, flags.size());
  for (size_t i = 0; i < flags.size(); ++i) {
    match.Set(static_cast<uint32_t>(i), Napi::Boolean::New(env, flags[i] != 0));
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("firstRow", Num(env, first));
  out.Set("match", match);
  return out;
}

}  // namespace

Napi::Value Workbook::GetAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "autoFilter", env.Null());
  }
  return GetAutoFilterImpl(env, handle_, false, ArgU32(info, 0));
}

Napi::Value Workbook::GetTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "autoFilter", env.Null());
  }
  return GetAutoFilterImpl(env, handle_, true, ArgU32(info, 0));
}

Napi::Value Workbook::SetAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  AutoFilterInput in;
  if (!ReadAutoFilter(info, 1, in)) {
    return MakeBindingArgumentError(env, "setAutoFilter expects (sheet:number, autoFilter:object)");
  }
  return MakeStatus(env, fm_sheet_set_auto_filter(handle_, ArgU32(info, 0), &in.filter));
}

Napi::Value Workbook::SetTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  AutoFilterInput in;
  if (!ReadAutoFilter(info, 1, in)) {
    return MakeBindingArgumentError(env, "setTableAutoFilter expects (tableIdx:number, autoFilter:object)");
  }
  return MakeStatus(env, fm_table_set_auto_filter(handle_, ArgU32(info, 0), &in.filter));
}

Napi::Value Workbook::RemoveAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  return handle_ == nullptr ? NullHandleError(env)
                            : MakeStatus(env, fm_sheet_remove_auto_filter(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::RemoveTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  return handle_ == nullptr ? NullHandleError(env)
                            : MakeStatus(env, fm_table_remove_auto_filter(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::ApplyAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  return handle_ == nullptr ? NullHandleError(env)
                            : MakeStatus(env, fm_sheet_apply_auto_filter(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::ApplyTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  return handle_ == nullptr ? NullHandleError(env)
                            : MakeStatus(env, fm_table_apply_auto_filter(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::ClearAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  return handle_ == nullptr ? NullHandleError(env)
                            : MakeStatus(env, fm_sheet_clear_auto_filter(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::ClearTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  return handle_ == nullptr ? NullHandleError(env)
                            : MakeStatus(env, fm_table_clear_auto_filter(handle_, ArgU32(info, 0)));
}

Napi::Value Workbook::EvaluateAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    Napi::Object out = Napi::Object::New(env);
    out.Set("status", NullHandleError(env));
    out.Set("firstRow", Num(env, 0));
    out.Set("match", Napi::Array::New(env));
    return out;
  }
  return EvaluateAutoFilterImpl(env, handle_, false, ArgU32(info, 0));
}

Napi::Value Workbook::EvaluateTableAutoFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    Napi::Object out = Napi::Object::New(env);
    out.Set("status", NullHandleError(env));
    out.Set("firstRow", Num(env, 0));
    out.Set("match", Napi::Array::New(env));
    return out;
  }
  return EvaluateAutoFilterImpl(env, handle_, true, ArgU32(info, 0));
}

// ---- Threaded comments and persons ----------------------------------

namespace {

// Reads the `Mention[]` at `arr`; `ids` backs the borrowed strings.
void ReadMentions(const Napi::Value& arr, std::vector<std::string>& ids, std::vector<fm_mention>& out) {
  if (!arr.IsArray()) {
    return;
  }
  const Napi::Array list = arr.As<Napi::Array>();
  ids.reserve(static_cast<size_t>(list.Length()) * 2U);
  out.reserve(list.Length());
  for (uint32_t i = 0; i < list.Length(); ++i) {
    if (!list.Get(i).IsObject()) {
      continue;
    }
    const Napi::Object m = list.Get(i).As<Napi::Object>();
    ids.push_back(PullString(m, "personId"));
    ids.push_back(PullString(m, "mentionId"));
    fm_mention c{};
    c.person_id = ids[ids.size() - 2].c_str();
    c.mention_id = ids.back().c_str();
    c.start = SpecPullU32(m, "start", 0U);
    c.length = SpecPullU32(m, "length", 0U);
    out.push_back(c);
  }
}

}  // namespace

Napi::Value Workbook::GetThreadedComments(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "comments", arr);
  }
  const uint32_t sheet = ArgU32(info, 0);
  size_t count = 0;
  fm_status_t rc = fm_sheet_threaded_comment_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_threaded_comment c{};
    rc = fm_sheet_threaded_comment_at(handle_, sheet, i, &c);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", CStr(env, c.id));
    item.Set("row", Num(env, c.row));
    item.Set("col", Num(env, c.col));
    item.Set("personId", CStr(env, c.person_id));
    item.Set("created", CStr(env, c.created));
    item.Set("text", CStr(env, c.text));
    item.Set("parentId", CStr(env, c.parent_id));
    item.Set("done", Bool(env, c.done));
    Napi::Array mentions = Napi::Array::New(env, c.mention_count);
    for (uint32_t m = 0; m < c.mention_count; ++m) {
      Napi::Object mention = Napi::Object::New(env);
      mention.Set("personId", CStr(env, c.mentions[m].person_id));
      mention.Set("mentionId", CStr(env, c.mentions[m].mention_id));
      mention.Set("start", Num(env, c.mentions[m].start));
      mention.Set("length", Num(env, c.mentions[m].length));
      mentions.Set(m, mention);
    }
    item.Set("mentions", mentions);
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return MakeFieldResult(env, MakeStatus(env, rc), "comments", arr);
}

Napi::Value Workbook::AddThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    return MakeBindingArgumentError(env, "addThreadedComment expects (sheet:number, comment:object)");
  }
  const Napi::Object spec = info[1].As<Napi::Object>();
  const std::string id = PullString(spec, "id");
  const std::string person_id = PullString(spec, "personId");
  const std::string created = PullString(spec, "created");
  const std::string text = PullString(spec, "text");
  const std::string parent_id = PullString(spec, "parentId");
  std::vector<std::string> mention_ids;
  std::vector<fm_mention> mentions;
  ReadMentions(spec.Get("mentions"), mention_ids, mentions);
  fm_threaded_comment c{};
  c.id = id.c_str();
  c.row = SpecPullU32(spec, "row", 0U);
  c.col = SpecPullU32(spec, "col", 0U);
  c.person_id = person_id.c_str();
  c.created = created.c_str();
  c.text = text.c_str();
  c.parent_id = parent_id.c_str();
  c.done = SpecPullBool(spec, "done", false) ? 1 : 0;
  c.mentions = mentions.empty() ? nullptr : mentions.data();
  c.mention_count = static_cast<uint32_t>(mentions.size());
  return MakeStatus(env, fm_sheet_add_threaded_comment(handle_, ArgU32(info, 0), &c));
}

Napi::Value Workbook::EditThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string id = ArgString(info, 1);
  const std::string text = ArgString(info, 2);
  std::vector<std::string> mention_ids;
  std::vector<fm_mention> mentions;
  ReadMentions(info.Length() > 3 ? info[3] : env.Undefined(), mention_ids, mentions);
  return MakeStatus(env, fm_sheet_edit_threaded_comment(handle_, ArgU32(info, 0), id.c_str(), text.c_str(),
                                                        mentions.empty() ? nullptr : mentions.data(),
                                                        static_cast<uint32_t>(mentions.size())));
}

Napi::Value Workbook::SetThreadResolved(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string thread_id = ArgString(info, 1);
  return MakeStatus(
      env, fm_sheet_set_thread_resolved(handle_, ArgU32(info, 0), thread_id.c_str(), ArgBool(info, 2) ? 1 : 0));
}

Napi::Value Workbook::RemoveThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string id = ArgString(info, 1);
  return MakeStatus(env, fm_sheet_remove_threaded_comment(handle_, ArgU32(info, 0), id.c_str()));
}

Napi::Value Workbook::GetPersons(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "persons", arr);
  }
  size_t count = 0;
  fm_status_t rc = fm_workbook_person_count(handle_, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_person p{};
    rc = fm_workbook_person_at(handle_, i, &p);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", CStr(env, p.id));
    item.Set("displayName", CStr(env, p.display_name));
    item.Set("userId", CStr(env, p.user_id));
    item.Set("providerId", CStr(env, p.provider_id));
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return MakeFieldResult(env, MakeStatus(env, rc), "persons", arr);
}

Napi::Value Workbook::AddPerson(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 1 || !info[0].IsObject()) {
    return MakeBindingArgumentError(env, "addPerson expects (person:object)");
  }
  const Napi::Object spec = info[0].As<Napi::Object>();
  const std::string id = PullString(spec, "id");
  const std::string display_name = PullString(spec, "displayName");
  const std::string user_id = PullString(spec, "userId");
  const std::string provider_id = PullString(spec, "providerId");
  fm_person p{};
  p.id = id.c_str();
  p.display_name = display_name.c_str();
  p.user_id = user_id.c_str();
  p.provider_id = provider_id.c_str();
  return MakeStatus(env, fm_workbook_add_person(handle_, &p));
}

Napi::Value Workbook::RemovePerson(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string id = ArgString(info, 0);
  return MakeStatus(env, fm_workbook_remove_person(handle_, id.c_str()));
}

// ---- Drawing images -------------------------------------------------

namespace {

// Returns the bytes of the Uint8Array at `info[idx]`; false when it is not one.
bool ReadImageBytes(const Napi::CallbackInfo& info, size_t idx, const uint8_t*& data, size_t& len) {
  if (info.Length() <= idx || !info[idx].IsTypedArray() ||
      info[idx].As<Napi::TypedArray>().TypedArrayType() != napi_uint8_array) {
    return false;
  }
  const Napi::Uint8Array u8 = info[idx].As<Napi::Uint8Array>();
  data = u8.Data();
  len = u8.ElementLength();
  return true;
}

}  // namespace

Napi::Value Workbook::ProbeImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", Num(env, 0));
  out.Set("pxWidth", Num(env, 0));
  out.Set("pxHeight", Num(env, 0));
  const uint8_t* data = nullptr;
  size_t len = 0;
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  if (!ReadImageBytes(info, 0, data, len)) {
    out.Set("status", MakeBindingArgumentError(env, "probeImage expects (bytes:Uint8Array)"));
    return out;
  }
  fm_image_info img{};
  const fm_status_t rc = fm_workbook_probe_image(handle_, data, len, &img);
  if (rc == 0) {
    out.Set("format", Num(env, img.format));
    out.Set("pxWidth", Num(env, img.px_width));
    out.Set("pxHeight", Num(env, img.px_height));
  }
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::ListDrawingObjects(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  size_t count = 0;
  fm_status_t rc = fm_sheet_drawing_object_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_drawing_object o{};
    rc = fm_sheet_drawing_object_at(handle_, sheet, i, &o);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("objectId", Num(env, o.object_id));
    item.Set("kind", Num(env, o.kind));
    item.Set("anchorKind", Num(env, o.anchor_kind));
    item.Set("editAs", Num(env, o.edit_as));
    item.Set("fromRow", Num(env, o.from_row));
    item.Set("fromCol", Num(env, o.from_col));
    item.Set("fromRowOff", Num(env, static_cast<double>(o.from_row_off)));
    item.Set("fromColOff", Num(env, static_cast<double>(o.from_col_off)));
    item.Set("toRow", Num(env, o.to_row));
    item.Set("toCol", Num(env, o.to_col));
    item.Set("toRowOff", Num(env, static_cast<double>(o.to_row_off)));
    item.Set("toColOff", Num(env, static_cast<double>(o.to_col_off)));
    item.Set("cx", Num(env, static_cast<double>(o.cx)));
    item.Set("cy", Num(env, static_cast<double>(o.cy)));
    item.Set("imageFormat", Num(env, o.image_format));
    item.Set("name", CStr(env, o.name));
    item.Set("descr", CStr(env, o.descr));
    item.Set("mediaPath", CStr(env, o.media_path));
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::GetImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("format", Num(env, 0));
  out.Set("bytes", Napi::Uint8Array::New(env, 0));
  out.Set("pxWidth", Num(env, 0));
  out.Set("pxHeight", Num(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint8_t* bytes = nullptr;
  size_t len = 0;
  fm_image_info img{};
  const fm_status_t rc = fm_sheet_get_image(handle_, ArgU32(info, 0), ArgU32(info, 1), &bytes, &len, &img);
  if (rc == 0) {
    // The C pointer dies on the next mutation, so the bytes are copied out.
    Napi::Uint8Array copy = Napi::Uint8Array::New(env, len);
    if (len != 0) {
      std::memcpy(copy.Data(), bytes, len);
    }
    out.Set("format", Num(env, img.format));
    out.Set("bytes", copy);
    out.Set("pxWidth", Num(env, img.px_width));
    out.Set("pxHeight", Num(env, img.px_height));
  }
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::InsertImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", Num(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  const bool has_opts = info.Length() > 2 && info[2].IsObject();
  if (!ReadImageBytes(info, 1, data, len) || (info.Length() > 2 && !info[2].IsUndefined() && !has_opts)) {
    out.Set("status",
            MakeBindingArgumentError(env, "insertImage expects (sheet:number, bytes:Uint8Array, opts?:object)"));
    return out;
  }
  const Napi::Object spec = has_opts ? info[2].As<Napi::Object>() : Napi::Object::New(env);
  const std::string name = PullString(spec, "name");
  const std::string descr = PullString(spec, "descr");
  fm_image_insert opts{};
  opts.name = name.c_str();
  opts.descr = descr.c_str();
  opts.anchor_kind = SpecPullInt32(spec, "anchorKind", FM_ANCHOR_KIND_ONE_CELL);
  opts.edit_as = SpecPullInt32(spec, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL);
  opts.row = SpecPullU32(spec, "row", 0U);
  opts.col = SpecPullU32(spec, "col", 0U);
  opts.row_off_emu = static_cast<int64_t>(SpecPullDouble(spec, "rowOffEmu", 0.0));
  opts.col_off_emu = static_cast<int64_t>(SpecPullDouble(spec, "colOffEmu", 0.0));
  opts.width_emu = static_cast<int64_t>(SpecPullDouble(spec, "widthEmu", 0.0));
  opts.height_emu = static_cast<int64_t>(SpecPullDouble(spec, "heightEmu", 0.0));
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_insert_image(handle_, ArgU32(info, 0), data, len, &opts, &object_id);
  out.Set("objectId", Num(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::RemoveImage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_remove_image(handle_, ArgU32(info, 0), ArgU32(info, 1)));
}

}  // namespace formulon_node
