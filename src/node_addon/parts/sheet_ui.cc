// Per-sheet UI feature bindings: merges, comments, hyperlinks, data
// validations, and threaded comments with their persons.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Merges ---------------------------------------------------------

namespace {

// Enumerates a per-sheet list through its count / at-index pair into a
// `ListResult` array, stopping at the first failed element read.
template <typename T>
Napi::Value SheetListResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                            fm_status_t (*count_fn)(const fm_workbook_t*, std::size_t, std::size_t*),
                            fm_status_t (*at_fn)(const fm_workbook_t*, std::size_t, std::size_t, T*),
                            Napi::Object (*to_js)(Napi::Env, const T&)) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const std::size_t sheet = Workbook::ArgU32(info, 0);
  std::size_t count = 0;
  fm_status_t rc = count_fn(handle, sheet, &count);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  for (std::size_t i = 0; i < count; ++i) {
    T entry{};
    rc = at_fn(handle, sheet, i, &entry);
    if (rc != 0) {
      return FinishListResult(env, arr, rc);
    }
    arr.Set(i, to_js(env, entry));
  }
  return FinishListResult(env, arr, 0);
}

// Reads element `(sheet, index)` of a per-sheet index-keyed collection into
// its `<Record> & {status}` result.
template <typename T>
Napi::Value SheetItemResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                            fm_status_t (*at_fn)(const fm_workbook_t*, std::size_t, std::size_t, T*),
                            Napi::Object (*to_js)(Napi::Env, const T&)) {
  Napi::Env env = info.Env();
  T entry{};
  const fm_status_t rc = handle != nullptr ? at_fn(handle, Workbook::ArgU32(info, 0), Workbook::ArgU32(info, 1), &entry)
                                           : kBindingInvalidHandle;
  return FinishItemResult(env, to_js(env, rc == 0 ? entry : T{}), rc);
}

// Reads the element count of a per-sheet collection into a `NumberResult`.
Napi::Value SheetCountResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                             fm_status_t (*count_fn)(const fm_workbook_t*, std::size_t, std::size_t*)) {
  std::size_t count = 0;
  const fm_status_t rc =
      handle != nullptr ? count_fn(handle, Workbook::ArgU32(info, 0), &count) : kBindingInvalidHandle;
  return MakeNumberResult(info.Env(), rc, static_cast<double>(count));
}

// Shared body of the per-sheet merge add / remove entries taking `(sheet, range)`.
using MergeEditFn = fm_status_t (*)(fm_workbook_t*, std::size_t, const fm_cell_rect_t*);

// `handle` aliases the wrapper's member, so the C call sees the handle as it
// is after the range getters ran, not a copy taken before them.
Napi::Value InvokeMergeEdit(const Napi::CallbackInfo& info, fm_workbook_t* const& handle, MergeEditFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t range{};
  if (!CellRectArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  fm_status_t rc = fn(handle, sheet, &range);
  return MakeStatus(env, rc);
}

}  // namespace

Napi::Value Workbook::AddMerge(const Napi::CallbackInfo& info) {
  return InvokeMergeEdit(info, handle_, &fm_sheet_add_merge);
}

Napi::Value Workbook::RemoveMerge(const Napi::CallbackInfo& info) {
  return InvokeMergeEdit(info, handle_, &fm_sheet_remove_merge);
}

Napi::Value Workbook::RemoveMergeAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_merge_at);
}

Napi::Value Workbook::ClearMerges(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_merges);
}

Napi::Value Workbook::GetMerges(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_merge_count, &fm_sheet_get_merge_at, &RangeToJs);
}

Napi::Value Workbook::GetMergesInRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t range{};
  if (!CellRectArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  uint32_t total = 0;
  fm_status_t rc = fm_sheet_get_merges_in_range(handle_, sheet, &range, nullptr, 0, &total);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  std::vector<fm_cell_rect_t> merges(total);
  uint32_t again = 0;
  rc = fm_sheet_get_merges_in_range(handle_, sheet, &range, merges.data(), total, &again);
  if (rc != 0) {
    return FinishListResult(env, arr, rc);
  }
  for (uint32_t i = 0; i < total; ++i) {
    arr.Set(i, RangeToJs(env, merges[i]));
  }
  return FinishListResult(env, arr, 0);
}

// ---- Comments -------------------------------------------------------

Napi::Value Workbook::GetComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  int32_t present = 0;
  fm_comment_t c{};
  const fm_status_t rc = fm_sheet_get_comment(handle_, sheet, row, col, &present, &c);
  if (rc != 0 || present == 0) {
    return MakeOptionalResult(env, rc, present, env.Null());
  }
  Napi::Object comment = Napi::Object::New(env);
  comment.Set("author", JsString(env, c.author));
  comment.Set("text", JsString(env, c.text));
  return MakeOptionalResult(env, rc, present, comment);
}

namespace {

Napi::Object CommentToJs(Napi::Env env, const fm_comment_t& c) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("row", Napi::Number::New(env, c.row));
  item.Set("col", Napi::Number::New(env, c.col));
  item.Set("author", JsString(env, c.author));
  item.Set("text", JsString(env, c.text));
  return item;
}

}  // namespace

Napi::Value Workbook::GetComments(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_comment_count, &fm_sheet_get_comment_at, &CommentToJs);
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

Napi::Value Workbook::GetCommentVisible(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  int32_t visible = 0;
  const fm_status_t rc =
      fm_sheet_get_comment_visible(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &visible);
  const bool absent = rc == FM_STATUS_NOT_FOUND;
  return MakeOptionalResult(env, absent ? 0 : rc, absent ? 0 : 1,
                            rc == 0 ? Napi::Value(Napi::Boolean::New(env, visible != 0)) : env.Null());
}

Napi::Value Workbook::SetCommentVisible(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_sheet_set_comment_visible(handle, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), ArgFlag(info, 3));
  });
}

// ---- Protection records ---------------------------------------------

Napi::Value Workbook::GetWorkbookProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "protection", env.Null());
  }
  fm_workbook_protection_t p{};
  const fm_status_t rc = fm_workbook_get_workbook_protection(handle_, &p);
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "protection", env.Null());
  }
  Napi::Object o = Napi::Object::New(env);
  o.Set("lockStructure", Napi::Number::New(env, p.lock_structure));
  o.Set("lockWindows", Napi::Number::New(env, p.lock_windows));
  o.Set("lockRevision", Napi::Number::New(env, p.lock_revision));
  o.Set("workbookSpinCount", Napi::Number::New(env, p.workbook_spin_count));
  o.Set("revisionsSpinCount", Napi::Number::New(env, p.revisions_spin_count));
  o.Set("workbookPassword", JsString(env, p.workbook_password));
  o.Set("workbookAlgorithmName", JsString(env, p.workbook_algorithm_name));
  o.Set("workbookHashValue", JsString(env, p.workbook_hash_value));
  o.Set("workbookSaltValue", JsString(env, p.workbook_salt_value));
  o.Set("revisionsPassword", JsString(env, p.revisions_password));
  o.Set("revisionsAlgorithmName", JsString(env, p.revisions_algorithm_name));
  o.Set("revisionsHashValue", JsString(env, p.revisions_hash_value));
  o.Set("revisionsSaltValue", JsString(env, p.revisions_salt_value));
  return MakeFieldResult(env, MakeOkStatus(env), "protection", o);
}

Napi::Value Workbook::SetWorkbookProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object in = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  std::string text[8];
  (void)reader.String(in, "workbookPassword", &text[0]);
  (void)reader.String(in, "workbookAlgorithmName", &text[1]);
  (void)reader.String(in, "workbookHashValue", &text[2]);
  (void)reader.String(in, "workbookSaltValue", &text[3]);
  (void)reader.String(in, "revisionsPassword", &text[4]);
  (void)reader.String(in, "revisionsAlgorithmName", &text[5]);
  (void)reader.String(in, "revisionsHashValue", &text[6]);
  (void)reader.String(in, "revisionsSaltValue", &text[7]);
  fm_workbook_protection_t p{};
  p.lock_structure = reader.I32(in, "lockStructure", 0);
  p.lock_windows = reader.I32(in, "lockWindows", 0);
  p.lock_revision = reader.I32(in, "lockRevision", 0);
  p.workbook_spin_count = reader.U32(in, "workbookSpinCount", 0U);
  p.revisions_spin_count = reader.U32(in, "revisionsSpinCount", 0U);
  if (!reader.ok()) {
    return env.Undefined();
  }
  p.workbook_password = text[0].c_str();
  p.workbook_algorithm_name = text[1].c_str();
  p.workbook_hash_value = text[2].c_str();
  p.workbook_salt_value = text[3].c_str();
  p.revisions_password = text[4].c_str();
  p.revisions_algorithm_name = text[5].c_str();
  p.revisions_hash_value = text[6].c_str();
  p.revisions_salt_value = text[7].c_str();
  return MakeStatus(env, fm_workbook_set_workbook_protection(handle_, &p));
}

namespace {

Napi::Object ProtectedRangeToJs(Napi::Env env, const fm_protected_range_t& r) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("spinCount", Napi::Number::New(env, r.spin_count));
  item.Set("name", JsString(env, r.name));
  Napi::Array ranges = Napi::Array::New(env, r.range_count);
  for (size_t i = 0; i < r.range_count; ++i) {
    ranges.Set(static_cast<uint32_t>(i), RangeToJs(env, r.ranges[i]));
  }
  item.Set("ranges", ranges);
  item.Set("password", JsString(env, r.password));
  item.Set("algorithmName", JsString(env, r.algorithm_name));
  item.Set("hashValue", JsString(env, r.hash_value));
  item.Set("saltValue", JsString(env, r.salt_value));
  item.Set("securityDescriptor", JsString(env, r.security_descriptor));
  return item;
}

// Reads the range spec at `info[idx]` and returns what `edit` makes of it.
// `handle` aliases the wrapper's member, as in `InvokeMergeEdit`.
template <typename Edit>
Napi::Value EditProtectedRange(fm_workbook_t* const& handle, const Napi::CallbackInfo& info, size_t idx, Edit edit) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object in = Workbook::ArgObject(info, idx);
  CheckedSpecReader reader(env);
  std::string text[6];
  (void)reader.String(in, "name", &text[0]);
  std::vector<fm_cell_rect_t> ranges;
  Napi::Array ranges_array;
  if (reader.RequiredArray(in, "ranges", &ranges_array)) {
    (void)ReadCellRects(reader, ranges_array, "ranges", &ranges);
  }
  (void)reader.String(in, "password", &text[1]);
  (void)reader.String(in, "algorithmName", &text[2]);
  (void)reader.String(in, "hashValue", &text[3]);
  (void)reader.String(in, "saltValue", &text[4]);
  (void)reader.String(in, "securityDescriptor", &text[5]);
  fm_protected_range_t r{};
  r.spin_count = reader.U32(in, "spinCount", 0U);
  if (!reader.ok()) {
    return env.Undefined();
  }
  r.name = text[0].c_str();
  r.ranges = ranges.data();
  r.range_count = ranges.size();
  r.password = text[1].c_str();
  r.algorithm_name = text[2].c_str();
  r.hash_value = text[3].c_str();
  r.salt_value = text[4].c_str();
  r.security_descriptor = text[5].c_str();
  return edit(handle, r);
}

}  // namespace

Napi::Value Workbook::GetProtectedRangeCount(const Napi::CallbackInfo& info) {
  return SheetCountResult(info, handle_, &fm_sheet_get_protected_range_count);
}

Napi::Value Workbook::GetProtectedRange(const Napi::CallbackInfo& info) {
  return SheetItemResult(info, handle_, &fm_sheet_get_protected_range, &ProtectedRangeToJs);
}

Napi::Value Workbook::AddProtectedRange(const Napi::CallbackInfo& info) {
  if (handle_ == nullptr) {
    return MakeIndexResult(info.Env(), NullHandleError(info.Env()), 0);
  }
  return EditProtectedRange(handle_, info, 1, [&](fm_workbook_t* handle, const fm_protected_range_t& r) {
    std::size_t out = 0;
    const fm_status_t rc = fm_sheet_add_protected_range(handle, ArgU32(info, 0), &r, &out);
    return MakeIndexResult(info.Env(), rc, static_cast<uint32_t>(out));
  });
}

Napi::Value Workbook::UpdateProtectedRange(const Napi::CallbackInfo& info) {
  return EditProtectedRange(handle_, info, 2, [&](fm_workbook_t* handle, const fm_protected_range_t& r) {
    return MakeStatus(info.Env(), fm_sheet_update_protected_range(handle, ArgU32(info, 0), ArgU32(info, 1), &r));
  });
}

Napi::Value Workbook::RemoveProtectedRange(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_protected_range);
}

// ---- Hyperlinks -----------------------------------------------------

namespace {

// Adds a hyperlink over `rect` whose texts are the already type-checked
// arguments `(target, display, tooltip, location)` from `text_index` on.
// Empty strings are forwarded as NULL.
fm_status_t AddHyperlinkFromArgs(fm_workbook_t* wb, const Napi::CallbackInfo& info, const fm_cell_rect_t& rect,
                                 std::size_t text_index) {
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  // Keep the std::string buffers alive until after the C ABI call so the
  // borrowed `const char*` pointers in `fm_hyperlink_t` stay valid.
  const std::string target = Workbook::ArgString(info, text_index);
  const std::string display = Workbook::ArgString(info, text_index + 1);
  const std::string tooltip = Workbook::ArgString(info, text_index + 2);
  const std::string location = Workbook::ArgString(info, text_index + 3);
  fm_hyperlink_t hl{};
  hl.rect = rect;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  return fm_sheet_add_hyperlink(wb, sheet, &hl);
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
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  return MakeStatus(env, AddHyperlinkFromArgs(handle_, info, fm_cell_rect_t{row, col, row, col}, 3));
}

Napi::Value Workbook::AddHyperlinkRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 6 || !info[0].IsNumber() || !info[1].IsObject() || !info[2].IsString() || !info[3].IsString() ||
      !info[4].IsString() || !info[5].IsString()) {
    Napi::TypeError::New(env,
                         "addHyperlinkRange expects (sheet:number, range:CellRect, target:string, "
                         "display:string, tooltip:string, location:string)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  return MakeStatus(env, AddHyperlinkFromArgs(handle_, info, rect, 2));
}

namespace {

Napi::Object HyperlinkToJs(Napi::Env env, const fm_hyperlink_t& h) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("rect", RangeToJs(env, h.rect));
  item.Set("target", JsString(env, h.target));
  item.Set("location", JsString(env, h.location));
  item.Set("display", JsString(env, h.display));
  item.Set("tooltip", JsString(env, h.tooltip));
  return item;
}

}  // namespace

Napi::Value Workbook::GetHyperlinks(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_hyperlink_count, &fm_sheet_get_hyperlink_at, &HyperlinkToJs);
}

Napi::Value Workbook::GetHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  fm_hyperlink_t h{};
  int32_t present = 0;
  const fm_status_t rc =
      fm_sheet_get_hyperlink(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &h);
  return MakeOptionalResult(env, rc, present,
                            rc == 0 && present != 0 ? Napi::Value(HyperlinkToJs(env, h)) : env.Null());
}

Napi::Value Workbook::RemoveHyperlink(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_hyperlink);
}

Napi::Value Workbook::RemoveHyperlinkAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_hyperlink_at);
}

Napi::Value Workbook::ClearHyperlinks(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_hyperlinks);
}

// ---- Validations ----------------------------------------------------

namespace {

Napi::Object ValidationToJs(Napi::Env env, const fm_data_validation_t& v) {
  Napi::Object item = Napi::Object::New(env);
  Napi::Array ranges = Napi::Array::New(env);
  for (size_t r = 0; r < v.range_count; ++r) {
    ranges.Set(static_cast<uint32_t>(r), RangeToJs(env, v.ranges[r]));
  }
  item.Set("ranges", ranges);
  item.Set("type", Napi::Number::New(env, v.type));
  item.Set("op", Napi::Number::New(env, v.op));
  item.Set("errorStyle", Napi::Number::New(env, v.error_style));
  item.Set("allowBlank", Napi::Boolean::New(env, v.allow_blank != 0));
  item.Set("showInputMessage", Napi::Boolean::New(env, v.show_input_message != 0));
  item.Set("showErrorMessage", Napi::Boolean::New(env, v.show_error_message != 0));
  item.Set("showDropDown", Napi::Boolean::New(env, v.show_dropdown != 0));
  item.Set("formula1", JsString(env, v.formula1));
  item.Set("formula2", JsString(env, v.formula2));
  item.Set("errorTitle", JsString(env, v.error_title));
  item.Set("errorMessage", JsString(env, v.error_message));
  item.Set("promptTitle", JsString(env, v.prompt_title));
  item.Set("promptMessage", JsString(env, v.prompt_message));
  return item;
}

}  // namespace

Napi::Value Workbook::GetValidationCount(const Napi::CallbackInfo& info) {
  return SheetCountResult(info, handle_, &fm_sheet_get_validation_count);
}

Napi::Value Workbook::GetValidation(const Napi::CallbackInfo& info) {
  return SheetItemResult(info, handle_, &fm_sheet_get_validation, &ValidationToJs);
}

Napi::Value Workbook::AddValidation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t sheet = ArgU32(info, 0);
  Napi::Object v = ArgObject(info, 1);
  CheckedSpecReader reader(env);

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  std::vector<fm_cell_rect_t> ranges_buf;
  Napi::Array ranges_arr;
  if (reader.RequiredArray(v, "ranges", &ranges_arr)) {
    const uint32_t n = ranges_arr.Length();
    ranges_buf.reserve(n);
    for (uint32_t i = 0; i < n && reader.ok(); ++i) {
      Napi::Value range_value;
      if (!reader.ArrayElement(ranges_arr, i, &range_value)) {
        continue;
      }
      Napi::Object range;
      if (!reader.Object(range_value, "ranges[]", &range)) {
        break;
      }
      fm_cell_rect_t parsed{};
      if (!ReadCellRect(reader, range, &parsed)) {
        break;
      }
      ranges_buf.push_back(parsed);
    }
  }
  std::string formula1;
  std::string formula2;
  std::string error_title;
  std::string error_message;
  std::string prompt_title;
  std::string prompt_message;
  reader.String(v, "formula1", &formula1);
  reader.String(v, "formula2", &formula2);
  reader.String(v, "errorTitle", &error_title);
  reader.String(v, "errorMessage", &error_message);
  reader.String(v, "promptTitle", &prompt_title);
  reader.String(v, "promptMessage", &prompt_message);

  fm_data_validation_t dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = ranges_buf.size();
  dv.type = reader.U8(v, "type", 0U);
  dv.op = reader.U8(v, "op", 0U);
  dv.error_style = reader.U8(v, "errorStyle", 0U);
  dv.allow_blank = reader.Bool(v, "allowBlank", false) ? 1 : 0;
  dv.show_input_message = reader.Bool(v, "showInputMessage", false) ? 1 : 0;
  dv.show_error_message = reader.Bool(v, "showErrorMessage", false) ? 1 : 0;
  dv.show_dropdown = reader.Bool(v, "showDropDown", true) ? 1 : 0;
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  if (!reader.ok()) {
    return env.Undefined();
  }
  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, &dv, &out);
  return MakeIndexResult(env, rc, static_cast<uint32_t>(out));
}

Napi::Value Workbook::RemoveValidationAt(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_validation);
}

Napi::Value Workbook::ClearValidations(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_validations);
}

// ---- Threaded comments and persons ----------------------------------

namespace {

// Reads the `Mention[]` at `arr`; `ids` backs the borrowed strings.
bool ReadMentions(CheckedSpecReader& reader, const Napi::Value& arr, std::vector<std::string>& ids,
                  std::vector<fm_mention_t>& out) {
  Napi::Array list;
  if (!reader.Array(arr, "mentions", &list)) {
    return reader.ok();
  }
  ids.reserve(static_cast<size_t>(list.Length()) * 2U);
  out.reserve(list.Length());
  for (uint32_t i = 0; i < list.Length() && reader.ok(); ++i) {
    Napi::Value mention_value;
    if (!reader.ArrayElement(list, i, &mention_value)) {
      continue;
    }
    Napi::Object m;
    if (!reader.Object(mention_value, "mentions[]", &m)) {
      break;
    }
    std::string person_id;
    std::string mention_id;
    reader.String(m, "personId", &person_id);
    reader.String(m, "mentionId", &mention_id);
    ids.push_back(std::move(person_id));
    ids.push_back(std::move(mention_id));
    fm_mention_t c{};
    c.person_id = ids[ids.size() - 2].c_str();
    c.mention_id = ids.back().c_str();
    c.start = reader.U32(m, "start", 0U);
    c.length = reader.U32(m, "length", 0U);
    if (!reader.ok()) {
      break;
    }
    out.push_back(c);
  }
  return reader.ok();
}

// The JS `ThreadedComment` record of `c`.
Napi::Object ThreadedCommentToJs(Napi::Env env, const fm_threaded_comment_t& c) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("id", JsString(env, c.id));
  item.Set("row", JsNumber(env, c.row));
  item.Set("col", JsNumber(env, c.col));
  item.Set("personId", JsString(env, c.person_id));
  item.Set("created", JsString(env, c.created));
  item.Set("text", JsString(env, c.text));
  item.Set("parentId", JsString(env, c.parent_id));
  item.Set("done", JsBool(env, c.done));
  Napi::Array mentions = Napi::Array::New(env, c.mention_count);
  for (uint32_t m = 0; m < c.mention_count; ++m) {
    Napi::Object mention = Napi::Object::New(env);
    mention.Set("personId", JsString(env, c.mentions[m].person_id));
    mention.Set("mentionId", JsString(env, c.mentions[m].mention_id));
    mention.Set("start", JsNumber(env, c.mentions[m].start));
    mention.Set("length", JsNumber(env, c.mentions[m].length));
    mentions.Set(m, mention);
  }
  item.Set("mentions", mentions);
  return item;
}

// The JS `Person` record of `p`.
Napi::Object PersonToJs(Napi::Env env, const fm_person_t& p) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("id", JsString(env, p.id));
  item.Set("displayName", JsString(env, p.display_name));
  item.Set("userId", JsString(env, p.user_id));
  item.Set("providerId", JsString(env, p.provider_id));
  return item;
}

}  // namespace

Napi::Value Workbook::GetThreadedComments(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "comments", arr);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_threaded_comment_t>(
      arr, [&](size_t* n) { return fm_sheet_get_threaded_comment_count(handle_, sheet, n); },
      [&](size_t i, fm_threaded_comment_t* c) { return fm_sheet_get_threaded_comment_at(handle_, sheet, i, c); },
      [&](const fm_threaded_comment_t& c) { return ThreadedCommentToJs(env, c); });
  return MakeFieldResult(env, MakeStatus(env, rc), "comments", arr);
}

Napi::Value Workbook::AddThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 1);
  std::string id;
  std::string person_id;
  std::string created;
  std::string text;
  std::string parent_id;
  reader.String(spec, "id", &id);
  reader.String(spec, "personId", &person_id);
  reader.String(spec, "created", &created);
  reader.String(spec, "text", &text);
  reader.String(spec, "parentId", &parent_id);
  std::vector<std::string> mention_ids;
  std::vector<fm_mention_t> mentions;
  Napi::Array mention_array;
  if (reader.Array(spec, "mentions", &mention_array)) {
    ReadMentions(reader, mention_array, mention_ids, mentions);
  }
  fm_threaded_comment_t c{};
  c.id = id.c_str();
  c.row = reader.U32(spec, "row", 0U);
  c.col = reader.U32(spec, "col", 0U);
  c.person_id = person_id.c_str();
  c.created = created.c_str();
  c.text = text.c_str();
  c.parent_id = parent_id.c_str();
  c.done = reader.Bool(spec, "done", false) ? 1 : 0;
  c.mentions = mentions.empty() ? nullptr : mentions.data();
  c.mention_count = static_cast<uint32_t>(mentions.size());
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_add_threaded_comment(handle_, ArgU32(info, 0), &c));
}

Napi::Value Workbook::EditThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  const std::string id = ArgString(info, 1);
  const std::string text = ArgString(info, 2);
  std::vector<std::string> mention_ids;
  std::vector<fm_mention_t> mentions;
  if (info.Length() > 3) {
    ReadMentions(reader, info[3], mention_ids, mentions);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_update_threaded_comment(handle_, ArgU32(info, 0), id.c_str(), text.c_str(),
                                                          mentions.empty() ? nullptr : mentions.data(),
                                                          static_cast<uint32_t>(mentions.size())));
}

Napi::Value Workbook::SetThreadResolved(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string thread_id = ArgString(info, 1);
    return fm_sheet_set_thread_resolved(handle, ArgU32(info, 0), thread_id.c_str(), ArgFlag(info, 2));
  });
}

Napi::Value Workbook::GetThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  const std::string id = ArgString(info, 1);
  fm_threaded_comment_t c{};
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_threaded_comment(handle_, ArgU32(info, 0), id.c_str(), &present, &c);
  return MakeOptionalResult(env, rc, present,
                            rc == 0 && present != 0 ? Napi::Value(ThreadedCommentToJs(env, c)) : env.Null());
}

Napi::Value Workbook::RemoveThreadedComment(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string id = ArgString(info, 1);
    return fm_sheet_remove_threaded_comment(handle, ArgU32(info, 0), id.c_str());
  });
}

Napi::Value Workbook::GetPersons(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "persons", arr);
  }
  const fm_status_t rc = CollectList<fm_person_t>(
      arr, [&](size_t* n) { return fm_workbook_get_person_count(handle_, n); },
      [&](size_t i, fm_person_t* p) { return fm_workbook_get_person_at(handle_, i, p); },
      [&](const fm_person_t& p) { return PersonToJs(env, p); });
  return MakeFieldResult(env, MakeStatus(env, rc), "persons", arr);
}

Napi::Value Workbook::GetPerson(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  const std::string id = ArgString(info, 0);
  fm_person_t p{};
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_person(handle_, id.c_str(), &present, &p);
  return MakeOptionalResult(env, rc, present, rc == 0 && present != 0 ? Napi::Value(PersonToJs(env, p)) : env.Null());
}

Napi::Value Workbook::AddPerson(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 0);
  std::string id;
  std::string display_name;
  std::string user_id;
  std::string provider_id;
  reader.String(spec, "id", &id);
  reader.String(spec, "displayName", &display_name);
  reader.String(spec, "userId", &user_id);
  reader.String(spec, "providerId", &provider_id);
  fm_person_t p{};
  p.id = id.c_str();
  p.display_name = display_name.c_str();
  p.user_id = user_id.c_str();
  p.provider_id = provider_id.c_str();
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_workbook_add_person(handle_, &p));
}

Napi::Value Workbook::RemovePerson(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string id = ArgString(info, 0);
    return fm_workbook_remove_person(handle, id.c_str());
  });
}

}  // namespace formulon_node
