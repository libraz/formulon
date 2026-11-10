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
// array, stopping at the first failed element read.
template <typename T>
Napi::Value SheetListResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                            fm_status_t (*count_fn)(const fm_workbook_t*, std::size_t, std::size_t*),
                            fm_status_t (*at_fn)(const fm_workbook_t*, std::size_t, std::size_t, T*),
                            Napi::Object (*to_js)(Napi::Env, const T&)) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = Workbook::ArgU32(info, 0);
  std::size_t count = 0;
  fm_status_t rc = count_fn(handle, sheet, &count);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  for (std::size_t i = 0; i < count; ++i) {
    T entry = InitOut<T>();
    rc = at_fn(handle, sheet, i, &entry);
    if (rc != 0) {
      return ThrowStatus(env, rc);
    }
    arr.Set(i, to_js(env, entry));
  }
  return arr;
}

// Reads the element count of a per-sheet collection.
Napi::Value SheetCountResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle,
                             fm_status_t (*count_fn)(const fm_workbook_t*, std::size_t, std::size_t*)) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t count = 0;
  const fm_status_t rc = count_fn(handle, Workbook::ArgU32(info, 0), &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

// Reads the `(sheet, range)` arguments of a merge edit and returns what `edit` makes of them. `handle` aliases the
// wrapper's member, so the C call sees the handle as it is after the range getters ran, not a copy taken before them.
template <typename Edit>
Napi::Value InvokeMergeEdit(const Napi::CallbackInfo& info, fm_workbook_t* const& handle, Edit edit) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t range{};
  if (!CellRectArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  return edit(handle, Workbook::ArgU32(info, 0), range);
}

}  // namespace

// Returns the merge as stored.
Napi::Value Workbook::AddMerge(const Napi::CallbackInfo& info) {
  return InvokeMergeEdit(info, handle_, [&](fm_workbook_t* handle, uint32_t sheet, const fm_cell_rect_t& range) {
    fm_cell_rect_t stored{};
    const fm_status_t rc = fm_sheet_add_merge(handle, sheet, &range, &stored);
    return ValueOrThrow(info.Env(), rc, rc == 0 ? Napi::Value(RangeToJs(info.Env(), stored)) : info.Env().Undefined());
  });
}

Napi::Value Workbook::RemoveMerge(const Napi::CallbackInfo& info) {
  return InvokeMergeEdit(info, handle_, [&](fm_workbook_t* handle, uint32_t sheet, const fm_cell_rect_t& range) {
    return ValueOrThrow(info.Env(), fm_sheet_remove_merge(handle, sheet, &range), info.Env().Undefined());
  });
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
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  fm_cell_rect_t range{};
  if (!CellRectArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
  std::vector<fm_cell_rect_t> merges;
  const fm_status_t rc = FillCallerBuffer(merges, [&](fm_cell_rect_t* out, size_t capacity, size_t* count) {
    return fm_sheet_get_merges_in_range(handle_, sheet, &range, out, capacity, count);
  });
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  for (uint32_t i = 0; i < merges.size(); ++i) {
    arr.Set(i, RangeToJs(env, merges[i]));
  }
  return arr;
}

// ---- Comments -------------------------------------------------------

Napi::Value Workbook::GetComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  int32_t present = 0;
  fm_comment_t c = FM_STRUCT_INIT(fm_comment_t);
  const fm_status_t rc = fm_sheet_get_comment(handle_, sheet, row, col, &present, &c);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object comment = Napi::Object::New(env);
  comment.Set("author", JsStringOrNull(env, c.author));
  comment.Set("text", JsString(env, c.text));
  return OptionalOrNull(env, rc, present, comment);
}

namespace {

// The JS `SheetComment` record of `c`: the note with its anchor `cell`.
Napi::Object CommentToJs(Napi::Env env, const fm_comment_t& c) {
  Napi::Object cell = Napi::Object::New(env);
  cell.Set("row", Napi::Number::New(env, c.row));
  cell.Set("col", Napi::Number::New(env, c.col));
  Napi::Object item = Napi::Object::New(env);
  item.Set("cell", cell);
  item.Set("author", JsStringOrNull(env, c.author));
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
    return ThrowInvalidHandle(env);
  }
  // `null` removes the note and does nothing when there is none; an omitted `author` is empty.
  CheckedSpecReader reader(env);
  Napi::Object in;
  const bool present = reader.Object(info[3], "", &in);
  std::string author;
  std::string text;
  fm_comment_t c = FM_STRUCT_INIT(fm_comment_t);
  if (present) {
    c.author = reader.String(in, "author", &author) ? author.c_str() : nullptr;
    text = reader.RequiredString(in, "text");
    c.text = text.c_str();
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc =
      fm_sheet_set_comment(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), present ? &c : nullptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::GetCommentVisible(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  int32_t visible = 0;
  const fm_status_t rc =
      fm_sheet_get_comment_visible(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &visible);
  return OptionalOrNull(env, rc, present, Napi::Boolean::New(env, visible != 0));
}

Napi::Value Workbook::SetCommentVisible(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const int32_t visible = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_comment_visible(handle, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), visible);
  });
}

// ---- Protection records ---------------------------------------------

Napi::Value Workbook::GetWorkbookProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_workbook_protection(handle_, &present, &p);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object o = Napi::Object::New(env);
  o.Set("lockStructure", JsBool(env, p.lock_structure));
  o.Set("lockWindows", JsBool(env, p.lock_windows));
  o.Set("lockRevision", JsBool(env, p.lock_revision));
  o.Set("workbookSpinCount", Napi::Number::New(env, p.workbook_spin_count));
  o.Set("revisionsSpinCount", Napi::Number::New(env, p.revisions_spin_count));
  o.Set("workbookLegacyPasswordHash", JsStringOrNull(env, p.workbook_legacy_password_hash));
  o.Set("workbookAlgorithmName", JsStringOrNull(env, p.workbook_algorithm_name));
  o.Set("workbookHashValue", JsStringOrNull(env, p.workbook_hash_value));
  o.Set("workbookSaltValue", JsStringOrNull(env, p.workbook_salt_value));
  o.Set("revisionsLegacyPasswordHash", JsStringOrNull(env, p.revisions_legacy_password_hash));
  o.Set("revisionsAlgorithmName", JsStringOrNull(env, p.revisions_algorithm_name));
  o.Set("revisionsHashValue", JsStringOrNull(env, p.revisions_hash_value));
  o.Set("revisionsSaltValue", JsStringOrNull(env, p.revisions_salt_value));
  return o;
}

Napi::Value Workbook::SetWorkbookProtection(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // `null` removes `<workbookProtection>`; a missing or `null` string is no stored value.
  CheckedSpecReader reader(env);
  Napi::Object in;
  if (!reader.Object(info[0], "", &in)) {
    return reader.ok() ? ValueOrThrow(env, fm_workbook_set_workbook_protection(handle_, nullptr), env.Undefined())
                       : env.Undefined();
  }
  std::string text[8];
  const auto optional = [&](const char* key, std::string* store) {
    return reader.String(in, key, store) ? store->c_str() : nullptr;
  };
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  p.lock_structure = reader.Bool(in, "lockStructure", false) ? 1 : 0;
  p.lock_windows = reader.Bool(in, "lockWindows", false) ? 1 : 0;
  p.lock_revision = reader.Bool(in, "lockRevision", false) ? 1 : 0;
  p.workbook_spin_count = reader.U32(in, "workbookSpinCount", 0U);
  p.revisions_spin_count = reader.U32(in, "revisionsSpinCount", 0U);
  p.workbook_legacy_password_hash = optional("workbookLegacyPasswordHash", &text[0]);
  p.workbook_algorithm_name = optional("workbookAlgorithmName", &text[1]);
  p.workbook_hash_value = optional("workbookHashValue", &text[2]);
  p.workbook_salt_value = optional("workbookSaltValue", &text[3]);
  p.revisions_legacy_password_hash = optional("revisionsLegacyPasswordHash", &text[4]);
  p.revisions_algorithm_name = optional("revisionsAlgorithmName", &text[5]);
  p.revisions_hash_value = optional("revisionsHashValue", &text[6]);
  p.revisions_salt_value = optional("revisionsSaltValue", &text[7]);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_workbook_set_workbook_protection(handle_, &p), env.Undefined());
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
  item.Set("legacyPasswordHash", JsStringOrNull(env, r.legacy_password_hash));
  item.Set("algorithmName", JsStringOrNull(env, r.algorithm_name));
  item.Set("hashValue", JsStringOrNull(env, r.hash_value));
  item.Set("saltValue", JsStringOrNull(env, r.salt_value));
  item.Set("securityDescriptor", JsStringOrNull(env, r.security_descriptor));
  return item;
}

// Reads the range spec at `info[idx]` and returns what `edit` makes of it.
// `handle` aliases the wrapper's member, as in `InvokeMergeEdit`.
template <typename Edit>
Napi::Value EditProtectedRange(fm_workbook_t* const& handle, const Napi::CallbackInfo& info, size_t idx, Edit edit) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
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
  (void)reader.String(in, "legacyPasswordHash", &text[1]);
  (void)reader.String(in, "algorithmName", &text[2]);
  (void)reader.String(in, "hashValue", &text[3]);
  (void)reader.String(in, "saltValue", &text[4]);
  (void)reader.String(in, "securityDescriptor", &text[5]);
  fm_protected_range_t r = FM_STRUCT_INIT(fm_protected_range_t);
  r.spin_count = reader.U32(in, "spinCount", 0U);
  if (!reader.ok()) {
    return env.Undefined();
  }
  r.name = text[0].c_str();
  r.ranges = ranges.data();
  r.range_count = static_cast<uint32_t>(ranges.size());
  r.legacy_password_hash = text[1].c_str();
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
  return SheetKeyedItemResult(info, handle_, &fm_sheet_get_protected_range_at, &fm_sheet_get_protected_range,
                              &ProtectedRangeToJs, "protectedRange");
}

Napi::Value Workbook::AddProtectedRange(const Napi::CallbackInfo& info) {
  return EditProtectedRange(handle_, info, 1, [&](fm_workbook_t* handle, const fm_protected_range_t& r) {
    const char* stored_name = nullptr;
    const fm_status_t rc = fm_sheet_add_protected_range(handle, ArgU32(info, 0), &r, &stored_name);
    return ValueOrThrow(info.Env(), rc, rc == 0 ? JsString(info.Env(), stored_name) : info.Env().Undefined());
  });
}

// Every member is optional: an omitted key keeps the stored value, and `null` clears `legacyPasswordHash`.
Napi::Value Workbook::UpdateProtectedRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const Napi::Object in = ArgObject(info, 2);
  CheckedSpecReader reader(env);
  std::string text[6];
  std::vector<fm_cell_rect_t> ranges;
  fm_protected_range_patch_t patch = FM_STRUCT_INIT(fm_protected_range_patch_t);
  patch.name_op = PatchString(reader, in, "name", false, &text[0], &patch.name);
  patch.spin_count_op = PatchU32(reader, in, "spinCount", &patch.spin_count);
  patch.legacy_password_hash_op =
      PatchString(reader, in, "legacyPasswordHash", true, &text[1], &patch.legacy_password_hash);
  patch.algorithm_name_op = PatchString(reader, in, "algorithmName", true, &text[2], &patch.algorithm_name);
  patch.hash_value_op = PatchString(reader, in, "hashValue", true, &text[3], &patch.hash_value);
  patch.salt_value_op = PatchString(reader, in, "saltValue", true, &text[4], &patch.salt_value);
  patch.security_descriptor_op =
      PatchString(reader, in, "securityDescriptor", true, &text[5], &patch.security_descriptor);
  if (PatchRects(reader, in, "ranges", &ranges)) {
    patch.ranges_op = FM_PATCH_OP_SET;
    patch.ranges = ranges.data();
    patch.range_count = static_cast<uint32_t>(ranges.size());
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  const std::string name = ArgString(info, 1);
  return ValueOrThrow(env, fm_sheet_update_protected_range(handle_, ArgU32(info, 0), name.c_str(), &patch),
                      env.Undefined());
}

Napi::Value Workbook::RemoveProtectedRange(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string name = ArgString(info, 1);
    return fm_sheet_remove_protected_range(handle, ArgU32(info, 0), name.c_str());
  });
}

// ---- Hyperlinks -----------------------------------------------------

namespace {

// Reads a `HyperlinkInput` into `hl`; `text` backs the borrowed string views. `hl->rect` keeps its value when
// `in` has no `range`, which the result reports.
bool ReadHyperlinkInput(CheckedSpecReader& reader, const Napi::Object& in, std::string (&text)[4], fm_hyperlink_t* hl) {
  const auto optional = [&](const char* key, std::string* store) {
    return reader.String(in, key, store) ? store->c_str() : nullptr;
  };
  hl->target = optional("target", &text[0]);
  hl->location = optional("location", &text[1]);
  hl->display = optional("display", &text[2]);
  hl->tooltip = optional("tooltip", &text[3]);
  Napi::Object range;
  if (!reader.Object(in, "range", &range)) {
    return false;
  }
  (void)ReadCellRect(reader, range, &hl->rect);
  return true;
}

}  // namespace

// `(sheet, row, col, link)`: `link` is a `HyperlinkInput` whose `range` (a `CellRect`, or omitted / `null` for the
// key cell) must start at `(row, col)`; `null` removes the hyperlink anchored there and does nothing when there
// is none. A missing or `null` string is no stored value.
Napi::Value Workbook::SetHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  CheckedSpecReader reader(env);
  Napi::Object in;
  if (!reader.Object(info[3], "", &in)) {
    return reader.ok() ? ValueOrThrow(env, fm_sheet_set_hyperlink(handle_, sheet, row, col, nullptr), env.Undefined())
                       : env.Undefined();
  }
  // The std::string buffers keep the borrowed `const char*` views of `fm_hyperlink_t` alive across the C call.
  std::string text[4];
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect = fm_cell_rect_t{row, col, row, col};
  (void)ReadHyperlinkInput(reader, in, text, &hl);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_hyperlink(handle_, sheet, row, col, &hl), env.Undefined());
}

// `(sheet, index, link)`: `link.range` is required; its top-left cell is the anchor.
Napi::Value Workbook::InsertHyperlinkAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  CheckedSpecReader reader(env);
  Napi::Object in;
  std::string text[4];
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  const bool has_range = reader.Object(info[2], "", &in) && ReadHyperlinkInput(reader, in, text, &hl);
  if (!reader.ok()) {
    return env.Undefined();
  }
  if (!has_range) {
    return ThrowBindingStatus(env, FM_STATUS_INVALID_ARGUMENT, "link.range is required");
  }
  return ValueOrThrow(env, fm_sheet_insert_hyperlink_at(handle_, sheet, index, &hl), env.Undefined());
}

namespace {

Napi::Object HyperlinkToJs(Napi::Env env, const fm_hyperlink_t& h) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("rect", RangeToJs(env, h.rect));
  item.Set("target", JsStringOrNull(env, h.target));
  item.Set("location", JsStringOrNull(env, h.location));
  item.Set("display", JsStringOrNull(env, h.display));
  item.Set("tooltip", JsStringOrNull(env, h.tooltip));
  return item;
}

}  // namespace

Napi::Value Workbook::GetHyperlinks(const Napi::CallbackInfo& info) {
  return SheetListResult(info, handle_, &fm_sheet_get_hyperlink_count, &fm_sheet_get_hyperlink_at, &HyperlinkToJs);
}

Napi::Value Workbook::GetHyperlink(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_hyperlink_t h = FM_STRUCT_INIT(fm_hyperlink_t);
  int32_t present = 0;
  const fm_status_t rc =
      fm_sheet_get_hyperlink(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &h);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(HyperlinkToJs(env, h)) : env.Null());
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
  item.Set("id", JsString(env, v.id));
  item.Set("ranges", ranges);
  item.Set("type", Napi::Number::New(env, v.type));
  item.Set("operator", Napi::Number::New(env, v.op));
  item.Set("errorStyle", Napi::Number::New(env, v.error_style));
  item.Set("allowBlank", JsTristate(env, v.allow_blank));
  item.Set("showInputMessage", JsTristate(env, v.show_input_message));
  item.Set("showErrorMessage", JsTristate(env, v.show_error_message));
  item.Set("showInCellDropdown", JsTristate(env, v.show_in_cell_dropdown));
  item.Set("formula1", JsStringOrNull(env, v.formula1));
  item.Set("formula2", JsStringOrNull(env, v.formula2));
  item.Set("errorTitle", JsStringOrNull(env, v.error_title));
  item.Set("errorMessage", JsStringOrNull(env, v.error_message));
  item.Set("promptTitle", JsStringOrNull(env, v.prompt_title));
  item.Set("promptMessage", JsStringOrNull(env, v.prompt_message));
  return item;
}

}  // namespace

Napi::Value Workbook::GetValidationCount(const Napi::CallbackInfo& info) {
  return SheetCountResult(info, handle_, &fm_sheet_get_validation_count);
}

Napi::Value Workbook::GetValidation(const Napi::CallbackInfo& info) {
  return SheetKeyedItemResult(info, handle_, &fm_sheet_get_validation_at, &fm_sheet_get_validation, &ValidationToJs,
                              "validation");
}

Napi::Value Workbook::AddValidation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
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

  fm_data_validation_t dv = FM_STRUCT_INIT(fm_data_validation_t);
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = ranges_buf.size();
  dv.type = reader.I32(v, "type", 0);
  dv.op = reader.I32(v, "operator", 0);
  dv.error_style = reader.I32(v, "errorStyle", 0);
  dv.allow_blank = reader.Tristate(v, "allowBlank");
  dv.show_input_message = reader.Tristate(v, "showInputMessage");
  dv.show_error_message = reader.Tristate(v, "showErrorMessage");
  dv.show_in_cell_dropdown = reader.Tristate(v, "showInCellDropdown");
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  if (!reader.ok()) {
    return env.Undefined();
  }
  const char* id = nullptr;
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, &dv, &id);
  return ValueOrThrow(env, rc, JsString(env, id));
}

// Every member is optional: an omitted key keeps the stored value, and `null` clears a formula or a text.
Napi::Value Workbook::UpdateValidation(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string id = ArgString(info, 1);
  const Napi::Object in = ArgObject(info, 2);
  CheckedSpecReader reader(env);
  std::string text[6];
  std::vector<fm_cell_rect_t> ranges;
  fm_data_validation_patch_t patch = FM_STRUCT_INIT(fm_data_validation_patch_t);
  if (PatchRects(reader, in, "ranges", &ranges)) {
    patch.ranges_op = FM_PATCH_OP_SET;
    patch.ranges = ranges.data();
    patch.range_count = static_cast<uint32_t>(ranges.size());
  }
  patch.type_op = PatchI32(reader, in, "type", &patch.type);
  patch.op_op = PatchI32(reader, in, "operator", &patch.op);
  patch.error_style_op = PatchI32(reader, in, "errorStyle", &patch.error_style);
  const auto tristate = [&](const char* key, int32_t* out) {
    int32_t flag = 0;
    const fm_patch_op_t op = PatchBool(reader, in, key, &flag);
    *out = flag != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
    return op;
  };
  patch.allow_blank_op = tristate("allowBlank", &patch.allow_blank);
  patch.show_input_message_op = tristate("showInputMessage", &patch.show_input_message);
  patch.show_error_message_op = tristate("showErrorMessage", &patch.show_error_message);
  patch.show_in_cell_dropdown_op = tristate("showInCellDropdown", &patch.show_in_cell_dropdown);
  patch.formula1_op = PatchString(reader, in, "formula1", true, &text[0], &patch.formula1);
  patch.formula2_op = PatchString(reader, in, "formula2", true, &text[1], &patch.formula2);
  patch.error_title_op = PatchString(reader, in, "errorTitle", true, &text[2], &patch.error_title);
  patch.error_message_op = PatchString(reader, in, "errorMessage", true, &text[3], &patch.error_message);
  patch.prompt_title_op = PatchString(reader, in, "promptTitle", true, &text[4], &patch.prompt_title);
  patch.prompt_message_op = PatchString(reader, in, "promptMessage", true, &text[5], &patch.prompt_message);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_validation(handle_, ArgU32(info, 0), id.c_str(), &patch), env.Undefined());
}

Napi::Value Workbook::RemoveValidationAt(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string id = ArgString(info, 1);
    return fm_sheet_remove_validation(handle, ArgU32(info, 0), id.c_str());
  });
}

Napi::Value Workbook::ClearValidations(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_validations);
}

// ---- Threaded comments and persons ----------------------------------

namespace {

// Reads the `MentionInput[]` at `arr`; `ids` backs the borrowed strings. An omitted `mentionId` is generated.
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
    c.end = reader.U32(m, "end", 0U);
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
  item.Set("created", c.created != nullptr ? JsString(env, c.created) : env.Null());
  item.Set("text", JsString(env, c.text));
  item.Set("parentId", c.parent_id != nullptr ? JsString(env, c.parent_id) : env.Null());
  item.Set("done", JsBool(env, c.done));
  Napi::Array mentions = Napi::Array::New(env, c.mentions_count);
  for (uint32_t m = 0; m < c.mentions_count; ++m) {
    Napi::Object mention = Napi::Object::New(env);
    mention.Set("personId", JsString(env, c.mentions[m].person_id));
    mention.Set("mentionId", JsString(env, c.mentions[m].mention_id));
    mention.Set("start", JsNumber(env, c.mentions[m].start));
    mention.Set("end", JsNumber(env, c.mentions[m].end));
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
  item.Set("userId", JsStringOrNull(env, p.user_id));
  item.Set("providerId", JsStringOrNull(env, p.provider_id));
  return item;
}

}  // namespace

Napi::Value Workbook::GetThreadedComments(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_threaded_comment_t>(
      arr, [&](size_t* n) { return fm_sheet_get_threaded_comment_count(handle_, sheet, n); },
      [&](size_t i, fm_threaded_comment_t* c) { return fm_sheet_get_threaded_comment_at(handle_, sheet, i, c); },
      [&](const fm_threaded_comment_t& c) { return ThreadedCommentToJs(env, c); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::AddThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 1);
  std::string id;
  std::string person_id;
  reader.String(spec, "id", &id);
  std::string created;
  std::string parent_id;
  reader.String(spec, "personId", &person_id);
  reader.String(spec, "created", &created);
  const std::string text = reader.RequiredString(spec, "text");
  // An omitted or null parentId is NULL, which opens a thread.
  const bool has_parent = reader.String(spec, "parentId", &parent_id);
  std::vector<std::string> mention_ids;
  std::vector<fm_mention_t> mentions;
  Napi::Array mention_array;
  if (reader.Array(spec, "mentions", &mention_array)) {
    ReadMentions(reader, mention_array, mention_ids, mentions);
  }
  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
  c.id = id.c_str();
  c.row = reader.U32(spec, "row", 0U);
  c.col = reader.U32(spec, "col", 0U);
  c.person_id = person_id.c_str();
  c.created = created.c_str();
  c.text = text.c_str();
  c.parent_id = has_parent ? parent_id.c_str() : nullptr;
  c.done = reader.Bool(spec, "done", false) ? 1 : 0;
  c.mentions = mentions.empty() ? nullptr : mentions.data();
  c.mentions_count = static_cast<uint32_t>(mentions.size());
  if (!reader.ok()) {
    return env.Undefined();
  }
  const char* new_id = nullptr;
  const fm_status_t rc =
      fm_sheet_add_threaded_comment(handle_, ArgU32(info, 0), FM_TEXT_OFFSET_UNIT_UTF16, &c, &new_id);
  return ValueOrThrow(env, rc, JsString(env, new_id));
}

Napi::Value Workbook::EditThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const std::string id = ArgString(info, 1);
  const Napi::Object patch = ArgObject(info, 2);
  fm_threaded_comment_patch_t out = FM_STRUCT_INIT(fm_threaded_comment_patch_t);
  std::string text;
  out.text_op = PatchString(reader, patch, "text", false, &text, &out.text);
  std::vector<std::string> mention_ids;
  std::vector<fm_mention_t> mentions;
  const Napi::Value mention_list = reader.Get(patch, "mentions");
  out.mentions_op = static_cast<fm_patch_op_t>(ClassifyPatchKey(reader, mention_list, "mentions", false));
  if (out.mentions_op == FM_PATCH_OP_SET) {
    ReadMentions(reader, mention_list, mention_ids, mentions);
  }
  out.done_op = PatchBool(reader, patch, "done", &out.done);
  if (!reader.ok()) {
    return env.Undefined();
  }
  out.mentions = mentions.empty() ? nullptr : mentions.data();
  out.mentions_count = static_cast<uint32_t>(mentions.size());
  const fm_status_t rc =
      fm_sheet_update_threaded_comment(handle_, ArgU32(info, 0), id.c_str(), FM_TEXT_OFFSET_UNIT_UTF16, &out);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::GetThreadedComment(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string id = ArgString(info, 1);
  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_threaded_comment(handle_, ArgU32(info, 0), id.c_str(), &present, &c);
  return OptionalOrNull(env, rc, present,
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
    return ThrowInvalidHandle(env);
  }
  const fm_status_t rc = CollectList<fm_person_t>(
      arr, [&](size_t* n) { return fm_workbook_get_person_count(handle_, n); },
      [&](size_t i, fm_person_t* p) { return fm_workbook_get_person_at(handle_, i, p); },
      [&](const fm_person_t& p) { return PersonToJs(env, p); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetPerson(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string id = ArgString(info, 0);
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_person(handle_, id.c_str(), &present, &p);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(PersonToJs(env, p)) : env.Null());
}

Napi::Value Workbook::AddPerson(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 0);
  std::string id;
  std::string display_name;
  std::string user_id;
  std::string provider_id;
  // An omitted `id` is generated; an omitted or `null` `userId` / `providerId` is absent.
  (void)reader.String(spec, "id", &id);
  display_name = reader.RequiredString(spec, "displayName");
  reader.String(spec, "userId", &user_id);
  reader.String(spec, "providerId", &provider_id);
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  p.id = id.c_str();
  p.display_name = display_name.c_str();
  p.user_id = user_id.c_str();
  p.provider_id = provider_id.c_str();
  if (!reader.ok()) {
    return env.Undefined();
  }
  const char* stored_id = nullptr;
  const fm_status_t rc = fm_workbook_add_person(handle_, &p, &stored_id);
  return ValueOrThrow(env, rc, rc == 0 ? JsString(env, stored_id) : env.Undefined());
}

Napi::Value Workbook::RemovePerson(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string id = ArgString(info, 0);
    return fm_workbook_remove_person(handle, id.c_str());
  });
}

}  // namespace formulon_node
