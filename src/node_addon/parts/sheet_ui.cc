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

// Shared bodies of the per-sheet Status entries taking `(sheet)`,
// `(sheet, index)` and `(sheet, range)`.
using SheetOpFn = fm_status_t (*)(fm_workbook_t*, uint32_t);
using SheetIndexOpFn = fm_status_t (*)(fm_workbook_t*, uint32_t, uint32_t);
using MergeEditFn = fm_status_t (*)(fm_workbook_t*, uint32_t, fm_merge_range);

Napi::Value InvokeSheetOp(const Napi::CallbackInfo& info, fm_workbook_t* handle, SheetOpFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  fm_status_t rc = fn(handle, sheet);
  return MakeStatus(env, rc);
}

Napi::Value InvokeSheetIndexOp(const Napi::CallbackInfo& info, fm_workbook_t* handle, SheetIndexOpFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  const uint32_t index = Workbook::ArgU32(info, 1);
  fm_status_t rc = fn(handle, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value InvokeMergeEdit(const Napi::CallbackInfo& info, fm_workbook_t* handle, MergeEditFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeErrorStatus(env, kBindingInvalidHandle);
  }
  CheckedSpecReader reader(env);
  fm_merge_range range{};
  if (!MergeRangeArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  const uint32_t sheet = Workbook::ArgU32(info, 0);
  fm_status_t rc = fn(handle, sheet, range);
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
  return InvokeSheetIndexOp(info, handle_, &fm_sheet_remove_merge_at);
}

Napi::Value Workbook::ClearMerges(const Napi::CallbackInfo& info) {
  return InvokeSheetOp(info, handle_, &fm_sheet_clear_merges);
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
  fm_merge_range range{};
  if (!MergeRangeArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  const uint32_t sheet = ArgU32(info, 0);
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
    arr.Set(i, RangeToJs(env, merges[i]));
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
  o.Set("author", JsString(env, c.author));
  o.Set("text", JsString(env, c.text));
  return o;
}

Napi::Value Workbook::GetCommentResult(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "comment", env.Null());
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  fm_comment c{};
  const fm_status_t rc = fm_sheet_get_comment_at(handle_, sheet, row, col, &c);
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "comment", env.Null());
  }
  Napi::Object comment = Napi::Object::New(env);
  comment.Set("author", JsString(env, c.author));
  comment.Set("text", JsString(env, c.text));
  return MakeFieldResult(env, MakeOkStatus(env), "comment", comment);
}

namespace {

Napi::Object CommentToJs(Napi::Env env, const fm_comment& c) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("row", Napi::Number::New(env, c.row));
  item.Set("col", Napi::Number::New(env, c.col));
  item.Set("author", JsString(env, c.author));
  item.Set("text", JsString(env, c.text));
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
  return InvokeSheetIndexOp(info, handle_, &fm_sheet_remove_hyperlink_at);
}

Napi::Value Workbook::ClearHyperlinks(const Napi::CallbackInfo& info) {
  return InvokeSheetOp(info, handle_, &fm_sheet_clear_hyperlinks);
}

// ---- Validations ----------------------------------------------------

namespace {

Napi::Object ValidationToJs(Napi::Env env, const fm_data_validation& v) {
  Napi::Object item = Napi::Object::New(env);
  Napi::Array ranges = Napi::Array::New(env);
  for (uint32_t r = 0; r < v.range_count; ++r) {
    ranges.Set(r, RangeToJs(env, v.ranges[r]));
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
  CheckedSpecReader reader(env);

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  std::vector<fm_merge_range> ranges_buf;
  Napi::Array ranges_arr;
  if (reader.Array(v, "ranges", &ranges_arr)) {
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
      fm_merge_range parsed{};
      if (!ReadMergeRange(reader, range, &parsed)) {
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

  fm_data_validation dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = static_cast<uint32_t>(ranges_buf.size());
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
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, dv);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::RemoveValidationAt(const Napi::CallbackInfo& info) {
  return InvokeSheetIndexOp(info, handle_, &fm_sheet_remove_validation_at);
}

Napi::Value Workbook::ClearValidations(const Napi::CallbackInfo& info) {
  return InvokeSheetOp(info, handle_, &fm_sheet_clear_validations);
}

// ---- Threaded comments and persons ----------------------------------

namespace {

// Reads the `Mention[]` at `arr`; `ids` backs the borrowed strings.
bool ReadMentions(CheckedSpecReader& reader, const Napi::Value& arr, std::vector<std::string>& ids,
                  std::vector<fm_mention>& out) {
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
    fm_mention c{};
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
  CheckedSpecReader reader(env);
  const Napi::Object spec = info[1].As<Napi::Object>();
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
  std::vector<fm_mention> mentions;
  Napi::Array mention_array;
  if (reader.Array(spec, "mentions", &mention_array)) {
    ReadMentions(reader, mention_array, mention_ids, mentions);
  }
  fm_threaded_comment c{};
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
  std::vector<fm_mention> mentions;
  if (info.Length() > 3) {
    ReadMentions(reader, info[3], mention_ids, mentions);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
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
    item.Set("id", JsString(env, p.id));
    item.Set("displayName", JsString(env, p.display_name));
    item.Set("userId", JsString(env, p.user_id));
    item.Set("providerId", JsString(env, p.provider_id));
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
  CheckedSpecReader reader(env);
  const Napi::Object spec = info[0].As<Napi::Object>();
  std::string id;
  std::string display_name;
  std::string user_id;
  std::string provider_id;
  reader.String(spec, "id", &id);
  reader.String(spec, "displayName", &display_name);
  reader.String(spec, "userId", &user_id);
  reader.String(spec, "providerId", &provider_id);
  fm_person p{};
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
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string id = ArgString(info, 0);
  return MakeStatus(env, fm_workbook_remove_person(handle_, id.c_str()));
}

}  // namespace formulon_node
