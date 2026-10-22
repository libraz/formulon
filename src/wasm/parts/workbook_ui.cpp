//
// JsWorkbook UI-feature surface: merges, hyperlinks, comments, threaded
// comments and persons, and data-validation rules. Each accessor returns a
// JS-friendly value (`Array<...>` or `null`) so JS callers don't have to
// step through count + getter pairs.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

// ---- Merges ------------------------------------------------------------

namespace {

// Reads `range` and hands it to `edit`, the add or remove half of the merge API.
JsStatus edit_merge(fm_workbook_t* wb, uint32_t sheet, const emscripten::val& range, const char* operation,
                    fm_status_t (*edit)(fm_workbook_t*, uint32_t, fm_merge_range)) {
  if (wb == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader(operation);
  const fm_merge_range m = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(edit(wb, sheet, m));
}

}  // namespace

JsStatus JsWorkbook::addMerge(uint32_t sheet, emscripten::val range) {
  return edit_merge(handle_, sheet, range, "addMerge", &fm_sheet_add_merge);
}

JsStatus JsWorkbook::removeMerge(uint32_t sheet, emscripten::val range) {
  return edit_merge(handle_, sheet, range, "removeMerge", &fm_sheet_remove_merge);
}

JsStatus JsWorkbook::removeMergeAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_merge_at(handle_, sheet, index));
}

JsStatus JsWorkbook::clearMerges(uint32_t sheet) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_merges(handle_, sheet));
}

namespace {

// Reads list element `index` of `sheet` into `*out`; the per-kind half of
// `sheet_list`.
using SheetListItemFn = fm_status_t (*)(fm_workbook_t*, uint32_t, uint32_t, emscripten::val*);

// Enumerates a per-sheet list through its count / at-index pair into a
// `ListResult` array, stopping at the first failed element read.
emscripten::val sheet_list(fm_workbook_t* wb, uint32_t sheet,
                           fm_status_t (*count_fn)(fm_workbook_t*, uint32_t, uint32_t*), SheetListItemFn item_fn) {
  emscripten::val arr = emscripten::val::array();
  if (wb == nullptr) {
    arr.set("status", invalid_handle_status());
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = count_fn(wb, sheet, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  for (uint32_t i = 0; i < count; ++i) {
    emscripten::val item = emscripten::val::undefined();
    rc = item_fn(wb, sheet, i, &item);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
    }
    arr.set(i, item);
  }
  arr.set("status", ok_status());
  return arr;
}

fm_status_t merge_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_merge_range m{};
  const fm_status_t rc = fm_sheet_get_merge_at(wb, sheet, index, &m);
  if (rc != 0) {
    return rc;
  }
  *out = merge_range_to_val(m);
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getMerges(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_merge_count, &merge_item);
}

emscripten::val JsWorkbook::getMergesInRange(uint32_t sheet, emscripten::val range) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", invalid_handle_status());
    return arr;
  }
  JsNarrowNumericReader reader("getMergesInRange");
  const fm_merge_range query = js_pull_range(range, &reader);
  if (!reader.ok()) {
    arr.set("status", binding_error_status(kInvalidArgument, reader.message().c_str()));
    return arr;
  }
  uint32_t total = 0;
  fm_status_t rc = fm_sheet_merges_in_range(handle_, sheet, query, nullptr, 0, &total);
  std::vector<fm_merge_range> found(rc == 0 ? total : 0);
  if (rc == 0 && total > 0) {
    uint32_t written = 0;
    rc = fm_sheet_merges_in_range(handle_, sheet, query, found.data(), total, &written);
    found.resize(rc == 0 ? written : 0);
  }
  if (rc != 0) {
    arr.set("status", error_status(rc));
    return arr;
  }
  for (uint32_t i = 0; i < found.size(); ++i) {
    arr.set(i, merge_range_to_val(found[i]));
  }
  arr.set("status", ok_status());
  return arr;
}

// ---- Comments ----------------------------------------------------------

emscripten::val JsWorkbook::getComment(uint32_t sheet, uint32_t row, uint32_t col) const {
  if (handle_ == nullptr) {
    return emscripten::val::null();
  }
  fm_comment c{};
  if (fm_sheet_get_comment_at(handle_, sheet, row, col, &c) != 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_fields(o, {{"author", c.author}, {"text", c.text}});
  return o;
}

emscripten::val JsWorkbook::getCommentResult(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val out = emscripten::val::object();
  if (handle_ == nullptr) {
    out.set("status", invalid_handle_status());
    out.set("comment", emscripten::val::null());
    return out;
  }
  fm_comment c{};
  const fm_status_t rc = fm_sheet_get_comment_at(handle_, sheet, row, col, &c);
  if (rc != 0) {
    out.set("status", status_from_rc(rc));
    out.set("comment", emscripten::val::null());
    return out;
  }
  emscripten::val comment = emscripten::val::object();
  js_set_cstr_fields(comment, {{"author", c.author}, {"text", c.text}});
  out.set("status", ok_status());
  out.set("comment", comment);
  return out;
}

namespace {

fm_status_t comment_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_comment c{};
  const fm_status_t rc = fm_sheet_get_comment_at_index(wb, sheet, index, &c);
  if (rc != 0) {
    return rc;
  }
  emscripten::val o = emscripten::val::object();
  o.set("row", c.row);
  o.set("col", c.col);
  js_set_cstr_fields(o, {{"author", c.author}, {"text", c.text}});
  *out = o;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getComments(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_comment_count, &comment_item);
}

JsStatus JsWorkbook::setComment(uint32_t sheet, uint32_t row, uint32_t col, const std::string& author,
                                const std::string& text) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  const char* author_c = author.empty() ? nullptr : author.c_str();
  const char* text_c = text.empty() ? nullptr : text.c_str();
  fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
  return status_from_rc(rc);
}

JsNumberResult JsWorkbook::getCommentVisible(uint32_t sheet, uint32_t row, uint32_t col) const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0);
  }
  int32_t visible = 0;
  const fm_status_t rc = fm_sheet_get_comment_visible(handle_, sheet, row, col, &visible);
  return number_result(rc, visible);
}

JsStatus JsWorkbook::setCommentVisible(uint32_t sheet, uint32_t row, uint32_t col, bool visible) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_comment_visible(handle_, sheet, row, col, visible ? 1 : 0));
}

// ---- Protection records ------------------------------------------------

emscripten::val JsWorkbook::getWorkbookProtection() const {
  emscripten::val out = emscripten::val::object();
  fm_workbook_protection_t p{};
  if (handle_ == nullptr) {
    out.set("status", invalid_handle_status());
    out.set("protection", emscripten::val::null());
    return out;
  }
  const fm_status_t rc = fm_workbook_get_protection(handle_, &p);
  if (rc != 0) {
    out.set("status", status_from_rc(rc));
    out.set("protection", emscripten::val::null());
    return out;
  }
  emscripten::val protection = emscripten::val::object();
  protection.set("lockStructure", p.lock_structure);
  protection.set("lockWindows", p.lock_windows);
  protection.set("lockRevision", p.lock_revision);
  protection.set("workbookSpinCount", p.workbook_spin_count);
  protection.set("revisionsSpinCount", p.revisions_spin_count);
  js_set_cstr_fields(protection, {{"workbookPassword", p.workbook_password},
                                  {"workbookAlgorithmName", p.workbook_algorithm_name},
                                  {"workbookHashValue", p.workbook_hash_value},
                                  {"workbookSaltValue", p.workbook_salt_value},
                                  {"revisionsPassword", p.revisions_password},
                                  {"revisionsAlgorithmName", p.revisions_algorithm_name},
                                  {"revisionsHashValue", p.revisions_hash_value},
                                  {"revisionsSaltValue", p.revisions_salt_value}});
  out.set("status", ok_status());
  out.set("protection", protection);
  return out;
}

JsStatus JsWorkbook::setWorkbookProtection(emscripten::val protection) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setWorkbookProtection");
  fm_workbook_protection_t p{};
  p.lock_structure = reader.i32(protection, "lockStructure", 0, "protection.lockStructure");
  p.lock_windows = reader.i32(protection, "lockWindows", 0, "protection.lockWindows");
  p.lock_revision = reader.i32(protection, "lockRevision", 0, "protection.lockRevision");
  p.workbook_spin_count = reader.u32(protection, "workbookSpinCount", 0U, "protection.workbookSpinCount");
  p.revisions_spin_count = reader.u32(protection, "revisionsSpinCount", 0U, "protection.revisionsSpinCount");
  std::string text[8];
  p.workbook_password = reader.optional_string(protection, "workbookPassword", text[0], "protection.workbookPassword");
  p.workbook_algorithm_name =
      reader.optional_string(protection, "workbookAlgorithmName", text[1], "protection.workbookAlgorithmName");
  p.workbook_hash_value =
      reader.optional_string(protection, "workbookHashValue", text[2], "protection.workbookHashValue");
  p.workbook_salt_value =
      reader.optional_string(protection, "workbookSaltValue", text[3], "protection.workbookSaltValue");
  p.revisions_password =
      reader.optional_string(protection, "revisionsPassword", text[4], "protection.revisionsPassword");
  p.revisions_algorithm_name =
      reader.optional_string(protection, "revisionsAlgorithmName", text[5], "protection.revisionsAlgorithmName");
  p.revisions_hash_value =
      reader.optional_string(protection, "revisionsHashValue", text[6], "protection.revisionsHashValue");
  p.revisions_salt_value =
      reader.optional_string(protection, "revisionsSaltValue", text[7], "protection.revisionsSaltValue");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_workbook_set_protection(handle_, &p));
}

namespace {

fm_status_t protected_range_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_protected_range_t r{};
  const fm_status_t rc = fm_sheet_get_protected_range(wb, sheet, index, &r);
  if (rc != 0) {
    return rc;
  }
  emscripten::val o = emscripten::val::object();
  o.set("spinCount", r.spin_count);
  js_set_cstr_fields(o, {{"name", r.name},
                         {"sqref", r.sqref},
                         {"password", r.password},
                         {"algorithmName", r.algorithm_name},
                         {"hashValue", r.hash_value},
                         {"saltValue", r.salt_value},
                         {"securityDescriptor", r.security_descriptor}});
  *out = o;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getProtectedRanges(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_protected_range_count, &protected_range_item);
}

namespace {

// Hands the range described by `range` to `edit`, the add or update half of the protected-range API.
template <typename Edit>
JsStatus edit_protected_range(fm_workbook_t* wb, const emscripten::val& range, const char* operation, Edit edit) {
  if (wb == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader(operation);
  std::string text[7];
  fm_protected_range_t r{};
  r.spin_count = reader.u32(range, "spinCount", 0U, "range.spinCount");
  r.name = reader.optional_string(range, "name", text[0], "range.name");
  r.sqref = reader.optional_string(range, "sqref", text[1], "range.sqref");
  r.password = reader.optional_string(range, "password", text[2], "range.password");
  r.algorithm_name = reader.optional_string(range, "algorithmName", text[3], "range.algorithmName");
  r.hash_value = reader.optional_string(range, "hashValue", text[4], "range.hashValue");
  r.salt_value = reader.optional_string(range, "saltValue", text[5], "range.saltValue");
  r.security_descriptor = reader.optional_string(range, "securityDescriptor", text[6], "range.securityDescriptor");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(edit(wb, r));
}

}  // namespace

JsStatus JsWorkbook::addProtectedRange(uint32_t sheet, emscripten::val range) {
  return edit_protected_range(handle_, range, "addProtectedRange",
                              [sheet](fm_workbook_t* wb, const fm_protected_range_t& r) {
                                return fm_sheet_add_protected_range(wb, sheet, &r);
                              });
}

JsStatus JsWorkbook::updateProtectedRange(uint32_t sheet, uint32_t index, emscripten::val range) {
  return edit_protected_range(handle_, range, "updateProtectedRange",
                              [sheet, index](fm_workbook_t* wb, const fm_protected_range_t& r) {
                                return fm_sheet_update_protected_range(wb, sheet, index, &r);
                              });
}

JsStatus JsWorkbook::removeProtectedRange(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_protected_range(handle_, sheet, index));
}

// ---- Hyperlinks --------------------------------------------------------

JsStatus JsWorkbook::addHyperlink(uint32_t sheet, uint32_t row, uint32_t col, const std::string& target,
                                  const std::string& display, const std::string& tooltip, const std::string& location) {
  return addHyperlinkRange(sheet, row, col, row, col, target, display, tooltip, location);
}

JsStatus JsWorkbook::addHyperlinkRange(uint32_t sheet, uint32_t row, uint32_t col, uint32_t lastRow, uint32_t lastCol,
                                       const std::string& target, const std::string& display,
                                       const std::string& tooltip, const std::string& location) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_hyperlink hl{};
  hl.row = row;
  hl.col = col;
  hl.last_row = lastRow;
  hl.last_col = lastCol;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  fm_status_t rc = fm_sheet_add_hyperlink(handle_, sheet, hl);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeHyperlink(uint32_t sheet, uint32_t row, uint32_t col) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_hyperlink(handle_, sheet, row, col));
}

JsStatus JsWorkbook::removeHyperlinkAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_hyperlink_at(handle_, sheet, index));
}

JsStatus JsWorkbook::clearHyperlinks(uint32_t sheet) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_hyperlinks(handle_, sheet));
}

namespace {

fm_status_t hyperlink_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_hyperlink h{};
  const fm_status_t rc = fm_sheet_get_hyperlink_at(wb, sheet, index, &h);
  if (rc != 0) {
    return rc;
  }
  emscripten::val item = emscripten::val::object();
  item.set("row", h.row);
  item.set("col", h.col);
  item.set("lastRow", h.last_row);
  item.set("lastCol", h.last_col);
  js_set_cstr_fields(item,
                     {{"target", h.target}, {"location", h.location}, {"display", h.display}, {"tooltip", h.tooltip}});
  *out = item;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getHyperlinks(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_hyperlink_count, &hyperlink_item);
}

// ---- Data validations --------------------------------------------------

namespace {

fm_status_t validation_item(fm_workbook_t* wb, uint32_t sheet, uint32_t index, emscripten::val* out) {
  fm_data_validation v{};
  const fm_status_t rc = fm_sheet_get_validation_at(wb, sheet, index, &v);
  if (rc != 0) {
    return rc;
  }
  emscripten::val item = emscripten::val::object();
  emscripten::val ranges = emscripten::val::array();
  for (uint32_t r = 0; r < v.range_count; ++r) {
    ranges.set(r, merge_range_to_val(v.ranges[r]));
  }
  item.set("ranges", ranges);
  item.set("type", v.type);
  item.set("op", v.op);
  item.set("errorStyle", v.error_style);
  item.set("allowBlank", v.allow_blank != 0);
  item.set("showInputMessage", v.show_input_message != 0);
  item.set("showErrorMessage", v.show_error_message != 0);
  item.set("showDropDown", v.show_dropdown != 0);
  js_set_cstr_fields(item, {{"formula1", v.formula1},
                            {"formula2", v.formula2},
                            {"errorTitle", v.error_title},
                            {"errorMessage", v.error_message},
                            {"promptTitle", v.prompt_title},
                            {"promptMessage", v.prompt_message}});
  *out = item;
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getValidations(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_validation_count, &validation_item);
}

JsStatus JsWorkbook::addValidation(uint32_t sheet, emscripten::val v) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  JsNarrowNumericReader reader("addValidation");
  const std::vector<fm_merge_range> ranges_buf = js_pull_ranges(v, "ranges", &reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  const std::string formula1 = reader.string(v, "formula1", "validation.formula1");
  const std::string formula2 = reader.string(v, "formula2", "validation.formula2");
  const std::string error_title = reader.string(v, "errorTitle", "validation.errorTitle");
  const std::string error_message = reader.string(v, "errorMessage", "validation.errorMessage");
  const std::string prompt_title = reader.string(v, "promptTitle", "validation.promptTitle");
  const std::string prompt_message = reader.string(v, "promptMessage", "validation.promptMessage");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }

  fm_data_validation dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = static_cast<uint32_t>(ranges_buf.size());
  dv.type = reader.u8(v, "type", 0U, "validation.type");
  dv.op = reader.u8(v, "op", 0U, "validation.op");
  dv.error_style = reader.u8(v, "errorStyle", 0U, "validation.errorStyle");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  dv.allow_blank = reader.boolean(v, "allowBlank", false, "validation.allowBlank") ? 1 : 0;
  dv.show_input_message = reader.boolean(v, "showInputMessage", false, "validation.showInputMessage") ? 1 : 0;
  dv.show_error_message = reader.boolean(v, "showErrorMessage", false, "validation.showErrorMessage") ? 1 : 0;
  dv.show_dropdown = reader.boolean(v, "showDropDown", true, "validation.showDropDown") ? 1 : 0;
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  fm_status_t rc = fm_sheet_add_validation(handle_, sheet, dv);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeValidationAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_validation_at(handle_, sheet, index));
}

JsStatus JsWorkbook::clearValidations(uint32_t sheet) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_clear_validations(handle_, sheet));
}

// ---- Threaded comments and persons -------------------------------------

namespace {

/// Owns the strings behind a list of `fm_mention`.
struct MentionStore {
  std::vector<std::string> person_ids;
  std::vector<std::string> mention_ids;
  std::vector<fm_mention> items;
};

void pull_mentions(const emscripten::val& list, MentionStore& st, JsNarrowNumericReader& reader) {
  const uint32_t n = reader.is_array(list, "mentions") ? reader.length(list, "mentions.length") : 0U;
  if (!reader.ok()) {
    return;
  }
  st.person_ids.resize(n);
  st.mention_ids.resize(n);
  st.items.resize(n);
  for (uint32_t i = 0; i < n; ++i) {
    if (!reader.ok()) {
      break;
    }
    const emscripten::val m = reader.array_element(list, i, "mentions[]");
    if (!reader.ok()) {
      break;
    }
    if (!js_value_present(m)) {
      continue;
    }
    st.person_ids[i] = reader.string(m, "personId", "mention.personId");
    st.mention_ids[i] = reader.string(m, "mentionId", "mention.mentionId");
    const uint32_t start = reader.u32(m, "start", 0U, "mention.start");
    const uint32_t length = reader.u32(m, "length", 0U, "mention.length");
    st.items[i] = {st.person_ids[i].c_str(), st.mention_ids[i].c_str(), start, length};
    if (!reader.ok()) {
      break;
    }
  }
}

emscripten::val mention_to_val(const fm_mention& m) {
  emscripten::val o = emscripten::val::object();
  o.set("start", m.start);
  o.set("length", m.length);
  js_set_cstr_fields(o, {{"personId", m.person_id}, {"mentionId", m.mention_id}});
  return o;
}

emscripten::val threaded_comment_to_val(const fm_threaded_comment& c) {
  emscripten::val o = emscripten::val::object();
  o.set("row", c.row);
  o.set("col", c.col);
  o.set("done", c.done != 0);
  js_set_cstr_fields(
      o,
      {{"id", c.id}, {"personId", c.person_id}, {"created", c.created}, {"text", c.text}, {"parentId", c.parent_id}});
  emscripten::val mentions = emscripten::val::array();
  for (uint32_t i = 0; i < c.mention_count; ++i) {
    mentions.set(i, mention_to_val(c.mentions[i]));
  }
  o.set("mentions", mentions);
  return o;
}

emscripten::val person_to_val(const fm_person& p) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr_fields(
      o, {{"id", p.id}, {"displayName", p.display_name}, {"userId", p.user_id}, {"providerId", p.provider_id}});
  return o;
}

}  // namespace

emscripten::val JsWorkbook::getThreadedComments(uint32_t sheet) const {
  emscripten::val o = emscripten::val::object();
  emscripten::val comments = emscripten::val::array();
  o.set("comments", comments);
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    return o;
  }
  size_t count = 0;
  fm_status_t rc = fm_sheet_threaded_comment_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_threaded_comment c{};
    rc = fm_sheet_threaded_comment_at(handle_, sheet, i, &c);
    if (rc == 0) {
      comments.call<void>("push", threaded_comment_to_val(c));
    }
  }
  o.set("status", status_from_rc(rc));
  return o;
}

JsStatus JsWorkbook::addThreadedComment(uint32_t sheet, emscripten::val comment) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("addThreadedComment");
  const std::string id = reader.string(comment, "id", "comment.id");
  const std::string person_id = reader.string(comment, "personId", "comment.personId");
  const std::string created = reader.string(comment, "created", "comment.created");
  const std::string text = reader.string(comment, "text", "comment.text");
  const std::string parent_id = reader.string(comment, "parentId", "comment.parentId");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  MentionStore mentions;
  pull_mentions(reader.value(comment, "mentions", "comment.mentions"), mentions, reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  fm_threaded_comment c{};
  c.id = id.c_str();
  c.row = reader.u32(comment, "row", 0U, "comment.row");
  c.col = reader.u32(comment, "col", 0U, "comment.col");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  c.person_id = person_id.c_str();
  c.created = created.c_str();
  c.text = text.c_str();
  c.parent_id = parent_id.c_str();
  c.done = reader.boolean(comment, "done", false, "done") ? 1 : 0;
  c.mentions = mentions.items.empty() ? nullptr : mentions.items.data();
  c.mention_count = static_cast<uint32_t>(mentions.items.size());
  return status_from_rc(fm_sheet_add_threaded_comment(handle_, sheet, &c));
}

JsStatus JsWorkbook::editThreadedComment(uint32_t sheet, const std::string& id, const std::string& text,
                                         emscripten::val mentions) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("editThreadedComment");
  MentionStore store;
  pull_mentions(mentions, store, reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_edit_threaded_comment(handle_, sheet, id.c_str(), text.c_str(),
                                                       store.items.empty() ? nullptr : store.items.data(),
                                                       static_cast<uint32_t>(store.items.size())));
}

JsStatus JsWorkbook::setThreadResolved(uint32_t sheet, const std::string& threadId, bool done) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_set_thread_resolved(handle_, sheet, threadId.c_str(), done ? 1 : 0));
}

JsStatus JsWorkbook::removeThreadedComment(uint32_t sheet, const std::string& id) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_sheet_remove_threaded_comment(handle_, sheet, id.c_str()));
}

emscripten::val JsWorkbook::getPersons() const {
  emscripten::val o = emscripten::val::object();
  emscripten::val persons = emscripten::val::array();
  o.set("persons", persons);
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    return o;
  }
  size_t count = 0;
  fm_status_t rc = fm_workbook_person_count(handle_, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_person p{};
    rc = fm_workbook_person_at(handle_, i, &p);
    if (rc == 0) {
      persons.call<void>("push", person_to_val(p));
    }
  }
  o.set("status", status_from_rc(rc));
  return o;
}

JsStatus JsWorkbook::addPerson(emscripten::val person) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("addPerson");
  const std::string id = reader.string(person, "id", "person.id");
  const std::string display_name = reader.string(person, "displayName", "person.displayName");
  const std::string user_id = reader.string(person, "userId", "person.userId");
  const std::string provider_id = reader.string(person, "providerId", "person.providerId");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  const fm_person p = {id.c_str(), display_name.c_str(), user_id.c_str(), provider_id.c_str()};
  return status_from_rc(fm_workbook_add_person(handle_, &p));
}

JsStatus JsWorkbook::removePerson(const std::string& id) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  return status_from_rc(fm_workbook_remove_person(handle_, id.c_str()));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
