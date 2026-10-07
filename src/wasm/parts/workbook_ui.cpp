//
// JsWorkbook UI-feature surface: merges, hyperlinks, comments, and
// data-validation rules. Each accessor returns a JS-friendly value
// (`Array<...>` or `null`) so JS callers don't have to step through
// count + getter pairs.

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

JsStatus JsWorkbook::addMerge(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  JsNarrowNumericReader reader("addMerge");
  const fm_merge_range m = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  fm_status_t rc = fm_sheet_add_merge(handle_, sheet, m);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeMerge(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  JsNarrowNumericReader reader("removeMerge");
  const fm_merge_range m = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  fm_status_t rc = fm_sheet_remove_merge(handle_, sheet, m);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeMergeAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_merge_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearMerges(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_clear_merges(handle_, sheet);
  return status_from_rc(rc);
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
    arr.set("status", error_status(7000));
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
    arr.set("status", error_status(7000));
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
    out.set("status", error_status(7000));
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
    return error_status(7000);
  }
  const char* author_c = author.empty() ? nullptr : author.c_str();
  const char* text_c = text.empty() ? nullptr : text.c_str();
  fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
  return status_from_rc(rc);
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
    return error_status(7000);
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
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_hyperlink(handle_, sheet, row, col);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeHyperlinkAt(uint32_t sheet, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_hyperlink_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearHyperlinks(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_clear_hyperlinks(handle_, sheet);
  return status_from_rc(rc);
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
    return error_status(7000);
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
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_remove_validation_at(handle_, sheet, index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::clearValidations(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_sheet_clear_validations(handle_, sheet);
  return status_from_rc(rc);
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
    if (m.isUndefined() || m.isNull()) {
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
    o.set("status", error_status(kBindingInvalidHandle));
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
    return error_status(kBindingInvalidHandle);
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
    return error_status(kBindingInvalidHandle);
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
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_set_thread_resolved(handle_, sheet, threadId.c_str(), done ? 1 : 0));
}

JsStatus JsWorkbook::removeThreadedComment(uint32_t sheet, const std::string& id) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_remove_threaded_comment(handle_, sheet, id.c_str()));
}

emscripten::val JsWorkbook::getPersons() const {
  emscripten::val o = emscripten::val::object();
  emscripten::val persons = emscripten::val::array();
  o.set("persons", persons);
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
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
    return error_status(kBindingInvalidHandle);
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
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_workbook_remove_person(handle_, id.c_str()));
}

// ---- Drawing images ----------------------------------------------------

namespace {

/// Envelope for the image calls: `status` plus the header info, zeroed on failure.
emscripten::val image_info_to_val(fm_status_t rc, const fm_image_info& info) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("format", rc == 0 ? info.format : 0);
  o.set("pxWidth", rc == 0 ? info.px_width : 0U);
  o.set("pxHeight", rc == 0 ? info.px_height : 0U);
  return o;
}

emscripten::val image_info_binding_error(JsStatus status) {
  emscripten::val o = emscripten::val::object();
  o.set("status", status);
  o.set("format", 0);
  o.set("pxWidth", 0U);
  o.set("pxHeight", 0U);
  return o;
}

/// Envelope for the calls that place a picture: `objectId`, zero unless the call succeeded.
emscripten::val image_id_result(JsStatus status, uint32_t object_id) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", object_id);
  o.set("status", status);
  return o;
}

emscripten::val drawing_object_to_val(const fm_drawing_object& d) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", d.object_id);
  o.set("kind", d.kind);
  o.set("anchorKind", d.anchor_kind);
  o.set("editAs", d.edit_as);
  o.set("fromRow", d.from_row);
  o.set("fromCol", d.from_col);
  o.set("fromRowOff", static_cast<double>(d.from_row_off));
  o.set("fromColOff", static_cast<double>(d.from_col_off));
  o.set("toRow", d.to_row);
  o.set("toCol", d.to_col);
  o.set("toRowOff", static_cast<double>(d.to_row_off));
  o.set("toColOff", static_cast<double>(d.to_col_off));
  o.set("cx", static_cast<double>(d.cx));
  o.set("cy", static_cast<double>(d.cy));
  o.set("imageFormat", d.image_format);
  js_set_cstr_fields(o, {{"name", d.name}, {"descr", d.descr}, {"mediaPath", d.media_path}});
  return o;
}

}  // namespace

emscripten::val JsWorkbook::probeImage(emscripten::val bytes) const {
  if (handle_ == nullptr) {
    return image_info_to_val(kBindingInvalidHandle, fm_image_info{});
  }
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_info_binding_error(binding_error_status(kInvalidArgument, bytes_result.message.c_str()));
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  fm_image_info info{};
  const fm_status_t rc = fm_workbook_probe_image(handle_, data.data(), data.size(), &info);
  return image_info_to_val(rc, info);
}

emscripten::val JsWorkbook::listDrawingObjects(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(kBindingInvalidHandle));
    return arr;
  }
  size_t count = 0;
  fm_status_t rc = fm_sheet_drawing_object_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_drawing_object d{};
    rc = fm_sheet_drawing_object_at(handle_, sheet, i, &d);
    if (rc == 0) {
      arr.call<void>("push", drawing_object_to_val(d));
    }
  }
  arr.set("status", status_from_rc(rc));
  return arr;
}

emscripten::val JsWorkbook::getImage(uint32_t sheet, uint32_t objectId) const {
  emscripten::val o;
  if (handle_ == nullptr) {
    o = image_info_to_val(kBindingInvalidHandle, fm_image_info{});
    o.set("bytes", bytes_to_val(nullptr, 0));
    return o;
  }
  const uint8_t* data = nullptr;
  size_t len = 0;
  fm_image_info info{};
  const fm_status_t rc = fm_sheet_get_image(handle_, sheet, objectId, &data, &len, &info);
  o = image_info_to_val(rc, info);
  o.set("bytes", bytes_to_val(rc == 0 ? data : nullptr, rc == 0 ? len : 0));
  return o;
}

emscripten::val JsWorkbook::insertImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    return image_id_result(error_status(kBindingInvalidHandle), 0U);
  }
  std::string name_store;
  std::string descr_store;
  JsNarrowNumericReader reader("insertImage");
  const emscripten::val options = opts.isUndefined() || opts.isNull() ? emscripten::val::object() : opts;
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_id_result(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), 0U);
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  fm_image_insert ins{};
  ins.name = reader.optional_string(options, "name", name_store, "image.name");
  ins.descr = reader.optional_string(options, "descr", descr_store, "image.descr");
  ins.anchor_kind = reader.i32(options, "anchorKind", FM_ANCHOR_KIND_ONE_CELL, "image.anchorKind");
  ins.edit_as = reader.i32(options, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "image.editAs");
  ins.row = reader.u32(options, "row", 0U, "image.row");
  ins.col = reader.u32(options, "col", 0U, "image.col");
  ins.row_off_emu = reader.i64(options, "rowOffEmu", 0, "image.rowOffEmu");
  ins.col_off_emu = reader.i64(options, "colOffEmu", 0, "image.colOffEmu");
  ins.width_emu = reader.i64(options, "widthEmu", 0, "image.widthEmu");
  ins.height_emu = reader.i64(options, "heightEmu", 0, "image.heightEmu");
  if (!reader.ok()) {
    return image_id_result(binding_error_status(kInvalidArgument, reader.message().c_str()), 0U);
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_insert_image(handle_, sheet, data.data(), data.size(), &ins, &id);
  return image_id_result(status_from_rc(rc), rc == 0 ? id : 0U);
}

JsStatus JsWorkbook::removeImage(uint32_t sheet, uint32_t objectId) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_remove_image(handle_, sheet, objectId));
}

JsStatus JsWorkbook::setImageAnchor(uint32_t sheet, uint32_t objectId, emscripten::val placement) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setImageAnchor");
  const emscripten::val options = placement.isUndefined() || placement.isNull() ? emscripten::val::object() : placement;
  fm_image_anchor anchor{};
  anchor.anchor_kind = reader.i32(options, "anchorKind", FM_ANCHOR_KIND_ONE_CELL, "placement.anchorKind");
  anchor.edit_as = reader.i32(options, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "placement.editAs");
  anchor.row = reader.u32(options, "row", 0U, "placement.row");
  anchor.col = reader.u32(options, "col", 0U, "placement.col");
  anchor.row_off_emu = reader.i64(options, "rowOffEmu", 0, "placement.rowOffEmu");
  anchor.col_off_emu = reader.i64(options, "colOffEmu", 0, "placement.colOffEmu");
  anchor.width_emu = reader.i64(options, "widthEmu", 0, "placement.widthEmu");
  anchor.height_emu = reader.i64(options, "heightEmu", 0, "placement.heightEmu");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_set_image_anchor(handle_, sheet, objectId, &anchor));
}

JsStatus JsWorkbook::setImageZOrder(uint32_t sheet, uint32_t objectId, uint32_t index) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_set_image_z_order(handle_, sheet, objectId, index));
}

emscripten::val JsWorkbook::snapshotImage(uint32_t sheet, uint32_t objectId) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(kBindingInvalidHandle));
    o.set("bytes", bytes_to_val(nullptr, 0));
    return o;
  }
  uint8_t* data = nullptr;
  size_t len = 0;
  const fm_status_t rc = fm_sheet_snapshot_image(handle_, sheet, objectId, &data, &len);
  o.set("status", status_from_rc(rc));
  o.set("bytes", bytes_to_val(rc == 0 ? data : nullptr, rc == 0 ? len : 0));
  fm_buffer_free(data);
  return o;
}

emscripten::val JsWorkbook::restoreImage(uint32_t sheet, emscripten::val bytes, emscripten::val opts) {
  if (handle_ == nullptr) {
    return image_id_result(error_status(kBindingInvalidHandle), 0U);
  }
  JsNarrowNumericReader reader("restoreImage");
  const emscripten::val options = opts.isUndefined() || opts.isNull() ? emscripten::val::object() : opts;
  const JsBytesReadResult bytes_result = val_to_bytes_checked(bytes);
  if (!bytes_result.ok) {
    return image_id_result(binding_error_status(kInvalidArgument, bytes_result.message.c_str()), 0U);
  }
  const std::vector<uint8_t>& data = bytes_result.bytes;
  const uint32_t flags = reader.boolean(options, "newId", false, "options.newId") ? FM_IMAGE_RESTORE_NEW_ID : 0U;
  if (!reader.ok()) {
    return image_id_result(binding_error_status(kInvalidArgument, reader.message().c_str()), 0U);
  }
  uint32_t id = 0;
  const fm_status_t rc = fm_sheet_restore_image(handle_, sheet, data.data(), data.size(), flags, &id);
  return image_id_result(status_from_rc(rc), rc == 0 ? id : 0U);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
