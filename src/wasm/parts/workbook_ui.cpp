//
// JsWorkbook UI-feature surface: merges, hyperlinks, comments, threaded
// comments and persons, and data-validation rules. Each accessor returns a
// JS-friendly value (`Array<...>` or `null`) so JS callers don't have to
// step through count + getter pairs.

#include <emscripten/val.h>

#include <cstddef>
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
void edit_merge(fm_workbook_t* wb, uint32_t sheet, const emscripten::val& range, const char* operation,
                fm_status_t (*edit)(fm_workbook_t*, std::size_t, const fm_cell_rect_t*)) {
  if (wb == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader(operation);
  const fm_cell_rect_t m = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(edit(wb, sheet, &m));
}

}  // namespace

void JsWorkbook::addMerge(uint32_t sheet, emscripten::val range) {
  edit_merge(handle_, sheet, range, "addMerge", &fm_sheet_add_merge);
}

void JsWorkbook::removeMerge(uint32_t sheet, emscripten::val range) {
  edit_merge(handle_, sheet, range, "removeMerge", &fm_sheet_remove_merge);
}

void JsWorkbook::removeMergeAt(uint32_t sheet, uint32_t index) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_merge_at(handle_, sheet, index) : kBindingInvalidHandle);
}

void JsWorkbook::clearMerges(uint32_t sheet) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_merges(handle_, sheet) : kBindingInvalidHandle);
}

namespace {

// Reads list element `index` of `sheet` into `*out`; the per-kind half of
// `sheet_list`.
using SheetListItemFn = fm_status_t (*)(const fm_workbook_t*, uint32_t, std::size_t, emscripten::val*);

// Enumerates a per-sheet list through its count / at-index pair into an
// array, stopping at the first failed element read (`undefined` after
// recording it).
emscripten::val sheet_list(const fm_workbook_t* wb, uint32_t sheet,
                           fm_status_t (*count_fn)(const fm_workbook_t*, std::size_t, std::size_t*),
                           SheetListItemFn item_fn) {
  if (wb == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  std::size_t count = 0;
  if (!succeeded(count_fn(wb, sheet, &count))) {
    return emscripten::val::undefined();
  }
  emscripten::val arr = emscripten::val::array();
  for (std::size_t i = 0; i < count; ++i) {
    emscripten::val item = emscripten::val::undefined();
    if (!succeeded(item_fn(wb, sheet, i, &item))) {
      return emscripten::val::undefined();
    }
    arr.set(i, item);
  }
  return arr;
}

fm_status_t merge_item(const fm_workbook_t* wb, uint32_t sheet, std::size_t index, emscripten::val* out) {
  fm_cell_rect_t m{};
  const fm_status_t rc = fm_sheet_get_merge_at(wb, sheet, index, &m);
  if (rc != 0) {
    return rc;
  }
  *out = cell_rect_to_val(m);
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getMerges(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_merge_count, &merge_item);
}

emscripten::val JsWorkbook::getMergesInRange(uint32_t sheet, emscripten::val range) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.getMergesInRange");
  const fm_cell_rect_t query = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  size_t total = 0;
  fm_status_t rc = fm_sheet_get_merges_in_range(handle_, sheet, &query, nullptr, 0, &total);
  std::vector<fm_cell_rect_t> found(rc == 0 ? total : 0);
  if (rc == 0 && total > 0) {
    size_t written = 0;
    rc = fm_sheet_get_merges_in_range(handle_, sheet, &query, found.data(), total, &written);
    found.resize(rc == 0 ? written : 0);
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return cell_rects_to_val(found.data(), found.size());
}

// ---- Comments ----------------------------------------------------------

emscripten::val JsWorkbook::getComment(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t present = 0;
  fm_comment_info_t c = FM_STRUCT_INIT(fm_comment_info_t);
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_comment(handle_, sheet, row, col, &present, &c) : kBindingInvalidHandle;
  if (rc != 0 || present == 0) {
    return optional_or_null(rc, present, emscripten::val::null());
  }
  emscripten::val comment = emscripten::val::object();
  js_set_cstr_fields(comment, {{"author", c.author}, {"text", c.text}});
  return optional_or_null(rc, present, comment);
}

namespace {

fm_status_t comment_item(const fm_workbook_t* wb, uint32_t sheet, std::size_t index, emscripten::val* out) {
  fm_comment_info_t c = FM_STRUCT_INIT(fm_comment_info_t);
  const fm_status_t rc = fm_sheet_get_comment_at(wb, sheet, index, &c);
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

void JsWorkbook::setComment(uint32_t sheet, uint32_t row, uint32_t col, const std::string& author,
                            const std::string& text) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  const char* author_c = author.empty() ? nullptr : author.c_str();
  const char* text_c = text.empty() ? nullptr : text.c_str();
  fm_status_t rc = fm_sheet_set_comment(handle_, sheet, row, col, author_c, text_c);
  succeeded(rc);
}

void JsWorkbook::removeComment(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_comment(handle_, sheet, row, col) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getCommentVisible(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t present = 0;
  int32_t visible = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_comment_visible(handle_, sheet, row, col, &present, &visible)
                                            : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 ? emscripten::val(visible != 0) : emscripten::val::null());
}

void JsWorkbook::setCommentVisible(uint32_t sheet, uint32_t row, uint32_t col, bool visible) {
  succeeded(handle_ != nullptr ? fm_sheet_set_comment_visible(handle_, sheet, row, col, visible ? 1 : 0)
                               : kBindingInvalidHandle);
}

// ---- Protection records ------------------------------------------------

emscripten::val JsWorkbook::getWorkbookProtection() const {
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_workbook_protection(handle_, &p) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val protection = emscripten::val::object();
  protection.set("lockStructure", p.lock_structure != 0);
  protection.set("lockWindows", p.lock_windows != 0);
  protection.set("lockRevision", p.lock_revision != 0);
  protection.set("workbookSpinCount", p.workbook_spin_count);
  protection.set("revisionsSpinCount", p.revisions_spin_count);
  js_set_cstr_or_null_fields(protection, {{"workbookLegacyPasswordHash", p.workbook_legacy_password_hash},
                                          {"workbookAlgorithmName", p.workbook_algorithm_name},
                                          {"workbookHashValue", p.workbook_hash_value},
                                          {"workbookSaltValue", p.workbook_salt_value},
                                          {"revisionsLegacyPasswordHash", p.revisions_legacy_password_hash},
                                          {"revisionsAlgorithmName", p.revisions_algorithm_name},
                                          {"revisionsHashValue", p.revisions_hash_value},
                                          {"revisionsSaltValue", p.revisions_salt_value}});
  return protection;
}

void JsWorkbook::setWorkbookProtection(emscripten::val protection) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Workbook.setWorkbookProtection");
  fm_workbook_protection_t p = FM_STRUCT_INIT(fm_workbook_protection_t);
  p.lock_structure = reader.boolean(protection, "lockStructure", false, "protection.lockStructure") ? 1 : 0;
  p.lock_windows = reader.boolean(protection, "lockWindows", false, "protection.lockWindows") ? 1 : 0;
  p.lock_revision = reader.boolean(protection, "lockRevision", false, "protection.lockRevision") ? 1 : 0;
  p.workbook_spin_count = reader.u32(protection, "workbookSpinCount", 0U, "protection.workbookSpinCount");
  p.revisions_spin_count = reader.u32(protection, "revisionsSpinCount", 0U, "protection.revisionsSpinCount");
  std::string text[8];
  p.workbook_legacy_password_hash = reader.optional_string(protection, "workbookLegacyPasswordHash", text[0],
                                                           "protection.workbookLegacyPasswordHash");
  p.workbook_algorithm_name =
      reader.optional_string(protection, "workbookAlgorithmName", text[1], "protection.workbookAlgorithmName");
  p.workbook_hash_value =
      reader.optional_string(protection, "workbookHashValue", text[2], "protection.workbookHashValue");
  p.workbook_salt_value =
      reader.optional_string(protection, "workbookSaltValue", text[3], "protection.workbookSaltValue");
  p.revisions_legacy_password_hash = reader.optional_string(protection, "revisionsLegacyPasswordHash", text[4],
                                                            "protection.revisionsLegacyPasswordHash");
  p.revisions_algorithm_name =
      reader.optional_string(protection, "revisionsAlgorithmName", text[5], "protection.revisionsAlgorithmName");
  p.revisions_hash_value =
      reader.optional_string(protection, "revisionsHashValue", text[6], "protection.revisionsHashValue");
  p.revisions_salt_value =
      reader.optional_string(protection, "revisionsSaltValue", text[7], "protection.revisionsSaltValue");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_set_workbook_protection(handle_, &p));
}

namespace {

emscripten::val protected_range_to_val(const fm_protected_range_t& r) {
  emscripten::val o = emscripten::val::object();
  o.set("spinCount", r.spin_count);
  o.set("ranges", cell_rects_to_val(r.ranges, r.range_count));
  js_set_cstr(o, "name", r.name);
  js_set_cstr_or_null(o, "legacyPasswordHash", r.legacy_password_hash);
  js_set_cstr_or_null_fields(o, {{"algorithmName", r.algorithm_name},
                                 {"hashValue", r.hash_value},
                                 {"saltValue", r.salt_value},
                                 {"securityDescriptor", r.security_descriptor}});
  return o;
}

}  // namespace

double JsWorkbook::getProtectedRangeCount(uint32_t sheet) const {
  return js_count(handle_, sheet, &fm_sheet_get_protected_range_count);
}

emscripten::val JsWorkbook::getProtectedRange(uint32_t sheet, uint32_t protectedRangeIndex) const {
  fm_protected_range_t r = FM_STRUCT_INIT(fm_protected_range_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_protected_range_at(handle_, sheet, protectedRangeIndex, &r)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return protected_range_to_val(r);
}

namespace {

// Reads the protected range described by `range` into `*out`, with `text` owning its strings and
// `ranges` its rectangles. False after the reader has recorded its rejection.
bool pull_protected_range(const emscripten::val& range, const char* operation, std::string (&text)[6],
                          std::vector<fm_cell_rect_t>& ranges, fm_protected_range_t* out) {
  JsNarrowNumericReader reader(operation);
  out->spin_count = reader.u32(range, "spinCount", 0U, "range.spinCount");
  out->name = reader.optional_string(range, "name", text[0], "range.name");
  const emscripten::val ranges_value = reader.value(range, "ranges", "range.ranges");
  if (reader.require_array(ranges_value, "range.ranges")) {
    ranges = js_pull_range_array(ranges_value, "range.ranges", &reader);
  }
  out->ranges = ranges.data();
  out->range_count = ranges.size();
  out->legacy_password_hash = reader.optional_string(range, "legacyPasswordHash", text[1], "range.legacyPasswordHash");
  out->algorithm_name = reader.optional_string(range, "algorithmName", text[2], "range.algorithmName");
  out->hash_value = reader.optional_string(range, "hashValue", text[3], "range.hashValue");
  out->salt_value = reader.optional_string(range, "saltValue", text[4], "range.saltValue");
  out->security_descriptor = reader.optional_string(range, "securityDescriptor", text[5], "range.securityDescriptor");
  return reader.ok();
}

}  // namespace

emscripten::val JsWorkbook::addProtectedRange(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  std::string text[6];
  std::vector<fm_cell_rect_t> ranges;
  fm_protected_range_t r = FM_STRUCT_INIT(fm_protected_range_t);
  if (!pull_protected_range(range, "addProtectedRange", text, ranges, &r)) {
    return emscripten::val::undefined();
  }
  if (!succeeded(fm_sheet_add_protected_range(handle_, sheet, &r))) {
    return emscripten::val::undefined();
  }
  return emscripten::val(text[0]);
}

void JsWorkbook::updateProtectedRange(uint32_t sheet, const std::string& name, emscripten::val range) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateProtectedRange");
  std::string text[6];
  std::vector<fm_cell_rect_t> ranges;
  fm_protected_range_patch_t p = FM_STRUCT_INIT(fm_protected_range_patch_t);
  const emscripten::val spin_count = reader.value(range, "spinCount", "range.spinCount");
  if (reader.ok() && js_value_present(spin_count)) {
    p.spin_count_op = FM_PATCH_OP_SET;
    p.spin_count = reader.u32_value(spin_count, 0U, "range.spinCount");
  }
  js_patch_string(reader, range, "name", "range.name", false, text[0], &p.name_op, &p.name);
  const emscripten::val ranges_value = reader.value(range, "ranges", "range.ranges");
  if (reader.ok() && js_value_present(ranges_value) && reader.require_array(ranges_value, "range.ranges")) {
    ranges = js_pull_range_array(ranges_value, "range.ranges", &reader);
    p.ranges_op = FM_PATCH_OP_SET;
    p.ranges = ranges.data();
    p.range_count = ranges.size();
  }
  js_patch_string(reader, range, "legacyPasswordHash", "range.legacyPasswordHash", true, text[1],
                  &p.legacy_password_hash_op, &p.legacy_password_hash);
  js_patch_string(reader, range, "algorithmName", "range.algorithmName", true, text[2], &p.algorithm_name_op,
                  &p.algorithm_name);
  js_patch_string(reader, range, "hashValue", "range.hashValue", true, text[3], &p.hash_value_op, &p.hash_value);
  js_patch_string(reader, range, "saltValue", "range.saltValue", true, text[4], &p.salt_value_op, &p.salt_value);
  js_patch_string(reader, range, "securityDescriptor", "range.securityDescriptor", true, text[5],
                  &p.security_descriptor_op, &p.security_descriptor);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_protected_range(handle_, sheet, name.c_str(), &p));
}

void JsWorkbook::removeProtectedRange(uint32_t sheet, const std::string& name) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_protected_range(handle_, sheet, name.c_str()) : kBindingInvalidHandle);
}

// ---- Hyperlinks --------------------------------------------------------

namespace {

fm_status_t set_hyperlink_rect(fm_workbook_t* wb, uint32_t sheet, const fm_cell_rect_t& rect, const std::string& target,
                               const std::string& display, const std::string& tooltip, const std::string& location) {
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect = rect;
  hl.target = target.empty() ? nullptr : target.c_str();
  hl.location = location.empty() ? nullptr : location.c_str();
  hl.display = display.empty() ? nullptr : display.c_str();
  hl.tooltip = tooltip.empty() ? nullptr : tooltip.c_str();
  return fm_sheet_set_hyperlink(wb, sheet, rect.first_row, rect.first_col, &hl);
}

}  // namespace

void JsWorkbook::addHyperlink(uint32_t sheet, uint32_t row, uint32_t col, const std::string& target,
                              const std::string& display, const std::string& tooltip, const std::string& location) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  succeeded(set_hyperlink_rect(handle_, sheet, fm_cell_rect_t{row, col, row, col}, target, display, tooltip, location));
}

void JsWorkbook::addHyperlinkRange(uint32_t sheet, emscripten::val range, const std::string& target,
                                   const std::string& display, const std::string& tooltip,
                                   const std::string& location) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.setHyperlinkRange");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(set_hyperlink_rect(handle_, sheet, rect, target, display, tooltip, location));
}

void JsWorkbook::removeHyperlink(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_hyperlink(handle_, sheet, row, col) : kBindingInvalidHandle);
}

void JsWorkbook::removeHyperlinkAt(uint32_t sheet, uint32_t index) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_hyperlink_at(handle_, sheet, index) : kBindingInvalidHandle);
}

void JsWorkbook::clearHyperlinks(uint32_t sheet) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_hyperlinks(handle_, sheet) : kBindingInvalidHandle);
}

namespace {

emscripten::val hyperlink_to_val(const fm_hyperlink_t& h) {
  emscripten::val item = emscripten::val::object();
  item.set("rect", cell_rect_to_val(h.rect));
  js_set_cstr(item, "target", h.target);
  js_set_cstr_or_null_fields(item, {{"location", h.location}, {"display", h.display}, {"tooltip", h.tooltip}});
  return item;
}

fm_status_t hyperlink_item(const fm_workbook_t* wb, uint32_t sheet, std::size_t index, emscripten::val* out) {
  fm_hyperlink_t h = FM_STRUCT_INIT(fm_hyperlink_t);
  const fm_status_t rc = fm_sheet_get_hyperlink_at(wb, sheet, index, &h);
  if (rc != 0) {
    return rc;
  }
  *out = hyperlink_to_val(h);
  return 0;
}

}  // namespace

emscripten::val JsWorkbook::getHyperlinks(uint32_t sheet) const {
  return sheet_list(handle_, sheet, &fm_sheet_get_hyperlink_count, &hyperlink_item);
}

emscripten::val JsWorkbook::getHyperlink(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_hyperlink_t h = FM_STRUCT_INIT(fm_hyperlink_t);
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_hyperlink(handle_, sheet, row, col, &present, &h) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? hyperlink_to_val(h) : emscripten::val::null());
}

// ---- Data validations --------------------------------------------------

namespace {

emscripten::val validation_to_val(const fm_data_validation_t& v) {
  emscripten::val item = emscripten::val::object();
  js_set_cstr(item, "id", v.id);
  item.set("ranges", cell_rects_to_val(v.ranges, v.range_count));
  item.set("type", v.type);
  item.set("operator", v.op);
  item.set("errorStyle", v.error_style);
  // Getters report a tristate as FALSE or TRUE, never DEFAULT.
  item.set("allowBlank", v.allow_blank == FM_TRISTATE_TRUE);
  item.set("showInputMessage", v.show_input_message == FM_TRISTATE_TRUE);
  item.set("showErrorMessage", v.show_error_message == FM_TRISTATE_TRUE);
  item.set("showInCellDropdown", v.show_in_cell_dropdown == FM_TRISTATE_TRUE);
  js_set_cstr_or_null_fields(item, {{"formula1", v.formula1},
                                    {"formula2", v.formula2},
                                    {"errorTitle", v.error_title},
                                    {"errorMessage", v.error_message},
                                    {"promptTitle", v.prompt_title},
                                    {"promptMessage", v.prompt_message}});
  return item;
}

}  // namespace

double JsWorkbook::getValidationCount(uint32_t sheet) const {
  return js_count(handle_, sheet, &fm_sheet_get_validation_count);
}

emscripten::val JsWorkbook::getValidation(uint32_t sheet, uint32_t validationIndex) const {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_validation_at(handle_, sheet, validationIndex, &v)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return validation_to_val(v);
}

emscripten::val JsWorkbook::addValidation(uint32_t sheet, emscripten::val v) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  JsNarrowNumericReader reader("Sheet.addValidation");
  const emscripten::val ranges_value = reader.value(v, "ranges", "validation.ranges");
  std::vector<fm_cell_rect_t> ranges_buf;
  if (reader.require_array(ranges_value, "validation.ranges")) {
    ranges_buf = js_pull_range_array(ranges_value, "validation.ranges", &reader);
  }
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  const std::string formula1 = reader.string(v, "formula1", "validation.formula1");
  const std::string formula2 = reader.string(v, "formula2", "validation.formula2");
  const std::string error_title = reader.string(v, "errorTitle", "validation.errorTitle");
  const std::string error_message = reader.string(v, "errorMessage", "validation.errorMessage");
  const std::string prompt_title = reader.string(v, "promptTitle", "validation.promptTitle");
  const std::string prompt_message = reader.string(v, "promptMessage", "validation.promptMessage");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }

  fm_data_validation_t dv = FM_STRUCT_INIT(fm_data_validation_t);
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = ranges_buf.size();
  dv.type = reader.u8(v, "type", 0U, "validation.type");
  dv.op = reader.u8(v, "operator", 0U, "validation.operator");
  dv.error_style = reader.u8(v, "errorStyle", 0U, "validation.errorStyle");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  // An omitted flag takes Excel's dialog default.
  dv.allow_blank = js_tristate(reader, v, "allowBlank", "validation.allowBlank");
  dv.show_input_message = js_tristate(reader, v, "showInputMessage", "validation.showInputMessage");
  dv.show_error_message = js_tristate(reader, v, "showErrorMessage", "validation.showErrorMessage");
  dv.show_in_cell_dropdown = js_tristate(reader, v, "showInCellDropdown", "validation.showInCellDropdown");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  dv.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  dv.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  dv.error_title = error_title.empty() ? nullptr : error_title.c_str();
  dv.error_message = error_message.empty() ? nullptr : error_message.c_str();
  dv.prompt_title = prompt_title.empty() ? nullptr : prompt_title.c_str();
  dv.prompt_message = prompt_message.empty() ? nullptr : prompt_message.c_str();
  const char* out_id = nullptr;
  if (!succeeded(fm_sheet_add_validation(handle_, sheet, &dv, &out_id))) {
    return emscripten::val::undefined();
  }
  return emscripten::val(string_from_cstr(out_id));
}

void JsWorkbook::updateValidation(uint32_t sheet, const std::string& validationId, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateValidation");
  fm_data_validation_patch_t p = FM_STRUCT_INIT(fm_data_validation_patch_t);
  std::vector<fm_cell_rect_t> ranges;
  const emscripten::val ranges_value = reader.value(patch, "ranges", "validation.ranges");
  if (reader.ok() && js_patch_key(reader, ranges_value, "validation.ranges", false) == JsPatchKey::kSet &&
      reader.require_array(ranges_value, "validation.ranges")) {
    ranges = js_pull_range_array(ranges_value, "validation.ranges", &reader);
    p.ranges_op = FM_PATCH_OP_SET;
    p.ranges = ranges.data();
    p.range_count = ranges.size();
  }
  const auto u8 = [&](const char* key, const char* field, fm_patch_op_t* op, int32_t* value) {
    const emscripten::val v = reader.value(patch, key, field);
    if (js_patch_key(reader, v, field, false) == JsPatchKey::kSet) {
      *op = FM_PATCH_OP_SET;
      *value = static_cast<int32_t>(reader.u8_value(v, 0U, field));
    }
  };
  // A flag is a boolean in JS; the C patch carries it as a tristate.
  const auto flag = [&](const char* key, const char* field, fm_patch_op_t* op, int32_t* value) {
    const emscripten::val v = reader.value(patch, key, field);
    if (js_patch_key(reader, v, field, false) == JsPatchKey::kSet) {
      *op = FM_PATCH_OP_SET;
      *value = reader.boolean_value(v, false) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
    }
  };
  u8("type", "validation.type", &p.type_op, &p.type);
  u8("operator", "validation.operator", &p.op_op, &p.op);
  u8("errorStyle", "validation.errorStyle", &p.error_style_op, &p.error_style);
  flag("allowBlank", "validation.allowBlank", &p.allow_blank_op, &p.allow_blank);
  flag("showInputMessage", "validation.showInputMessage", &p.show_input_message_op, &p.show_input_message);
  flag("showErrorMessage", "validation.showErrorMessage", &p.show_error_message_op, &p.show_error_message);
  flag("showInCellDropdown", "validation.showInCellDropdown", &p.show_in_cell_dropdown_op, &p.show_in_cell_dropdown);
  std::string text[6];
  js_patch_string(reader, patch, "formula1", "validation.formula1", true, text[0], &p.formula1_op, &p.formula1);
  js_patch_string(reader, patch, "formula2", "validation.formula2", true, text[1], &p.formula2_op, &p.formula2);
  js_patch_string(reader, patch, "errorTitle", "validation.errorTitle", true, text[2], &p.error_title_op,
                  &p.error_title);
  js_patch_string(reader, patch, "errorMessage", "validation.errorMessage", true, text[3], &p.error_message_op,
                  &p.error_message);
  js_patch_string(reader, patch, "promptTitle", "validation.promptTitle", true, text[4], &p.prompt_title_op,
                  &p.prompt_title);
  js_patch_string(reader, patch, "promptMessage", "validation.promptMessage", true, text[5], &p.prompt_message_op,
                  &p.prompt_message);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_validation(handle_, sheet, validationId.c_str(), &p));
}

void JsWorkbook::removeValidationAt(uint32_t sheet, const std::string& validationId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_validation(handle_, sheet, validationId.c_str())
                               : kBindingInvalidHandle);
}

void JsWorkbook::clearValidations(uint32_t sheet) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_validations(handle_, sheet) : kBindingInvalidHandle);
}

// ---- Threaded comments and persons -------------------------------------

namespace {

/// Owns the strings behind a list of `fm_mention_t`.
struct MentionStore {
  std::vector<std::string> person_ids;
  std::vector<std::string> mention_ids;
  std::vector<fm_mention_t> items;
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

emscripten::val mention_to_val(const fm_mention_t& m) {
  emscripten::val o = emscripten::val::object();
  o.set("start", m.start);
  o.set("length", m.length);
  js_set_cstr_fields(o, {{"personId", m.person_id}, {"mentionId", m.mention_id}});
  return o;
}

emscripten::val threaded_comment_to_val(const fm_threaded_comment_t& c) {
  emscripten::val o = emscripten::val::object();
  o.set("row", c.row);
  o.set("col", c.col);
  o.set("done", c.done != 0);
  js_set_cstr_fields(o, {{"id", c.id}, {"personId", c.person_id}, {"text", c.text}});
  js_set_cstr_or_null(o, "parentId", c.parent_id);
  js_set_cstr_or_null(o, "created", c.created);
  emscripten::val mentions = emscripten::val::array();
  for (uint32_t i = 0; i < c.mentions_count; ++i) {
    mentions.set(i, mention_to_val(c.mentions[i]));
  }
  o.set("mentions", mentions);
  return o;
}

emscripten::val person_to_val(const fm_person_t& p) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr_fields(
      o, {{"id", p.id}, {"displayName", p.display_name}, {"userId", p.user_id}, {"providerId", p.provider_id}});
  return o;
}

fm_status_t threaded_comment_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
  const fm_status_t rc = fm_sheet_get_threaded_comment_at(handle, sheet, index, &c);
  if (rc == 0) {
    *out = threaded_comment_to_val(c);
  }
  return rc;
}

// Persons are workbook-scoped; the sheet argument is ignored.
fm_status_t person_count(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_person_count(handle, out);
}

fm_status_t person_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  const fm_status_t rc = fm_workbook_get_person_at(handle, index, &p);
  if (rc == 0) {
    *out = person_to_val(p);
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::getThreadedComments(uint32_t sheet) const {
  emscripten::val comments = emscripten::val::array();
  if (!succeeded(
          js_collect_list(handle_, sheet, &fm_sheet_get_threaded_comment_count, &threaded_comment_item, comments))) {
    return emscripten::val::undefined();
  }
  return comments;
}

emscripten::val JsWorkbook::addThreadedComment(uint32_t sheet, emscripten::val comment) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Sheet.addThreadedComment");
  if (!reader.require_object(comment, "comment")) {
    return emscripten::val::undefined();
  }
  // An omitted id and an omitted `created` are filled in by the engine.
  const std::string id = reader.string(comment, "id", "comment.id");
  const std::string person_id = reader.string(comment, "personId", "comment.personId");
  const std::string created = reader.string(comment, "created", "comment.created");
  const std::string text = reader.required_string(comment, "text", "text");
  const std::string parent_id = reader.string(comment, "parentId", "comment.parentId");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  MentionStore mentions;
  pull_mentions(reader.value(comment, "mentions", "comment.mentions"), mentions, reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
  c.id = id.c_str();
  c.row = reader.u32(comment, "row", 0U, "comment.row");
  c.col = reader.u32(comment, "col", 0U, "comment.col");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  c.person_id = person_id.c_str();
  c.created = created.c_str();
  c.text = text.c_str();
  c.parent_id = parent_id.c_str();
  c.done = reader.boolean(comment, "done", false, "done") ? 1 : 0;
  c.mentions = mentions.items.empty() ? nullptr : mentions.items.data();
  c.mentions_count = static_cast<uint32_t>(mentions.items.size());
  const char* out_id = nullptr;
  if (!succeeded(fm_sheet_add_threaded_comment(handle_, sheet, &c, &out_id))) {
    return emscripten::val::undefined();
  }
  return emscripten::val(string_from_cstr(out_id));
}

void JsWorkbook::editThreadedComment(uint32_t sheet, const std::string& id, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateThreadedComment");
  fm_threaded_comment_patch_t out = FM_STRUCT_INIT(fm_threaded_comment_patch_t);
  std::string text;
  js_patch_string(reader, patch, "text", "patch.text", false, text, &out.text_op, &out.text);
  MentionStore store;
  const emscripten::val mentions = reader.value(patch, "mentions", "patch.mentions");
  if (reader.ok() && js_value_present(mentions)) {
    out.mentions_op = FM_PATCH_OP_SET;
    pull_mentions(mentions, store, reader);
  }
  js_patch_flag(reader, patch, "done", "patch.done", &out.done_op, &out.done);
  if (!reader.ok()) {
    return;
  }
  out.mentions = store.items.empty() ? nullptr : store.items.data();
  out.mentions_count = static_cast<uint32_t>(store.items.size());
  succeeded(fm_sheet_update_threaded_comment(handle_, sheet, id.c_str(), &out));
}

void JsWorkbook::removeThreadedComment(uint32_t sheet, const std::string& id) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_threaded_comment(handle_, sheet, id.c_str()) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getPersons() const {
  emscripten::val persons = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, 0U, &person_count, &person_item, persons))) {
    return emscripten::val::undefined();
  }
  return persons;
}

emscripten::val JsWorkbook::addPerson(emscripten::val person) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  JsNarrowNumericReader reader("Workbook.addPerson");
  const std::string id = reader.required_string(person, "id", "person.id");
  const std::string display_name = reader.required_string(person, "displayName", "person.displayName");
  // An omitted `userId` / `providerId` is the empty string.
  const std::string user_id = reader.string(person, "userId", "person.userId");
  const std::string provider_id = reader.string(person, "providerId", "person.providerId");
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  p.id = id.c_str();
  p.display_name = display_name.c_str();
  p.user_id = user_id.c_str();
  p.provider_id = provider_id.c_str();
  const char* out_id = nullptr;
  if (!succeeded(fm_workbook_add_person(handle_, &p, &out_id))) {
    return emscripten::val::undefined();
  }
  return emscripten::val(string_from_cstr(out_id));
}

void JsWorkbook::removePerson(const std::string& id) {
  succeeded(handle_ != nullptr ? fm_workbook_remove_person(handle_, id.c_str()) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getPerson(const std::string& id) const {
  fm_person_t p = FM_STRUCT_INIT(fm_person_t);
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_person(handle_, id.c_str(), &present, &p) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? person_to_val(p) : emscripten::val::null());
}

emscripten::val JsWorkbook::getThreadedComment(uint32_t sheet, const std::string& id) const {
  fm_threaded_comment_t c = FM_STRUCT_INIT(fm_threaded_comment_t);
  int32_t present = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_threaded_comment(handle_, sheet, id.c_str(), &present, &c)
                                            : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? threaded_comment_to_val(c) : emscripten::val::null());
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
