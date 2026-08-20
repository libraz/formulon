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

// ---- AutoFilter --------------------------------------------------------

emscripten::val JsWorkbook::getSheetAutoFilterXml(uint32_t sheet) const {
  emscripten::val out = emscripten::val::object();
  if (handle_ == nullptr) {
    out.set("status", error_status(7000));
    out.set("xml", std::string());
    return out;
  }
  const char* xml = nullptr;
  fm_status_t rc = fm_sheet_get_auto_filter_xml(handle_, sheet, &xml);
  if (rc != 0) {
    out.set("status", error_status(rc));
    out.set("xml", std::string());
    return out;
  }
  out.set("status", ok_status());
  out.set("xml", xml != nullptr ? std::string(xml) : std::string());
  return out;
}

JsStatus JsWorkbook::setSheetAutoFilterXml(uint32_t sheet, const std::string& xml) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_set_auto_filter_xml(handle_, sheet, xml.c_str()));
}

// ---- Merges ------------------------------------------------------------

JsStatus JsWorkbook::addMerge(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_merge_range m;
  m.first_row = range["firstRow"].as<uint32_t>();
  m.last_row = range["lastRow"].as<uint32_t>();
  m.first_col = range["firstCol"].as<uint32_t>();
  m.last_col = range["lastCol"].as<uint32_t>();
  fm_status_t rc = fm_sheet_add_merge(handle_, sheet, m);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::removeMerge(uint32_t sheet, emscripten::val range) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_merge_range m;
  m.first_row = range["firstRow"].as<uint32_t>();
  m.last_row = range["lastRow"].as<uint32_t>();
  m.first_col = range["firstCol"].as<uint32_t>();
  m.last_col = range["lastCol"].as<uint32_t>();
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

emscripten::val JsWorkbook::getMerges(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(7000));
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_merge_count(handle_, sheet, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  uint32_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_merge_range m{};
    rc = fm_sheet_get_merge_at(handle_, sheet, i, &m);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
    }
    arr.set(emitted, merge_range_to_val(m));
    ++emitted;
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
  o.set("author", c.author != nullptr ? std::string(c.author) : std::string());
  o.set("text", c.text != nullptr ? std::string(c.text) : std::string());
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
  comment.set("author", c.author != nullptr ? std::string(c.author) : std::string());
  comment.set("text", c.text != nullptr ? std::string(c.text) : std::string());
  out.set("status", ok_status());
  out.set("comment", comment);
  return out;
}

emscripten::val JsWorkbook::getComments(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(7000));
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_comment_count(handle_, sheet, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  uint32_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_comment c{};
    rc = fm_sheet_get_comment_at_index(handle_, sheet, i, &c);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
    }
    emscripten::val o = emscripten::val::object();
    o.set("row", c.row);
    o.set("col", c.col);
    o.set("author", c.author != nullptr ? std::string(c.author) : std::string());
    o.set("text", c.text != nullptr ? std::string(c.text) : std::string());
    arr.set(emitted, o);
    ++emitted;
  }
  arr.set("status", ok_status());
  return arr;
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
  if (handle_ == nullptr) {
    return error_status(7000);
  }
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
  return status_from_rc(rc);
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

emscripten::val JsWorkbook::getHyperlinks(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(7000));
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_hyperlink_count(handle_, sheet, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  uint32_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_hyperlink h{};
    rc = fm_sheet_get_hyperlink_at(handle_, sheet, i, &h);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
    }
    emscripten::val item = emscripten::val::object();
    item.set("row", h.row);
    item.set("col", h.col);
    item.set("lastRow", h.last_row);
    item.set("lastCol", h.last_col);
    item.set("target", h.target != nullptr ? std::string(h.target) : std::string());
    item.set("location", h.location != nullptr ? std::string(h.location) : std::string());
    item.set("display", h.display != nullptr ? std::string(h.display) : std::string());
    item.set("tooltip", h.tooltip != nullptr ? std::string(h.tooltip) : std::string());
    arr.set(emitted, item);
    ++emitted;
  }
  arr.set("status", ok_status());
  return arr;
}

// ---- Data validations --------------------------------------------------

emscripten::val JsWorkbook::getValidations(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (handle_ == nullptr) {
    arr.set("status", error_status(7000));
    return arr;
  }
  uint32_t count = 0;
  fm_status_t rc = fm_sheet_get_validation_count(handle_, sheet, &count);
  if (rc != 0) {
    arr.set("status", status_from_rc(rc));
    return arr;
  }
  uint32_t emitted = 0;
  for (uint32_t i = 0; i < count; ++i) {
    fm_data_validation v{};
    rc = fm_sheet_get_validation_at(handle_, sheet, i, &v);
    if (rc != 0) {
      arr.set("status", status_from_rc(rc));
      return arr;
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
    item.set("formula1", v.formula1 != nullptr ? std::string(v.formula1) : std::string());
    item.set("formula2", v.formula2 != nullptr ? std::string(v.formula2) : std::string());
    item.set("errorTitle", v.error_title != nullptr ? std::string(v.error_title) : std::string());
    item.set("errorMessage", v.error_message != nullptr ? std::string(v.error_message) : std::string());
    item.set("promptTitle", v.prompt_title != nullptr ? std::string(v.prompt_title) : std::string());
    item.set("promptMessage", v.prompt_message != nullptr ? std::string(v.prompt_message) : std::string());
    arr.set(emitted, item);
    ++emitted;
  }
  arr.set("status", ok_status());
  return arr;
}

JsStatus JsWorkbook::addValidation(uint32_t sheet, emscripten::val v) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_add_validation` returns.
  std::vector<fm_merge_range> ranges_buf;
  if (v.hasOwnProperty("ranges")) {
    emscripten::val ranges_js = v["ranges"];
    if (ranges_js.isArray()) {
      const uint32_t n = ranges_js["length"].as<uint32_t>();
      ranges_buf.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        emscripten::val rng = ranges_js[i];
        fm_merge_range m{};
        m.first_row = rng["firstRow"].as<uint32_t>();
        m.last_row = rng["lastRow"].as<uint32_t>();
        m.first_col = rng["firstCol"].as<uint32_t>();
        m.last_col = rng["lastCol"].as<uint32_t>();
        ranges_buf.push_back(m);
      }
    }
  }
  const std::string formula1 = js_pull_string(v, "formula1");
  const std::string formula2 = js_pull_string(v, "formula2");
  const std::string error_title = js_pull_string(v, "errorTitle");
  const std::string error_message = js_pull_string(v, "errorMessage");
  const std::string prompt_title = js_pull_string(v, "promptTitle");
  const std::string prompt_message = js_pull_string(v, "promptMessage");

  fm_data_validation dv{};
  dv.ranges = ranges_buf.empty() ? nullptr : ranges_buf.data();
  dv.range_count = static_cast<uint32_t>(ranges_buf.size());
  dv.type = js_pull_u8(v, "type", 0U);
  dv.op = js_pull_u8(v, "op", 0U);
  dv.error_style = js_pull_u8(v, "errorStyle", 0U);
  dv.allow_blank = js_pull_bool(v, "allowBlank", false) ? 1 : 0;
  dv.show_input_message = js_pull_bool(v, "showInputMessage", false) ? 1 : 0;
  dv.show_error_message = js_pull_bool(v, "showErrorMessage", false) ? 1 : 0;
  dv.show_dropdown = js_pull_bool(v, "showDropDown", true) ? 1 : 0;
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

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
