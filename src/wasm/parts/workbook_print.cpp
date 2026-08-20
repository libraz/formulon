//
// JsWorkbook print-authoring surface: raw print-settings XML, the
// fit-to-page helper, print area / titles, manual page breaks, and the
// typed patch setters.
//
// The three surfaces map one-for-one onto the C ABI; the only translation
// this layer performs is the JS shape. Getters return `{ status, ... }` so
// an absent setting stays distinguishable from a rejected call, and the
// break enumerators return arrays rather than count + getter pairs,
// matching the rest of the JS surface.
//
// The patch setters read each field with the shared `js_pull_*` helpers and
// derive the `_engaged` flag from key presence, so `{ orientation: 1 }`
// changes the orientation and nothing else - the same semantics the C
// struct spells out with explicit flags.
//
// @size-budget: 8 KB

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

namespace {

using XmlGetter = fm_status_t (*)(const fm_workbook_t*, size_t, const char**);
using XmlSetter = fm_status_t (*)(fm_workbook_t*, size_t, const char*);

emscripten::val xml_result(const fm_workbook_t* handle, uint32_t sheet, XmlGetter getter) {
  emscripten::val out = emscripten::val::object();
  if (handle == nullptr) {
    out.set("status", error_status(7000));
    out.set("xml", std::string());
    return out;
  }
  const char* xml = nullptr;
  const fm_status_t rc = getter(handle, sheet, &xml);
  if (rc != 0) {
    out.set("status", error_status(rc));
    out.set("xml", std::string());
    return out;
  }
  out.set("status", ok_status());
  out.set("xml", xml != nullptr ? std::string(xml) : std::string());
  return out;
}

JsStatus xml_set(fm_workbook_t* handle, uint32_t sheet, const std::string& xml, XmlSetter setter) {
  if (handle == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(setter(handle, sheet, xml.c_str()));
}

emscripten::val breaks_array(const fm_workbook_t* handle, uint32_t sheet, bool rows) {
  emscripten::val out = emscripten::val::object();
  emscripten::val items = emscripten::val::array();
  if (handle == nullptr) {
    out.set("status", error_status(7000));
    out.set("breaks", items);
    return out;
  }
  const size_t count = rows ? fm_sheet_row_break_count(handle, sheet) : fm_sheet_col_break_count(handle, sheet);
  for (size_t i = 0; i < count; ++i) {
    fm_page_break_t brk{};
    const fm_status_t rc =
        rows ? fm_sheet_row_break_at(handle, sheet, i, &brk) : fm_sheet_col_break_at(handle, sheet, i, &brk);
    if (rc != 0) {
      out.set("status", error_status(rc));
      out.set("breaks", emscripten::val::array());
      return out;
    }
    emscripten::val entry = emscripten::val::object();
    entry.set("id", brk.id);
    entry.set("min", brk.min);
    entry.set("max", brk.max);
    entry.set("manual", brk.manual != 0);
    items.call<void>("push", entry);
  }
  out.set("status", ok_status());
  out.set("breaks", items);
  return out;
}

/// Reads an optional string field as the C tri-state: a missing key stays
/// `nullptr` (leave the section alone), a present one owns its bytes in
/// `storage` so the pointer outlives this call.
const char* pull_optional_string(const emscripten::val& object, const char* key, std::string& storage) {
  const emscripten::val field = object[key];
  if (field.isUndefined() || field.isNull()) {
    return nullptr;
  }
  storage = field.as<std::string>();
  return storage.c_str();
}

bool has_field(const emscripten::val& object, const char* key) {
  const emscripten::val field = object[key];
  return !field.isUndefined() && !field.isNull();
}

}  // namespace

// ---- Raw XML -----------------------------------------------------------

emscripten::val JsWorkbook::getSheetPageSetupXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_page_setup_xml);
}

JsStatus JsWorkbook::setSheetPageSetupXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_page_setup_xml);
}

emscripten::val JsWorkbook::getSheetPageMarginsXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_page_margins_xml);
}

JsStatus JsWorkbook::setSheetPageMarginsXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_page_margins_xml);
}

emscripten::val JsWorkbook::getSheetPrintOptionsXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_print_options_xml);
}

JsStatus JsWorkbook::setSheetPrintOptionsXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_print_options_xml);
}

emscripten::val JsWorkbook::getSheetHeaderFooterXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_header_footer_xml);
}

JsStatus JsWorkbook::setSheetHeaderFooterXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_header_footer_xml);
}

emscripten::val JsWorkbook::getSheetSheetPrXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_sheet_pr_xml);
}

JsStatus JsWorkbook::setSheetSheetPrXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_sheet_pr_xml);
}

JsStatus JsWorkbook::setSheetFitToPage(uint32_t sheet, bool enabled) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_set_fit_to_page(handle_, sheet, enabled ? 1 : 0));
}

// ---- Print area / titles -----------------------------------------------

emscripten::val JsWorkbook::getSheetPrintArea(uint32_t sheet) const {
  emscripten::val out = emscripten::val::object();
  if (handle_ == nullptr) {
    out.set("status", error_status(7000));
    out.set("ranges", std::string());
    return out;
  }
  const char* ranges = nullptr;
  const fm_status_t rc = fm_sheet_get_print_area(handle_, sheet, &ranges);
  if (rc != 0) {
    out.set("status", error_status(rc));
    out.set("ranges", std::string());
    return out;
  }
  out.set("status", ok_status());
  out.set("ranges", ranges != nullptr ? std::string(ranges) : std::string());
  return out;
}

JsStatus JsWorkbook::setSheetPrintArea(uint32_t sheet, const std::string& rangesA1) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_set_print_area(handle_, sheet, rangesA1.c_str()));
}

emscripten::val JsWorkbook::getSheetPrintTitles(uint32_t sheet) const {
  emscripten::val out = emscripten::val::object();
  if (handle_ == nullptr) {
    out.set("status", error_status(7000));
    out.set("repeatRows", std::string());
    out.set("repeatCols", std::string());
    return out;
  }
  const char* rows = nullptr;
  const char* cols = nullptr;
  const fm_status_t rc = fm_sheet_get_print_titles(handle_, sheet, &rows, &cols);
  if (rc != 0) {
    out.set("status", error_status(rc));
    out.set("repeatRows", std::string());
    out.set("repeatCols", std::string());
    return out;
  }
  out.set("status", ok_status());
  out.set("repeatRows", rows != nullptr ? std::string(rows) : std::string());
  out.set("repeatCols", cols != nullptr ? std::string(cols) : std::string());
  return out;
}

JsStatus JsWorkbook::setSheetPrintTitles(uint32_t sheet, const std::string& repeatRows, const std::string& repeatCols) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_set_print_titles(handle_, sheet, repeatRows.c_str(), repeatCols.c_str()));
}

// ---- Manual page breaks -------------------------------------------------

JsStatus JsWorkbook::addSheetRowBreak(uint32_t sheet, uint32_t row, bool manual) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_add_row_break(handle_, sheet, row, manual ? 1 : 0));
}

JsStatus JsWorkbook::addSheetColBreak(uint32_t sheet, uint32_t col, bool manual) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_add_col_break(handle_, sheet, col, manual ? 1 : 0));
}

JsStatus JsWorkbook::removeSheetRowBreak(uint32_t sheet, uint32_t row) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_remove_row_break(handle_, sheet, row));
}

JsStatus JsWorkbook::removeSheetColBreak(uint32_t sheet, uint32_t col) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_remove_col_break(handle_, sheet, col));
}

JsStatus JsWorkbook::clearSheetBreaks(uint32_t sheet) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_clear_breaks(handle_, sheet));
}

emscripten::val JsWorkbook::getSheetRowBreaks(uint32_t sheet) const {
  return breaks_array(handle_, sheet, /*rows=*/true);
}

emscripten::val JsWorkbook::getSheetColBreaks(uint32_t sheet) const {
  return breaks_array(handle_, sheet, /*rows=*/false);
}

// ---- Typed patch setters ------------------------------------------------

JsStatus JsWorkbook::setSheetPageSetup(uint32_t sheet, emscripten::val setup) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_page_setup_t out{};
  out.orientation_engaged = has_field(setup, "orientation") ? 1 : 0;
  out.orientation = js_pull_u32(setup, "orientation", 0U);
  out.paper_size_engaged = has_field(setup, "paperSize") ? 1 : 0;
  out.paper_size = js_pull_u32(setup, "paperSize", 0U);
  out.scale_engaged = has_field(setup, "scale") ? 1 : 0;
  out.scale = js_pull_u32(setup, "scale", 0U);
  out.fit_to_width_engaged = has_field(setup, "fitToWidth") ? 1 : 0;
  out.fit_to_width = js_pull_u32(setup, "fitToWidth", 0U);
  out.fit_to_height_engaged = has_field(setup, "fitToHeight") ? 1 : 0;
  out.fit_to_height = js_pull_u32(setup, "fitToHeight", 0U);
  out.fit_to_page_engaged = has_field(setup, "fitToPage") ? 1 : 0;
  out.fit_to_page = js_pull_bool(setup, "fitToPage", false) ? 1 : 0;
  return status_from_rc(fm_sheet_set_page_setup(handle_, sheet, &out));
}

JsStatus JsWorkbook::setSheetPageMargins(uint32_t sheet, emscripten::val margins) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_page_margins_t out{};
  out.left_engaged = has_field(margins, "left") ? 1 : 0;
  out.left = js_pull_double(margins, "left", 0.0);
  out.right_engaged = has_field(margins, "right") ? 1 : 0;
  out.right = js_pull_double(margins, "right", 0.0);
  out.top_engaged = has_field(margins, "top") ? 1 : 0;
  out.top = js_pull_double(margins, "top", 0.0);
  out.bottom_engaged = has_field(margins, "bottom") ? 1 : 0;
  out.bottom = js_pull_double(margins, "bottom", 0.0);
  out.header_engaged = has_field(margins, "header") ? 1 : 0;
  out.header = js_pull_double(margins, "header", 0.0);
  out.footer_engaged = has_field(margins, "footer") ? 1 : 0;
  out.footer = js_pull_double(margins, "footer", 0.0);
  return status_from_rc(fm_sheet_set_page_margins(handle_, sheet, &out));
}

JsStatus JsWorkbook::setSheetPrintOptions(uint32_t sheet, emscripten::val options) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_print_options_t out{};
  out.grid_lines_engaged = has_field(options, "gridLines") ? 1 : 0;
  out.grid_lines = js_pull_bool(options, "gridLines", false) ? 1 : 0;
  out.headings_engaged = has_field(options, "headings") ? 1 : 0;
  out.headings = js_pull_bool(options, "headings", false) ? 1 : 0;
  out.horizontal_centered_engaged = has_field(options, "horizontalCentered") ? 1 : 0;
  out.horizontal_centered = js_pull_bool(options, "horizontalCentered", false) ? 1 : 0;
  out.vertical_centered_engaged = has_field(options, "verticalCentered") ? 1 : 0;
  out.vertical_centered = js_pull_bool(options, "verticalCentered", false) ? 1 : 0;
  return status_from_rc(fm_sheet_set_print_options(handle_, sheet, &out));
}

JsStatus JsWorkbook::setSheetHeaderFooter(uint32_t sheet, emscripten::val headerFooter) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  // The six section strings must outlive the call, so their storage is
  // declared here rather than inside the helper.
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
  fm_header_footer_t out{};
  out.odd_header = pull_optional_string(headerFooter, "oddHeader", odd_header);
  out.odd_footer = pull_optional_string(headerFooter, "oddFooter", odd_footer);
  out.even_header = pull_optional_string(headerFooter, "evenHeader", even_header);
  out.even_footer = pull_optional_string(headerFooter, "evenFooter", even_footer);
  out.first_header = pull_optional_string(headerFooter, "firstHeader", first_header);
  out.first_footer = pull_optional_string(headerFooter, "firstFooter", first_footer);
  out.different_odd_even_engaged = has_field(headerFooter, "differentOddEven") ? 1 : 0;
  out.different_odd_even = js_pull_bool(headerFooter, "differentOddEven", false) ? 1 : 0;
  out.different_first_engaged = has_field(headerFooter, "differentFirst") ? 1 : 0;
  out.different_first = js_pull_bool(headerFooter, "differentFirst", false) ? 1 : 0;
  out.scale_with_doc_engaged = has_field(headerFooter, "scaleWithDoc") ? 1 : 0;
  out.scale_with_doc = js_pull_bool(headerFooter, "scaleWithDoc", false) ? 1 : 0;
  out.align_with_margins_engaged = has_field(headerFooter, "alignWithMargins") ? 1 : 0;
  out.align_with_margins = js_pull_bool(headerFooter, "alignWithMargins", false) ? 1 : 0;
  return status_from_rc(fm_sheet_set_header_footer(handle_, sheet, &out));
}

// Both typed getters below emit their whole declared payload on every exit
// path; only `status.ok` and the values differ. See the same note in
// `parts/workbook_styles.cpp`.

emscripten::val JsWorkbook::getSheetPageSetup(uint32_t sheet) const {
  emscripten::val out = emscripten::val::object();
  fm_page_setup_t setup{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_page_setup(handle_, sheet, &setup) : 7000;
  if (rc != 0) {
    setup = fm_page_setup_t{};
  }
  out.set("status", status_from_rc(rc));
  out.set("orientation", setup.orientation);
  out.set("paperSize", setup.paper_size);
  out.set("scale", setup.scale);
  out.set("fitToWidth", setup.fit_to_width);
  out.set("fitToHeight", setup.fit_to_height);
  out.set("fitToPage", setup.fit_to_page != 0);
  // Presence flags: the value fields always carry the effective setting,
  // so these are the only way to tell an explicit value from a default.
  out.set("orientationStated", setup.orientation_engaged != 0);
  out.set("paperSizeStated", setup.paper_size_engaged != 0);
  out.set("scaleStated", setup.scale_engaged != 0);
  out.set("fitToWidthStated", setup.fit_to_width_engaged != 0);
  out.set("fitToHeightStated", setup.fit_to_height_engaged != 0);
  out.set("fitToPageStated", setup.fit_to_page_engaged != 0);
  return out;
}

emscripten::val JsWorkbook::getSheetPageMargins(uint32_t sheet) const {
  emscripten::val out = emscripten::val::object();
  fm_page_margins_t margins{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_page_margins(handle_, sheet, &margins) : 7000;
  if (rc != 0) {
    margins = fm_page_margins_t{};
  }
  out.set("status", status_from_rc(rc));
  out.set("left", margins.left);
  out.set("right", margins.right);
  out.set("top", margins.top);
  out.set("bottom", margins.bottom);
  out.set("header", margins.header);
  out.set("footer", margins.footer);
  out.set("leftStated", margins.left_engaged != 0);
  out.set("rightStated", margins.right_engaged != 0);
  out.set("topStated", margins.top_engaged != 0);
  out.set("bottomStated", margins.bottom_engaged != 0);
  out.set("headerStated", margins.header_engaged != 0);
  out.set("footerStated", margins.footer_engaged != 0);
  return out;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
