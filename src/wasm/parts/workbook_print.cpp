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
// The patch setters snapshot each field through JsNarrowNumericReader before
// deriving its `_engaged` flag. This keeps getter/proxy failures and lossy
// primitive conversions inside the binding boundary, while `{ orientation:
// 1 }` still changes the orientation and nothing else.
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

bool value_present(const emscripten::val& value) {
  return !value.isUndefined() && !value.isNull();
}

emscripten::val xml_result(const fm_workbook_t* handle, uint32_t sheet, XmlGetter getter, const char* key = "xml") {
  emscripten::val out = emscripten::val::object();
  if (handle == nullptr) {
    out.set("status", error_status(7000));
    out.set(key, std::string());
    return out;
  }
  const char* xml = nullptr;
  const fm_status_t rc = getter(handle, sheet, &xml);
  if (rc != 0) {
    out.set("status", error_status(rc));
    out.set(key, std::string());
    return out;
  }
  out.set("status", ok_status());
  js_set_cstr(out, key, xml);
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
  return xml_result(handle_, sheet, &fm_sheet_get_print_area, "ranges");
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
  js_set_cstr(out, "repeatRows", rows);
  js_set_cstr(out, "repeatCols", cols);
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
  JsNarrowNumericReader reader("setSheetPageSetup");
  fm_page_setup_t out{};
  const emscripten::val orientation = reader.value(setup, "orientation", "pageSetup.orientation");
  out.orientation_engaged = value_present(orientation) ? 1 : 0;
  out.orientation = reader.u32_value(orientation, 0U, "pageSetup.orientation");
  const emscripten::val paper_size = reader.value(setup, "paperSize", "pageSetup.paperSize");
  out.paper_size_engaged = value_present(paper_size) ? 1 : 0;
  out.paper_size = reader.u32_value(paper_size, 0U, "pageSetup.paperSize");
  const emscripten::val scale = reader.value(setup, "scale", "pageSetup.scale");
  out.scale_engaged = value_present(scale) ? 1 : 0;
  out.scale = reader.u32_value(scale, 0U, "pageSetup.scale");
  const emscripten::val fit_to_width = reader.value(setup, "fitToWidth", "pageSetup.fitToWidth");
  out.fit_to_width_engaged = value_present(fit_to_width) ? 1 : 0;
  out.fit_to_width = reader.u32_value(fit_to_width, 0U, "pageSetup.fitToWidth");
  const emscripten::val fit_to_height = reader.value(setup, "fitToHeight", "pageSetup.fitToHeight");
  out.fit_to_height_engaged = value_present(fit_to_height) ? 1 : 0;
  out.fit_to_height = reader.u32_value(fit_to_height, 0U, "pageSetup.fitToHeight");
  const emscripten::val fit_to_page = reader.value(setup, "fitToPage", "pageSetup.fitToPage");
  out.fit_to_page_engaged = value_present(fit_to_page) ? 1 : 0;
  out.fit_to_page = reader.boolean_value(fit_to_page, false, "pageSetup.fitToPage") ? 1 : 0;
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_set_page_setup(handle_, sheet, &out));
}

JsStatus JsWorkbook::setSheetPageMargins(uint32_t sheet, emscripten::val margins) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  JsNarrowNumericReader reader("setSheetPageMargins");
  fm_page_margins_t out{};
  const emscripten::val left = reader.value(margins, "left", "pageMargins.left");
  out.left_engaged = value_present(left) ? 1 : 0;
  out.left = reader.number_value(left, 0.0, "pageMargins.left");
  const emscripten::val right = reader.value(margins, "right", "pageMargins.right");
  out.right_engaged = value_present(right) ? 1 : 0;
  out.right = reader.number_value(right, 0.0, "pageMargins.right");
  const emscripten::val top = reader.value(margins, "top", "pageMargins.top");
  out.top_engaged = value_present(top) ? 1 : 0;
  out.top = reader.number_value(top, 0.0, "pageMargins.top");
  const emscripten::val bottom = reader.value(margins, "bottom", "pageMargins.bottom");
  out.bottom_engaged = value_present(bottom) ? 1 : 0;
  out.bottom = reader.number_value(bottom, 0.0, "pageMargins.bottom");
  const emscripten::val header = reader.value(margins, "header", "pageMargins.header");
  out.header_engaged = value_present(header) ? 1 : 0;
  out.header = reader.number_value(header, 0.0, "pageMargins.header");
  const emscripten::val footer = reader.value(margins, "footer", "pageMargins.footer");
  out.footer_engaged = value_present(footer) ? 1 : 0;
  out.footer = reader.number_value(footer, 0.0, "pageMargins.footer");
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
  return status_from_rc(fm_sheet_set_page_margins(handle_, sheet, &out));
}

JsStatus JsWorkbook::setSheetPrintOptions(uint32_t sheet, emscripten::val options) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  JsNarrowNumericReader reader("setSheetPrintOptions");
  fm_print_options_t out{};
  const emscripten::val grid_lines = reader.value(options, "gridLines", "printOptions.gridLines");
  out.grid_lines_engaged = value_present(grid_lines) ? 1 : 0;
  out.grid_lines = reader.boolean_value(grid_lines, false, "printOptions.gridLines") ? 1 : 0;
  const emscripten::val headings = reader.value(options, "headings", "printOptions.headings");
  out.headings_engaged = value_present(headings) ? 1 : 0;
  out.headings = reader.boolean_value(headings, false, "printOptions.headings") ? 1 : 0;
  const emscripten::val horizontal_centered =
      reader.value(options, "horizontalCentered", "printOptions.horizontalCentered");
  out.horizontal_centered_engaged = value_present(horizontal_centered) ? 1 : 0;
  out.horizontal_centered = reader.boolean_value(horizontal_centered, false, "printOptions.horizontalCentered") ? 1 : 0;
  const emscripten::val vertical_centered = reader.value(options, "verticalCentered", "printOptions.verticalCentered");
  out.vertical_centered_engaged = value_present(vertical_centered) ? 1 : 0;
  out.vertical_centered = reader.boolean_value(vertical_centered, false, "printOptions.verticalCentered") ? 1 : 0;
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
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
  JsNarrowNumericReader reader("setSheetHeaderFooter");
  fm_header_footer_t out{};
  const emscripten::val odd_header_value = reader.value(headerFooter, "oddHeader", "headerFooter.oddHeader");
  out.odd_header = reader.optional_string_value(odd_header_value, odd_header, "headerFooter.oddHeader");
  const emscripten::val odd_footer_value = reader.value(headerFooter, "oddFooter", "headerFooter.oddFooter");
  out.odd_footer = reader.optional_string_value(odd_footer_value, odd_footer, "headerFooter.oddFooter");
  const emscripten::val even_header_value = reader.value(headerFooter, "evenHeader", "headerFooter.evenHeader");
  out.even_header = reader.optional_string_value(even_header_value, even_header, "headerFooter.evenHeader");
  const emscripten::val even_footer_value = reader.value(headerFooter, "evenFooter", "headerFooter.evenFooter");
  out.even_footer = reader.optional_string_value(even_footer_value, even_footer, "headerFooter.evenFooter");
  const emscripten::val first_header_value = reader.value(headerFooter, "firstHeader", "headerFooter.firstHeader");
  out.first_header = reader.optional_string_value(first_header_value, first_header, "headerFooter.firstHeader");
  const emscripten::val first_footer_value = reader.value(headerFooter, "firstFooter", "headerFooter.firstFooter");
  out.first_footer = reader.optional_string_value(first_footer_value, first_footer, "headerFooter.firstFooter");
  const emscripten::val different_odd_even =
      reader.value(headerFooter, "differentOddEven", "headerFooter.differentOddEven");
  out.different_odd_even_engaged = value_present(different_odd_even) ? 1 : 0;
  out.different_odd_even = reader.boolean_value(different_odd_even, false, "headerFooter.differentOddEven") ? 1 : 0;
  const emscripten::val different_first = reader.value(headerFooter, "differentFirst", "headerFooter.differentFirst");
  out.different_first_engaged = value_present(different_first) ? 1 : 0;
  out.different_first = reader.boolean_value(different_first, false, "headerFooter.differentFirst") ? 1 : 0;
  const emscripten::val scale_with_doc = reader.value(headerFooter, "scaleWithDoc", "headerFooter.scaleWithDoc");
  out.scale_with_doc_engaged = value_present(scale_with_doc) ? 1 : 0;
  out.scale_with_doc = reader.boolean_value(scale_with_doc, false, "headerFooter.scaleWithDoc") ? 1 : 0;
  const emscripten::val align_with_margins =
      reader.value(headerFooter, "alignWithMargins", "headerFooter.alignWithMargins");
  out.align_with_margins_engaged = value_present(align_with_margins) ? 1 : 0;
  out.align_with_margins = reader.boolean_value(align_with_margins, false, "headerFooter.alignWithMargins") ? 1 : 0;
  if (!reader.ok()) {
    return binding_error_status(kInvalidArgument, reader.message().c_str());
  }
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
