//
// JsWorkbook print-authoring surface: raw print-settings XML, the
// fit-to-page helper, print area / titles, manual page breaks, and the
// typed patch setters.
//
// The three surfaces map one-for-one onto the C ABI; the only translation
// this layer performs is the JS shape. A rejected call is thrown by the
// pre-js wrapper, so an absent setting stays distinguishable from a failure,
// and the break enumerators return arrays rather than count + getter pairs,
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

emscripten::val xml_result(const fm_workbook_t* handle, uint32_t sheet, XmlGetter getter) {
  const char* xml = nullptr;
  if (!succeeded(handle != nullptr ? getter(handle, sheet, &xml) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return cstr_or_null(xml);
}

void xml_set(fm_workbook_t* handle, uint32_t sheet, const std::string& xml, XmlSetter setter) {
  succeeded(handle != nullptr ? setter(handle, sheet, xml.c_str()) : kBindingInvalidHandle);
}

emscripten::val breaks_array(const fm_workbook_t* handle, uint32_t sheet, bool rows) {
  if (handle == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  size_t count = 0;
  const fm_status_t count_rc = rows ? fm_sheet_get_row_break_count(handle, sheet, &count)
                                    : fm_sheet_get_column_break_count(handle, sheet, &count);
  if (!succeeded(count_rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val items = emscripten::val::array();
  for (size_t i = 0; i < count; ++i) {
    uint32_t position = 0;
    uint32_t first = 0;
    uint32_t last = 0;
    int32_t manual = 0;
    fm_status_t rc = 0;
    if (rows) {
      fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
      rc = fm_sheet_get_row_break_at(handle, sheet, i, &brk);
      position = brk.row;
      first = brk.first;
      last = brk.last;
      manual = brk.manual;
    } else {
      fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
      rc = fm_sheet_get_column_break_at(handle, sheet, i, &brk);
      position = brk.col;
      first = brk.first;
      last = brk.last;
      manual = brk.manual;
    }
    if (!succeeded(rc)) {
      return emscripten::val::undefined();
    }
    emscripten::val entry = emscripten::val::object();
    entry.set(rows ? "row" : "col", position);
    entry.set("first", first);
    entry.set("last", last);
    entry.set("manual", manual != 0);
    items.call<void>("push", entry);
  }
  return items;
}

}  // namespace

// ---- Raw XML -----------------------------------------------------------

emscripten::val JsWorkbook::getSheetPageSetupXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_page_setup_xml);
}

void JsWorkbook::setSheetPageSetupXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_page_setup_xml);
}

emscripten::val JsWorkbook::getSheetPageMarginsXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_page_margins_xml);
}

void JsWorkbook::setSheetPageMarginsXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_page_margins_xml);
}

emscripten::val JsWorkbook::getSheetPrintOptionsXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_print_options_xml);
}

void JsWorkbook::setSheetPrintOptionsXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_print_options_xml);
}

emscripten::val JsWorkbook::getSheetHeaderFooterXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_header_footer_xml);
}

void JsWorkbook::setSheetHeaderFooterXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_header_footer_xml);
}

emscripten::val JsWorkbook::getSheetSheetPrXml(uint32_t sheet) const {
  return xml_result(handle_, sheet, &fm_sheet_get_sheet_pr_xml);
}

void JsWorkbook::setSheetSheetPrXml(uint32_t sheet, const std::string& xml) {
  return xml_set(handle_, sheet, xml, &fm_sheet_set_sheet_pr_xml);
}

// `null` when the sheet has no `<tabColor>` (the C ABI reports kind None).
emscripten::val JsWorkbook::getSheetTabColor(uint32_t sheet) const {
  fm_color_spec_t color{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_tab_color(handle_, sheet, &color) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_color_spec_or_null(color);
}

// `null` clears the tab colour; an object is read as an `fm_color_spec_t`.
void JsWorkbook::setSheetTabColor(uint32_t sheet, emscripten::val color) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (color.isNull()) {
    succeeded(fm_sheet_set_tab_color(handle_, sheet, nullptr));
    return;
  }
  JsNarrowNumericReader reader("setTabColor");
  const fm_color_spec_t spec = js_pull_color_spec_value(color, "color", &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_tab_color(handle_, sheet, &spec));
}

// ---- Print area / titles -----------------------------------------------

emscripten::val JsWorkbook::getSheetPrintArea(uint32_t sheet) const {
  const fm_cell_rect_t* found = nullptr;
  size_t count = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_print_area(handle_, sheet, &found, &count)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return cell_rects_to_val(found, count);
}

void JsWorkbook::setSheetPrintArea(uint32_t sheet, emscripten::val ranges) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setPrintArea");
  const std::vector<fm_cell_rect_t> rects = js_pull_range_array(ranges, "ranges", &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_print_area(handle_, sheet, rects.data(), rects.size()));
}

emscripten::val JsWorkbook::getSheetPrintTitles(uint32_t sheet) const {
  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_print_titles(handle_, sheet, &titles) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return print_titles_to_val(titles);
}

void JsWorkbook::setSheetPrintTitles(uint32_t sheet, emscripten::val titles) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (!js_value_present(titles)) {
    succeeded(fm_sheet_set_print_titles(handle_, sheet, nullptr));
    return;
  }
  JsNarrowNumericReader reader("Sheet.setPrintTitles");
  if (!reader.require_object(titles, "titles")) {
    return;
  }
  fm_print_titles_t in = FM_STRUCT_INIT(fm_print_titles_t);
  const emscripten::val rows = reader.value(titles, "rows", "rows");
  if (reader.ok() && js_value_present(rows) && reader.require_object(rows, "rows")) {
    in.has_rows = 1;
    in.first_row = reader.required_u32(rows, "first", "rows.first");
    in.last_row = reader.required_u32(rows, "last", "rows.last");
  }
  const emscripten::val columns = reader.value(titles, "columns", "columns");
  if (reader.ok() && js_value_present(columns) && reader.require_object(columns, "columns")) {
    in.has_columns = 1;
    in.first_col = reader.required_u32(columns, "first", "columns.first");
    in.last_col = reader.required_u32(columns, "last", "columns.last");
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_print_titles(handle_, sheet, &in));
}

// ---- Manual page breaks -------------------------------------------------

// `null` removes the break; a record `{ first?, last?, manual? }` upserts it. Both span ends omitted (or 0)
// cover the whole sheet, and `manual` defaults to true.
void JsWorkbook::setSheetRowBreak(uint32_t sheet, uint32_t row, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (input.isNull()) {
    succeeded(fm_sheet_set_row_break(handle_, sheet, row, nullptr));
    return;
  }
  JsNarrowNumericReader reader("setRowBreak");
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.first = reader.u32(input, "first", 0U, "break.first");
  brk.last = reader.u32(input, "last", 0U, "break.last");
  brk.manual = reader.boolean(input, "manual", true, "break.manual") ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_row_break(handle_, sheet, row, &brk));
}

void JsWorkbook::setSheetColBreak(uint32_t sheet, uint32_t col, emscripten::val input) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (input.isNull()) {
    succeeded(fm_sheet_set_column_break(handle_, sheet, col, nullptr));
    return;
  }
  JsNarrowNumericReader reader("setColumnBreak");
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  brk.first = reader.u32(input, "first", 0U, "break.first");
  brk.last = reader.u32(input, "last", 0U, "break.last");
  brk.manual = reader.boolean(input, "manual", true, "break.manual") ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_column_break(handle_, sheet, col, &brk));
}

void JsWorkbook::clearSheetBreaks(uint32_t sheet) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_breaks(handle_, sheet) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getSheetRowBreaks(uint32_t sheet) const {
  return breaks_array(handle_, sheet, /*rows=*/true);
}

emscripten::val JsWorkbook::getSheetColBreaks(uint32_t sheet) const {
  return breaks_array(handle_, sheet, /*rows=*/false);
}

// ---- Typed patch setters ------------------------------------------------

void JsWorkbook::setSheetPageSetup(uint32_t sheet, emscripten::val setup) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updatePageSetup");
  fm_page_setup_t out = FM_STRUCT_INIT(fm_page_setup_t);
  const emscripten::val orientation = reader.value(setup, "orientation", "pageSetup.orientation");
  out.orientation_engaged = js_value_present(orientation) ? 1 : 0;
  out.orientation = reader.u32_value(orientation, 0U, "pageSetup.orientation");
  const emscripten::val paper_size = reader.value(setup, "paperSize", "pageSetup.paperSize");
  out.paper_size_engaged = js_value_present(paper_size) ? 1 : 0;
  out.paper_size = reader.u32_value(paper_size, 0U, "pageSetup.paperSize");
  const emscripten::val scale = reader.value(setup, "scalePct", "pageSetup.scalePct");
  out.scale_pct_engaged = js_value_present(scale) ? 1 : 0;
  out.scale_pct = reader.u32_value(scale, 0U, "pageSetup.scalePct");
  const emscripten::val fit_to_width = reader.value(setup, "fitToWidth", "pageSetup.fitToWidth");
  out.fit_to_width_engaged = js_value_present(fit_to_width) ? 1 : 0;
  out.fit_to_width = reader.u32_value(fit_to_width, 0U, "pageSetup.fitToWidth");
  const emscripten::val fit_to_height = reader.value(setup, "fitToHeight", "pageSetup.fitToHeight");
  out.fit_to_height_engaged = js_value_present(fit_to_height) ? 1 : 0;
  out.fit_to_height = reader.u32_value(fit_to_height, 0U, "pageSetup.fitToHeight");
  const emscripten::val fit_to_page = reader.value(setup, "fitToPage", "pageSetup.fitToPage");
  out.fit_to_page_engaged = js_value_present(fit_to_page) ? 1 : 0;
  out.fit_to_page = reader.boolean_value(fit_to_page, false) ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_page_setup(handle_, sheet, &out));
}

void JsWorkbook::setSheetPageMargins(uint32_t sheet, emscripten::val margins) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updatePageMargins");
  fm_page_margins_t out = FM_STRUCT_INIT(fm_page_margins_t);
  const emscripten::val left = reader.value(margins, "leftIn", "pageMargins.leftIn");
  out.left_in_engaged = js_value_present(left) ? 1 : 0;
  out.left_in = reader.number_value(left, 0.0, "pageMargins.leftIn");
  const emscripten::val right = reader.value(margins, "rightIn", "pageMargins.rightIn");
  out.right_in_engaged = js_value_present(right) ? 1 : 0;
  out.right_in = reader.number_value(right, 0.0, "pageMargins.rightIn");
  const emscripten::val top = reader.value(margins, "topIn", "pageMargins.topIn");
  out.top_in_engaged = js_value_present(top) ? 1 : 0;
  out.top_in = reader.number_value(top, 0.0, "pageMargins.topIn");
  const emscripten::val bottom = reader.value(margins, "bottomIn", "pageMargins.bottomIn");
  out.bottom_in_engaged = js_value_present(bottom) ? 1 : 0;
  out.bottom_in = reader.number_value(bottom, 0.0, "pageMargins.bottomIn");
  const emscripten::val header = reader.value(margins, "headerIn", "pageMargins.headerIn");
  out.header_in_engaged = js_value_present(header) ? 1 : 0;
  out.header_in = reader.number_value(header, 0.0, "pageMargins.headerIn");
  const emscripten::val footer = reader.value(margins, "footerIn", "pageMargins.footerIn");
  out.footer_in_engaged = js_value_present(footer) ? 1 : 0;
  out.footer_in = reader.number_value(footer, 0.0, "pageMargins.footerIn");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_page_margins(handle_, sheet, &out));
}

void JsWorkbook::setSheetPrintOptions(uint32_t sheet, emscripten::val options) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updatePrintOptions");
  fm_print_options_t out = FM_STRUCT_INIT(fm_print_options_t);
  const emscripten::val grid_lines = reader.value(options, "gridLines", "printOptions.gridLines");
  out.grid_lines_engaged = js_value_present(grid_lines) ? 1 : 0;
  out.grid_lines = reader.boolean_value(grid_lines, false) ? 1 : 0;
  const emscripten::val headings = reader.value(options, "headings", "printOptions.headings");
  out.headings_engaged = js_value_present(headings) ? 1 : 0;
  out.headings = reader.boolean_value(headings, false) ? 1 : 0;
  const emscripten::val horizontal_centered =
      reader.value(options, "horizontalCentered", "printOptions.horizontalCentered");
  out.horizontal_centered_engaged = js_value_present(horizontal_centered) ? 1 : 0;
  out.horizontal_centered = reader.boolean_value(horizontal_centered, false) ? 1 : 0;
  const emscripten::val vertical_centered = reader.value(options, "verticalCentered", "printOptions.verticalCentered");
  out.vertical_centered_engaged = js_value_present(vertical_centered) ? 1 : 0;
  out.vertical_centered = reader.boolean_value(vertical_centered, false) ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_print_options(handle_, sheet, &out));
}

void JsWorkbook::setSheetHeaderFooter(uint32_t sheet, emscripten::val headerFooter) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // The six section strings must outlive the call, so their storage is
  // declared here rather than inside the helper.
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
  JsNarrowNumericReader reader("updateHeaderFooter");
  fm_header_footer_t out = FM_STRUCT_INIT(fm_header_footer_t);
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
  out.different_odd_even_engaged = js_value_present(different_odd_even) ? 1 : 0;
  out.different_odd_even = reader.boolean_value(different_odd_even, false) ? 1 : 0;
  const emscripten::val different_first = reader.value(headerFooter, "differentFirst", "headerFooter.differentFirst");
  out.different_first_engaged = js_value_present(different_first) ? 1 : 0;
  out.different_first = reader.boolean_value(different_first, false) ? 1 : 0;
  const emscripten::val scale_with_doc = reader.value(headerFooter, "scaleWithDoc", "headerFooter.scaleWithDoc");
  out.scale_with_doc_engaged = js_value_present(scale_with_doc) ? 1 : 0;
  out.scale_with_doc = reader.boolean_value(scale_with_doc, false) ? 1 : 0;
  const emscripten::val align_with_margins =
      reader.value(headerFooter, "alignWithMargins", "headerFooter.alignWithMargins");
  out.align_with_margins_engaged = js_value_present(align_with_margins) ? 1 : 0;
  out.align_with_margins = reader.boolean_value(align_with_margins, false) ? 1 : 0;
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_header_footer(handle_, sheet, &out));
}

emscripten::val JsWorkbook::getSheetPageSetup(uint32_t sheet) const {
  fm_page_setup_t setup = FM_STRUCT_INIT(fm_page_setup_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_page_setup(handle_, sheet, &setup) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val out = emscripten::val::object();
  js_set_present_or_null(out, "orientation", setup.orientation_engaged, setup.orientation);
  out.set("paperSize", setup.paper_size);
  out.set("scalePct", setup.scale_pct);
  out.set("fitToWidth", setup.fit_to_width);
  out.set("fitToHeight", setup.fit_to_height);
  out.set("fitToPage", setup.fit_to_page != 0);
  // Presence flags: the value fields always carry the effective setting,
  // so these are the only way to tell an explicit value from a default.
  out.set("orientationStated", setup.orientation_engaged != 0);
  out.set("paperSizeStated", setup.paper_size_engaged != 0);
  out.set("scalePctStated", setup.scale_pct_engaged != 0);
  out.set("fitToWidthStated", setup.fit_to_width_engaged != 0);
  out.set("fitToHeightStated", setup.fit_to_height_engaged != 0);
  out.set("fitToPageStated", setup.fit_to_page_engaged != 0);
  return out;
}

emscripten::val JsWorkbook::getSheetPageMargins(uint32_t sheet) const {
  fm_page_margins_t margins = FM_STRUCT_INIT(fm_page_margins_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_page_margins(handle_, sheet, &margins) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val out = emscripten::val::object();
  out.set("leftIn", margins.left_in);
  out.set("rightIn", margins.right_in);
  out.set("topIn", margins.top_in);
  out.set("bottomIn", margins.bottom_in);
  out.set("headerIn", margins.header_in);
  out.set("footerIn", margins.footer_in);
  out.set("leftInStated", margins.left_in_engaged != 0);
  out.set("rightInStated", margins.right_in_engaged != 0);
  out.set("topInStated", margins.top_in_engaged != 0);
  out.set("bottomInStated", margins.bottom_in_engaged != 0);
  out.set("headerInStated", margins.header_in_engaged != 0);
  out.set("footerInStated", margins.footer_in_engaged != 0);
  return out;
}

emscripten::val JsWorkbook::getSheetPrintOptions(uint32_t sheet) const {
  fm_print_options_t options = FM_STRUCT_INIT(fm_print_options_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_print_options(handle_, sheet, &options) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val out = emscripten::val::object();
  out.set("gridLines", options.grid_lines != 0);
  out.set("headings", options.headings != 0);
  out.set("horizontalCentered", options.horizontal_centered != 0);
  out.set("verticalCentered", options.vertical_centered != 0);
  out.set("gridLinesStated", options.grid_lines_engaged != 0);
  out.set("headingsStated", options.headings_engaged != 0);
  out.set("horizontalCenteredStated", options.horizontal_centered_engaged != 0);
  out.set("verticalCenteredStated", options.vertical_centered_engaged != 0);
  return out;
}

emscripten::val JsWorkbook::getSheetHeaderFooter(uint32_t sheet) const {
  fm_header_footer_t hf = FM_STRUCT_INIT(fm_header_footer_t);
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_header_footer(handle_, sheet, &hf) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val out = emscripten::val::object();
  js_set_cstr_or_null_fields(out, {{"oddHeader", hf.odd_header},
                                   {"oddFooter", hf.odd_footer},
                                   {"evenHeader", hf.even_header},
                                   {"evenFooter", hf.even_footer},
                                   {"firstHeader", hf.first_header},
                                   {"firstFooter", hf.first_footer}});
  out.set("differentOddEven", hf.different_odd_even != 0);
  out.set("differentFirst", hf.different_first != 0);
  out.set("scaleWithDoc", hf.scale_with_doc != 0);
  out.set("alignWithMargins", hf.align_with_margins != 0);
  out.set("differentOddEvenStated", hf.different_odd_even_engaged != 0);
  out.set("differentFirstStated", hf.different_first_engaged != 0);
  out.set("scaleWithDocStated", hf.scale_with_doc_engaged != 0);
  out.set("alignWithMarginsStated", hf.align_with_margins_engaged != 0);
  return out;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
