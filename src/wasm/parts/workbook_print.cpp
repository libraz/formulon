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
// deriving its `has_*` flag. This keeps getter/proxy failures and lossy
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
      manual = brk.manual == FM_TRISTATE_TRUE ? 1 : 0;
    } else {
      fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
      rc = fm_sheet_get_column_break_at(handle, sheet, i, &brk);
      position = brk.col;
      first = brk.first;
      last = brk.last;
      manual = brk.manual == FM_TRISTATE_TRUE ? 1 : 0;
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
  fm_color_t color{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_tab_color(handle_, sheet, &color) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_color_spec_or_null(color);
}

// `null` clears the tab colour; an object is read as an `fm_color_t`.
void JsWorkbook::setSheetTabColor(uint32_t sheet, emscripten::val color) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (color.isNull()) {
    succeeded(fm_sheet_set_tab_color(handle_, sheet, nullptr));
    return;
  }
  JsNarrowNumericReader reader("Sheet.setTabColor");
  const fm_color_t spec = js_pull_color_spec_value(color, "color", &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_tab_color(handle_, sheet, &spec));
}

// ---- Print area / titles -----------------------------------------------

emscripten::val JsWorkbook::getSheetPrintArea(uint32_t sheet) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  size_t count = 0;
  fm_status_t rc = fm_sheet_get_print_area(handle_, sheet, nullptr, 0, &count);
  std::vector<fm_cell_rect_t> found(count);
  if (rc == 0 && count > 0) {
    rc = fm_sheet_get_print_area(handle_, sheet, found.data(), found.size(), &count);
  }
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return cell_rects_to_val(found.data(), found.size());
}

void JsWorkbook::setSheetPrintArea(uint32_t sheet, emscripten::val ranges) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.setPrintArea");
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
  JsNarrowNumericReader reader("Sheet.setRowBreak");
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.first = reader.u32(input, "first", 0U, "first");
  brk.last = reader.u32(input, "last", 0U, "last");
  brk.manual = reader.boolean(input, "manual", true, "manual") ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
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
  JsNarrowNumericReader reader("Sheet.setColumnBreak");
  fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
  brk.first = reader.u32(input, "first", 0U, "first");
  brk.last = reader.u32(input, "last", 0U, "last");
  brk.manual = reader.boolean(input, "manual", true, "manual") ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
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
  JsNarrowNumericReader reader("Sheet.updatePageSetup");
  fm_page_setup_patch_t out = FM_STRUCT_INIT(fm_page_setup_patch_t);
  js_patch_int(reader, setup, "orientation", "orientation", true, &out.orientation_op, &out.orientation);
  js_patch_u32(reader, setup, "paperSize", "paperSize", false, &out.paper_size_op, &out.paper_size);
  js_patch_u32(reader, setup, "scalePct", "scalePct", false, &out.scale_pct_op, &out.scale_pct);
  js_patch_u32(reader, setup, "fitToWidth", "fitToWidth", false, &out.fit_to_width_op, &out.fit_to_width);
  js_patch_u32(reader, setup, "fitToHeight", "fitToHeight", false, &out.fit_to_height_op, &out.fit_to_height);
  js_patch_flag(reader, setup, "fitToPage", "fitToPage", &out.fit_to_page_op, &out.fit_to_page);
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
  JsNarrowNumericReader reader("Sheet.updatePageMargins");
  fm_page_margins_patch_t out = FM_STRUCT_INIT(fm_page_margins_patch_t);
  js_patch_number(reader, margins, "leftIn", "leftIn", false, &out.left_in_op, &out.left_in);
  js_patch_number(reader, margins, "rightIn", "rightIn", false, &out.right_in_op, &out.right_in);
  js_patch_number(reader, margins, "topIn", "topIn", false, &out.top_in_op, &out.top_in);
  js_patch_number(reader, margins, "bottomIn", "bottomIn", false, &out.bottom_in_op, &out.bottom_in);
  js_patch_number(reader, margins, "headerIn", "headerIn", false, &out.header_in_op, &out.header_in);
  js_patch_number(reader, margins, "footerIn", "footerIn", false, &out.footer_in_op, &out.footer_in);
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
  JsNarrowNumericReader reader("Sheet.updatePrintOptions");
  fm_print_options_patch_t out = FM_STRUCT_INIT(fm_print_options_patch_t);
  js_patch_flag(reader, options, "gridLines", "gridLines", &out.grid_lines_op, &out.grid_lines);
  js_patch_flag(reader, options, "headings", "headings", &out.headings_op, &out.headings);
  js_patch_flag(reader, options, "horizontalCentered", "horizontalCentered", &out.horizontal_centered_op,
                &out.horizontal_centered);
  js_patch_flag(reader, options, "verticalCentered", "verticalCentered", &out.vertical_centered_op,
                &out.vertical_centered);
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
  // The six section strings must outlive the call, so their storage is declared here.
  std::string odd_header;
  std::string odd_footer;
  std::string even_header;
  std::string even_footer;
  std::string first_header;
  std::string first_footer;
  JsNarrowNumericReader reader("Sheet.updateHeaderFooter");
  fm_header_footer_patch_t out = FM_STRUCT_INIT(fm_header_footer_patch_t);
  const auto section = [&](const char* key, std::string& storage, fm_patch_op_t* op, const char** value) {
    js_patch_string(reader, headerFooter, key, key, true, storage, op, value);
  };
  section("oddHeader", odd_header, &out.odd_header_op, &out.odd_header);
  section("oddFooter", odd_footer, &out.odd_footer_op, &out.odd_footer);
  section("evenHeader", even_header, &out.even_header_op, &out.even_header);
  section("evenFooter", even_footer, &out.even_footer_op, &out.even_footer);
  section("firstHeader", first_header, &out.first_header_op, &out.first_header);
  section("firstFooter", first_footer, &out.first_footer_op, &out.first_footer);
  js_patch_flag(reader, headerFooter, "differentOddEven", "differentOddEven", &out.different_odd_even_op,
                &out.different_odd_even);
  js_patch_flag(reader, headerFooter, "differentFirst", "differentFirst", &out.different_first_op,
                &out.different_first);
  js_patch_flag(reader, headerFooter, "scaleWithDoc", "scaleWithDoc", &out.scale_with_doc_op, &out.scale_with_doc);
  js_patch_flag(reader, headerFooter, "alignWithMargins", "alignWithMargins", &out.align_with_margins_op,
                &out.align_with_margins);
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
  js_set_present_or_null(out, "orientation", setup.has_orientation, setup.orientation);
  out.set("paperSize", setup.paper_size);
  out.set("scalePct", setup.scale_pct);
  out.set("fitToWidth", setup.fit_to_width);
  out.set("fitToHeight", setup.fit_to_height);
  out.set("fitToPage", setup.fit_to_page != 0);
  // Presence flags: the value fields always carry the effective setting,
  // so these are the only way to tell an explicit value from a default.
  out.set("paperSizeStated", setup.has_paper_size != 0);
  out.set("scalePctStated", setup.has_scale_pct != 0);
  out.set("fitToWidthStated", setup.has_fit_to_width != 0);
  out.set("fitToHeightStated", setup.has_fit_to_height != 0);
  out.set("fitToPageStated", setup.has_fit_to_page != 0);
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
  out.set("leftInStated", margins.has_left_in != 0);
  out.set("rightInStated", margins.has_right_in != 0);
  out.set("topInStated", margins.has_top_in != 0);
  out.set("bottomInStated", margins.has_bottom_in != 0);
  out.set("headerInStated", margins.has_header_in != 0);
  out.set("footerInStated", margins.has_footer_in != 0);
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
  out.set("gridLinesStated", options.has_grid_lines != 0);
  out.set("headingsStated", options.has_headings != 0);
  out.set("horizontalCenteredStated", options.has_horizontal_centered != 0);
  out.set("verticalCenteredStated", options.has_vertical_centered != 0);
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
  out.set("differentOddEvenStated", hf.has_different_odd_even != 0);
  out.set("differentFirstStated", hf.has_different_first != 0);
  out.set("scaleWithDocStated", hf.has_scale_with_doc != 0);
  out.set("alignWithMarginsStated", hf.has_align_with_margins != 0);
  return out;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
