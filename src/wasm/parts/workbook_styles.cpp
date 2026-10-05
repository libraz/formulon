//
// JsWorkbook styles surface: per-cell xfIndex get/set, font / fill /
// border / numFmt / xf record getters and adders, named cell-style and
// cellStyleXfs accessors, plus the matching count accessors. The four
// add* paths share the `js_pull_*` helpers in `parts/embind_common.h`
// so the embind glue is emitted once per field type rather than per
// call site.

#include <emscripten/val.h>

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "utils/error.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

void js_pull_cell_xf(const emscripten::val& record, fm_cell_xf* xf) {
  xf->font_index = js_pull_u32(record, "fontIndex", 0U);
  xf->fill_index = js_pull_u32(record, "fillIndex", 0U);
  xf->border_index = js_pull_u32(record, "borderIndex", 0U);
  xf->num_fmt_id = js_pull_u16(record, "numFmtId", 0U);
  xf->horizontal_align = js_pull_u8(record, "horizontalAlign", 0U);
  xf->vertical_align = js_pull_u8(record, "verticalAlign", 2U);
  xf->wrap_text = js_pull_bool(record, "wrapText", false) ? 1 : 0;
  xf->justify_last_line = js_pull_bool(record, "justifyLastLine", false) ? 1 : 0;
  xf->xf_id = js_pull_u32(record, "xfId", 0U);
  xf->apply_number_format = js_pull_bool(record, "applyNumberFormat", false) ? 1 : 0;
  xf->apply_font = js_pull_bool(record, "applyFont", false) ? 1 : 0;
  xf->apply_fill = js_pull_bool(record, "applyFill", false) ? 1 : 0;
  xf->apply_border = js_pull_bool(record, "applyBorder", false) ? 1 : 0;
  xf->apply_alignment = js_pull_bool(record, "applyAlignment", false) ? 1 : 0;
  xf->apply_protection = js_pull_bool(record, "applyProtection", false) ? 1 : 0;
  xf->quote_prefix = js_pull_bool(record, "quotePrefix", false) ? 1 : 0;
  const bool has_supplied_protection = js_has(record, "locked") || js_has(record, "hidden");
  xf->has_protection = js_pull_bool(record, "hasProtection", has_supplied_protection) ? 1 : 0;
  xf->locked = js_pull_bool(record, "locked", true) ? 1 : 0;
  xf->hidden = js_pull_bool(record, "hidden", false) ? 1 : 0;

  const bool has_explicit_horizontal_align = js_has(record, "hasHorizontalAlign");
  const bool has_explicit_vertical_align = js_has(record, "hasVerticalAlign");
  const bool has_explicit_wrap_text = js_has(record, "hasWrapText");
  const bool has_explicit_justify_last_line = js_has(record, "hasJustifyLastLine");
  xf->has_horizontal_align = has_explicit_horizontal_align ? (js_pull_bool(record, "hasHorizontalAlign", false) ? 1 : 0)
                                                           : (js_has(record, "horizontalAlign") ? 1 : 0);
  xf->has_vertical_align = has_explicit_vertical_align ? (js_pull_bool(record, "hasVerticalAlign", false) ? 1 : 0)
                                                       : (js_has(record, "verticalAlign") ? 1 : 0);
  xf->has_wrap_text = has_explicit_wrap_text ? (js_pull_bool(record, "hasWrapText", false) ? 1 : 0)
                                             : (js_has(record, "wrapText") ? 1 : 0);
  xf->has_justify_last_line = has_explicit_justify_last_line
                                  ? (js_pull_bool(record, "hasJustifyLastLine", false) ? 1 : 0)
                                  : (js_has(record, "justifyLastLine") ? 1 : 0);
  const bool has_explicit_alignment = js_has(record, "hasAlignment");
  const bool has_supplied_alignment =
      js_has(record, "horizontalAlign") || js_has(record, "verticalAlign") || js_has(record, "wrapText") ||
      js_has(record, "justifyLastLine") || js_has(record, "textRotation") || js_has(record, "indent") ||
      js_has(record, "relativeIndent") || js_has(record, "shrinkToFit") || js_has(record, "readingOrder") ||
      js_has(record, "hasHorizontalAlign") || js_has(record, "hasVerticalAlign") || js_has(record, "hasWrapText") ||
      js_has(record, "hasJustifyLastLine");
  xf->has_alignment =
      has_explicit_alignment ? (js_pull_bool(record, "hasAlignment", false) ? 1 : 0) : (has_supplied_alignment ? 1 : 0);

  if (js_has(record, "textRotation")) {
    xf->has_text_rotation = 1;
    xf->text_rotation = js_pull_u32(record, "textRotation", 0U);
  }
  if (js_has(record, "indent")) {
    xf->has_indent = 1;
    xf->indent = js_pull_u32(record, "indent", 0U);
  }
  if (js_has(record, "relativeIndent")) {
    xf->has_relative_indent = 1;
    xf->relative_indent = js_pull_i32(record, "relativeIndent", 0);
  }
  if (js_has(record, "shrinkToFit")) {
    xf->has_shrink_to_fit = 1;
    xf->shrink_to_fit = js_pull_bool(record, "shrinkToFit", false) ? 1 : 0;
  }
  if (js_has(record, "readingOrder")) {
    xf->has_reading_order = 1;
    xf->reading_order = js_pull_u32(record, "readingOrder", 0U);
  }
}

/// Builds the JS mirror of a font record. Shared by `getFont` and the
/// `<dxf>` font projection so both surface the same field set.
emscripten::val js_font_record(const fm_font_record& f) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", f.name);
  o.set("size", f.size);
  o.set("colorArgb", f.color_argb);
  o.set("bold", f.bold != 0);
  o.set("italic", f.italic != 0);
  o.set("strike", f.strike != 0);
  o.set("hasBold", f.has_bold != 0);
  o.set("hasItalic", f.has_italic != 0);
  o.set("hasStrike", f.has_strike != 0);
  o.set("underline", static_cast<std::uint32_t>(f.underline));
  o.set("vertAlign", static_cast<std::uint32_t>(f.vert_align));
  o.set("hasFamily", f.has_family != 0);
  o.set("family", static_cast<std::uint32_t>(f.family));
  o.set("hasCharset", f.has_charset != 0);
  o.set("charset", static_cast<std::uint32_t>(f.charset));
  o.set("scheme", static_cast<std::uint32_t>(f.scheme));
  o.set("color", js_color_spec(f.color));
  return o;
}

/// Reads a font record out of a JS object. `name_storage` owns the font
/// name for the duration of the C ABI call, which borrows the pointer.
void js_pull_font_record(const emscripten::val& record, std::string* name_storage, fm_font_record* out) {
  *name_storage = js_pull_string(record, "name");
  out->name = name_storage->c_str();
  out->size = js_pull_double(record, "size", 11.0);
  out->bold = js_pull_bool(record, "bold", false) ? 1 : 0;
  out->italic = js_pull_bool(record, "italic", false) ? 1 : 0;
  out->strike = js_pull_bool(record, "strike", false) ? 1 : 0;
  out->has_bold = js_pull_bool(record, "hasBold", false) ? 1 : 0;
  out->has_italic = js_pull_bool(record, "hasItalic", false) ? 1 : 0;
  out->has_strike = js_pull_bool(record, "hasStrike", false) ? 1 : 0;
  out->underline = js_pull_u8(record, "underline", 0U);
  out->vert_align = js_pull_u8(record, "vertAlign", 0U);
  out->has_family = js_pull_bool(record, "hasFamily", false) ? 1 : 0;
  out->family = js_pull_u8(record, "family", 0U);
  out->has_charset = js_pull_bool(record, "hasCharset", false) ? 1 : 0;
  out->charset = js_pull_u8(record, "charset", 0U);
  out->scheme = js_pull_u8(record, "scheme", 0U);
  out->color_argb = js_pull_u32(record, "colorArgb", 0xFF000000U);
  out->color = js_pull_color_spec(record, "color");
}

emscripten::val js_fill_record(const fm_fill_record& f) {
  emscripten::val o = emscripten::val::object();
  o.set("pattern", static_cast<std::uint32_t>(f.pattern));
  o.set("fgArgb", f.fg_argb);
  o.set("bgArgb", f.bg_argb);
  o.set("fg", js_color_spec(f.fg));
  o.set("bg", js_color_spec(f.bg));
  return o;
}

fm_fill_record js_pull_fill_record(const emscripten::val& record) {
  fm_fill_record fr{};
  fr.pattern = js_pull_u8(record, "pattern", 0U);
  fr.fg_argb = js_pull_u32(record, "fgArgb", 0U);
  fr.bg_argb = js_pull_u32(record, "bgArgb", 0U);
  fr.fg = js_pull_color_spec(record, "fg");
  fr.bg = js_pull_color_spec(record, "bg");
  return fr;
}

emscripten::val js_border_record(const fm_border_record& b) {
  emscripten::val o = emscripten::val::object();
  o.set("left", js_border_side(b.left));
  o.set("right", js_border_side(b.right));
  o.set("top", js_border_side(b.top));
  o.set("bottom", js_border_side(b.bottom));
  o.set("diagonal", js_border_side(b.diagonal));
  o.set("diagonalUp", b.diagonal_up != 0);
  o.set("diagonalDown", b.diagonal_down != 0);
  return o;
}

/// Builds the JS mirror of a cell format. Shared by `getCellXf` and
/// `getCellStyleXf`; a zero-initialised `xf` yields the all-default record
/// those two return on their failure paths, so the key set a caller sees
/// does not depend on whether the read succeeded. `xfId` is set by the
/// callers: a named-style xf is its own parent, so `getCellStyleXf`
/// always reports 0 there.
emscripten::val js_cell_xf_record(const fm_cell_xf& xf) {
  emscripten::val o = emscripten::val::object();
  o.set("fontIndex", xf.font_index);
  o.set("fillIndex", xf.fill_index);
  o.set("borderIndex", xf.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.vertical_align));
  o.set("wrapText", xf.wrap_text != 0);
  o.set("justifyLastLine", xf.justify_last_line != 0);
  o.set("hasAlignment", xf.has_alignment != 0);
  o.set("hasHorizontalAlign", xf.has_horizontal_align != 0);
  o.set("hasVerticalAlign", xf.has_vertical_align != 0);
  o.set("hasWrapText", xf.has_wrap_text != 0);
  o.set("hasJustifyLastLine", xf.has_justify_last_line != 0);
  o.set("applyNumberFormat", xf.apply_number_format != 0);
  o.set("applyFont", xf.apply_font != 0);
  o.set("applyFill", xf.apply_fill != 0);
  o.set("applyBorder", xf.apply_border != 0);
  o.set("applyAlignment", xf.apply_alignment != 0);
  o.set("applyProtection", xf.apply_protection != 0);
  o.set("quotePrefix", xf.quote_prefix != 0);
  o.set("hasProtection", xf.has_protection != 0);
  o.set("locked", xf.locked != 0);
  o.set("hidden", xf.hidden != 0);
  return o;
}

void js_set_cell_xf_alignment(const fm_cell_xf& xf, emscripten::val* object) {
  if (xf.has_text_rotation != 0) {
    object->set("textRotation", xf.text_rotation);
  }
  if (xf.has_indent != 0) {
    object->set("indent", xf.indent);
  }
  if (xf.has_relative_indent != 0) {
    object->set("relativeIndent", xf.relative_indent);
  }
  if (xf.has_shrink_to_fit != 0) {
    object->set("shrinkToFit", xf.shrink_to_fit != 0);
  }
  if (xf.has_reading_order != 0) {
    object->set("readingOrder", xf.reading_order);
  }
}

// Builds the `getCellXf` / `getCellStyleXf` result from the C call's status
// and record; a failed read reports the all-default record.
emscripten::val js_cell_xf_result(fm_status_t rc, fm_cell_xf xf) {
  if (rc != 0) {
    xf = fm_cell_xf{};
  }
  emscripten::val o = js_cell_xf_record(xf);
  o.set("status", status_from_rc(rc));
  o.set("xfId", xf.xf_id);
  js_set_cell_xf_alignment(xf, &o);
  return o;
}

}  // namespace

// ---- Per-cell xf index get/set -----------------------------------------

// Every getter below emits its declared payload keys on all three exit
// paths -- dead handle, non-zero rc, success -- so `status.ok` is the only
// thing a caller has to branch on. A key that is present only on success
// makes the TypeScript declaration a lie and turns a missed status check
// into a `TypeError` instead of a wrong-but-typed value.

emscripten::val JsWorkbook::getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  uint32_t xf = 0;
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    o.set("xfIndex", xf);
    return o;
  }
  fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
  o.set("status", status_from_rc(rc));
  o.set("xfIndex", rc == 0 ? xf : 0U);
  return o;
}

JsStatus JsWorkbook::setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf_index);
  return status_from_rc(rc);
}

JsStatus JsWorkbook::setRangeXfIndex(uint32_t sheet, uint32_t firstRow, uint32_t firstCol, uint32_t lastRow,
                                     uint32_t lastCol, uint32_t xf_index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_sheet_set_range_xf_index(handle_, sheet, firstRow, firstCol, lastRow, lastCol, xf_index));
}

// ---- Style record getters ----------------------------------------------

emscripten::val JsWorkbook::getCellXf(uint32_t xf_index) const {
  fm_cell_xf xf{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_cell_xf(handle_, xf_index, &xf) : 7000;
  return js_cell_xf_result(rc, xf);
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  fm_font_record f{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_font(handle_, font_index, &f) : 7000;
  if (rc != 0) {
    f = fm_font_record{};
  }
  emscripten::val o = js_font_record(f);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getFill(uint32_t fill_index) const {
  fm_fill_record f{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_fill(handle_, fill_index, &f) : 7000;
  if (rc != 0) {
    f = fm_fill_record{};
  }
  emscripten::val o = js_fill_record(f);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getBorder(uint32_t border_index) const {
  fm_border_record b{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_border(handle_, border_index, &b) : 7000;
  if (rc != 0) {
    b = fm_border_record{};
  }
  emscripten::val o = js_border_record(b);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getNumFmt(uint32_t num_fmt_id) const {
  emscripten::val o = emscripten::val::object();
  const char* s = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s) : 7000;
  o.set("status", status_from_rc(rc));
  o.set("numFmtId", rc == 0 ? num_fmt_id : 0U);
  js_set_cstr(o, "formatCode", rc == 0 ? s : nullptr);
  return o;
}

emscripten::val JsWorkbook::getDxf(uint32_t dxf_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_dxf_record d{};
  fm_status_t rc = fm_styles_get_dxf(handle_, dxf_index, &d);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  if (d.font_engaged != 0) {
    o.set("font", js_font_record(d.font));
  }
  if (d.fill_engaged != 0) {
    o.set("fill", js_fill_record(d.fill));
  }
  if (d.border_engaged != 0) {
    o.set("border", js_border_record(d.border));
  }
  if (d.num_fmt_engaged != 0) {
    emscripten::val num_fmt = emscripten::val::object();
    num_fmt.set("numFmtId", static_cast<uint32_t>(d.num_fmt_id));
    js_set_cstr(num_fmt, "formatCode", d.num_fmt_code);
    o.set("numFmt", num_fmt);
  }
  if (d.alignment_xml != nullptr && d.alignment_xml[0] != '\0') {
    o.set("alignmentXml", std::string(d.alignment_xml));
  }
  if (d.protection_xml != nullptr && d.protection_xml[0] != '\0') {
    o.set("protectionXml", std::string(d.protection_xml));
  }
  return o;
}

// ---- Style record adders -----------------------------------------------

JsAddStyleResult JsWorkbook::addFont(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  std::string name;
  fm_font_record fr{};
  js_pull_font_record(record, &name, &fr);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsStatus JsWorkbook::setFont(uint32_t font_index, emscripten::val record) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  std::string name;
  fm_font_record fr{};
  js_pull_font_record(record, &name, &fr);
  return status_from_rc(fm_styles_set_font(handle_, font_index, fr));
}

JsStatus JsWorkbook::setDefaultFont(emscripten::val record) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  std::string name;
  fm_font_record fr{};
  js_pull_font_record(record, &name, &fr);
  return status_from_rc(fm_workbook_set_default_font(handle_, fr));
}

JsAddStyleResult JsWorkbook::addFill(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  const fm_fill_record fr = js_pull_fill_record(record);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddStyleResult JsWorkbook::addBorder(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_border_record br{};
  br.left = js_pull_border_side(record["left"]);
  br.right = js_pull_border_side(record["right"]);
  br.top = js_pull_border_side(record["top"]);
  br.bottom = js_pull_border_side(record["bottom"]);
  br.diagonal = js_pull_border_side(record["diagonal"]);
  br.diagonal_up = js_pull_bool(record, "diagonalUp", false) ? 1 : 0;
  br.diagonal_down = js_pull_bool(record, "diagonalDown", false) ? 1 : 0;
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddNumFmtResult JsWorkbook::addNumFmt(const std::string& format_code) {
  JsAddNumFmtResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  uint16_t id = 0;
  fm_status_t rc = fm_styles_add_num_fmt(handle_, format_code.c_str(), &id);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.numFmtId = id;
  return r;
}

JsAddStyleResult JsWorkbook::addXf(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }
  fm_cell_xf xf{};
  js_pull_cell_xf(record, &xf);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

JsAddStyleResult JsWorkbook::addDxf(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = error_status(7000);
    return r;
  }

  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  fm_dxf_record dxf{};

  emscripten::val font = record["font"];
  if (!font.isUndefined() && !font.isNull()) {
    dxf.font_engaged = 1;
    js_pull_font_record(font, &font_name, &dxf.font);
  }

  emscripten::val fill = record["fill"];
  if (!fill.isUndefined() && !fill.isNull()) {
    dxf.fill_engaged = 1;
    dxf.fill = js_pull_fill_record(fill);
  }

  emscripten::val border = record["border"];
  if (!border.isUndefined() && !border.isNull()) {
    dxf.border_engaged = 1;
    dxf.border.left = js_pull_border_side(border["left"]);
    dxf.border.right = js_pull_border_side(border["right"]);
    dxf.border.top = js_pull_border_side(border["top"]);
    dxf.border.bottom = js_pull_border_side(border["bottom"]);
    dxf.border.diagonal = js_pull_border_side(border["diagonal"]);
    dxf.border.diagonal_up = js_pull_bool(border, "diagonalUp", false) ? 1 : 0;
    dxf.border.diagonal_down = js_pull_bool(border, "diagonalDown", false) ? 1 : 0;
  }

  emscripten::val num_fmt = record["numFmt"];
  if (!num_fmt.isUndefined() && !num_fmt.isNull()) {
    dxf.num_fmt_engaged = 1;
    dxf.num_fmt_id = js_pull_u16(num_fmt, "numFmtId", 0U);
    num_fmt_code = js_pull_string(num_fmt, "formatCode");
    dxf.num_fmt_code = num_fmt_code.c_str();
  }

  alignment_xml = js_pull_string(record, "alignmentXml");
  protection_xml = js_pull_string(record, "protectionXml");
  dxf.alignment_xml = alignment_xml.c_str();
  dxf.protection_xml = protection_xml.c_str();

  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_dxf(handle_, dxf, &idx);
  if (rc != 0) {
    r.status = error_status(rc);
    return r;
  }
  r.status = ok_status();
  r.index = idx;
  return r;
}

// ---- Style count accessors ---------------------------------------------
//
// `fontCount` / `fillCount` / `borderCount` / `xfCount` are now emitted
// by the binding codegen (see `src/wasm/generated/styles_counts.cpp`).

JsNumberResult JsWorkbook::dxfCount() const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  uint32_t n = 0;
  const fm_status_t rc = fm_styles_get_dxf_count(handle_, &n);
  return number_result(rc, static_cast<double>(n));
}
// `cellStyleCount` / `cellStyleXfCount` stay here because they have no
// N-API counterpart and are therefore not part of the cross-binding
// manifest.

JsNumberResult JsWorkbook::cellStyleCount() const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  uint32_t n = 0;
  const fm_status_t rc = fm_styles_get_cell_style_count(handle_, &n);
  return number_result(rc, static_cast<double>(n));
}

JsNumberResult JsWorkbook::cellStyleXfCount() const {
  if (handle_ == nullptr) {
    return number_result(kBindingInvalidHandle, 0.0);
  }
  uint32_t n = 0;
  const fm_status_t rc = fm_styles_get_cell_style_xf_count(handle_, &n);
  return number_result(rc, static_cast<double>(n));
}

emscripten::val JsWorkbook::getCellStyle(uint32_t index) const {
  emscripten::val o = emscripten::val::object();
  fm_cell_style_record_t cs{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_cell_style(handle_, index, &cs) : 7000;
  if (rc != 0) {
    cs = fm_cell_style_record_t{};
  }
  o.set("status", status_from_rc(rc));
  js_set_cstr(o, "name", cs.name);
  o.set("xfId", cs.xf_id);
  o.set("builtinId", cs.builtin_id);
  o.set("iLevel", cs.i_level);
  o.set("hidden", cs.hidden != 0);
  o.set("customBuiltin", cs.custom_builtin != 0);
  return o;
}

emscripten::val JsWorkbook::getCellStyleXf(uint32_t index) const {
  fm_cell_xf xf{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_cell_style_xf(handle_, index, &xf) : 7000;
  return js_cell_xf_result(rc, xf);
}

JsAddStyleResult JsWorkbook::addCellStyleXf(emscripten::val record) {
  JsAddStyleResult out;
  if (handle_ == nullptr) {
    out.status = error_status(7000);
    return out;
  }
  fm_cell_xf xf{};
  js_pull_cell_xf(record, &xf);
  uint32_t index = 0;
  const fm_status_t rc = fm_styles_add_cell_style_xf(handle_, xf, &index);
  out.status = rc == 0 ? ok_status() : error_status(rc);
  out.index = index;
  return out;
}

JsStatus JsWorkbook::setCellStyle(emscripten::val record) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const std::string name = js_pull_string(record, "name");
  fm_cell_style_record_t cs{};
  cs.name = name.c_str();
  cs.xf_id = js_pull_u32(record, "xfId", 0U);
  cs.builtin_id = js_pull_u32(record, "builtinId", FM_CELL_STYLE_BUILTIN_ID_NONE);
  cs.i_level = js_pull_u32(record, "iLevel", 0U);
  cs.hidden = js_pull_bool(record, "hidden", false) ? 1 : 0;
  cs.custom_builtin = js_pull_bool(record, "customBuiltin", false) ? 1 : 0;
  return status_from_rc(fm_styles_set_cell_style(handle_, &cs));
}

JsStatus JsWorkbook::removeCellStyle(const std::string& name) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_styles_remove_cell_style(handle_, name.c_str()));
}

namespace {

emscripten::val js_resolved_color(uint32_t argb, int32_t resolution) {
  emscripten::val o = emscripten::val::object();
  o.set("argb", argb);
  o.set("resolution", resolution);
  return o;
}

}  // namespace

emscripten::val JsWorkbook::getTheme() const {
  fm_theme_colors colors{};
  fm_theme_fonts fonts{};
  int32_t source = 0;
  fm_status_t rc = handle_ != nullptr ? fm_workbook_get_theme_colors(handle_, &colors, &source) : 7000;
  if (rc == 0) {
    rc = fm_workbook_get_theme_fonts(handle_, &fonts);
  }
  if (rc != 0) {
    colors = fm_theme_colors{};
    fonts = fm_theme_fonts{};
    source = 0;
  }
  emscripten::val list = emscripten::val::array();
  for (uint32_t i = 0; i < 12U; ++i) {
    list.set(i, colors.argb[i]);
  }
  emscripten::val f = emscripten::val::object();
  js_set_cstr_fields(f, {{"majorLatin", fonts.major_latin},
                         {"majorEastAsian", fonts.major_east_asian},
                         {"minorLatin", fonts.minor_latin},
                         {"minorEastAsian", fonts.minor_east_asian}});
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("source", source);
  o.set("colors", list);
  o.set("fonts", f);
  return o;
}

JsStatus JsWorkbook::setThemeColors(emscripten::val colors) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_theme_colors tc{};
  if (!js_pull_u32_array(colors, tc.argb, 12U)) {
    return binding_error_status(static_cast<int32_t>(formulon::FormulonErrorCode::kInvalidArgument),
                                "setThemeColors: `colors` must be an array of 12 ARGB numbers");
  }
  return status_from_rc(fm_workbook_set_theme_colors(handle_, &tc));
}

JsStatus JsWorkbook::setThemeFonts(emscripten::val fonts) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  const std::string major_latin = js_pull_string(fonts, "majorLatin");
  const std::string major_ea = js_pull_string(fonts, "majorEastAsian");
  const std::string minor_latin = js_pull_string(fonts, "minorLatin");
  const std::string minor_ea = js_pull_string(fonts, "minorEastAsian");
  fm_theme_fonts tf{};
  tf.major_latin = major_latin.c_str();
  tf.major_east_asian = major_ea.c_str();
  tf.minor_latin = minor_latin.c_str();
  tf.minor_east_asian = minor_ea.c_str();
  return status_from_rc(fm_workbook_set_theme_fonts(handle_, &tf));
}

emscripten::val JsWorkbook::resolveColor(emscripten::val spec, int32_t context) const {
  uint32_t argb = 0;
  int32_t resolution = 0;
  emscripten::val holder = emscripten::val::object();
  holder.set("c", spec);
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_resolve_color(handle_, js_pull_color_spec(holder, "c"),
                                                                        context, &argb, &resolution)
                                            : 7000;
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("argb", rc == 0 ? argb : 0U);
  o.set("resolution", rc == 0 ? resolution : 0);
  return o;
}

emscripten::val JsWorkbook::getEffectiveStyle(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_effective_style e{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_effective_style(handle_, sheet, row, col, &e) : 7000;
  if (rc != 0) {
    e = fm_effective_style{};
  }
  emscripten::val o = emscripten::val::object();
  o.set("status", status_from_rc(rc));
  o.set("xfIndex", e.xf_index);
  o.set("source", e.source);
  o.set("fontIndex", e.font_index);
  o.set("fillIndex", e.fill_index);
  o.set("borderIndex", e.border_index);
  o.set("font", js_resolved_color(e.font_argb, e.font_resolution));
  o.set("fillForeground", js_resolved_color(e.fill_fg_argb, e.fill_fg_resolution));
  o.set("fillBackground", js_resolved_color(e.fill_bg_argb, e.fill_bg_resolution));
  static const char* const kSides[5] = {"left", "right", "top", "bottom", "diagonal"};
  emscripten::val borders = emscripten::val::object();
  for (int i = 0; i < 5; ++i) {
    borders.set(kSides[i], js_resolved_color(e.border_argb[i], e.border_resolution[i]));
  }
  o.set("borders", borders);
  o.set("locked", e.locked != 0);
  o.set("hidden", e.hidden != 0);
  js_set_cstr(o, "numFmtCode", e.num_fmt_code);
  return o;
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
