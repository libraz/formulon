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
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

bool js_has_field(const emscripten::val& object, const char* key) {
  const emscripten::val field = object[key];
  return !field.isUndefined() && !field.isNull();
}

std::int32_t js_pull_i32(const emscripten::val& object, const char* key, std::int32_t dflt) {
  const emscripten::val field = object[key];
  if (field.isUndefined() || field.isNull()) {
    return dflt;
  }
  return field.as<std::int32_t>();
}

void js_pull_cell_xf_ex2(const emscripten::val& record, fm_cell_xf_ex2* xf) {
  xf->base.font_index = js_pull_u32(record, "fontIndex", 0U);
  xf->base.fill_index = js_pull_u32(record, "fillIndex", 0U);
  xf->base.border_index = js_pull_u32(record, "borderIndex", 0U);
  xf->base.num_fmt_id = js_pull_u16(record, "numFmtId", 0U);
  xf->base.horizontal_align = js_pull_u8(record, "horizontalAlign", 0U);
  xf->base.vertical_align = js_pull_u8(record, "verticalAlign", 2U);
  xf->base.wrap_text = js_pull_bool(record, "wrapText", false) ? 1 : 0;
  xf->justify_last_line = js_pull_bool(record, "justifyLastLine", false) ? 1 : 0;
  xf->xf_id = js_pull_u32(record, "xfId", 0U);

  const bool has_explicit_horizontal_align = js_has_field(record, "hasHorizontalAlign");
  const bool has_explicit_vertical_align = js_has_field(record, "hasVerticalAlign");
  const bool has_explicit_wrap_text = js_has_field(record, "hasWrapText");
  const bool has_explicit_justify_last_line = js_has_field(record, "hasJustifyLastLine");
  xf->has_horizontal_align = has_explicit_horizontal_align ? (js_pull_bool(record, "hasHorizontalAlign", false) ? 1 : 0)
                                                           : (js_has_field(record, "horizontalAlign") ? 1 : 0);
  xf->has_vertical_align = has_explicit_vertical_align ? (js_pull_bool(record, "hasVerticalAlign", false) ? 1 : 0)
                                                       : (js_has_field(record, "verticalAlign") ? 1 : 0);
  xf->has_wrap_text = has_explicit_wrap_text ? (js_pull_bool(record, "hasWrapText", false) ? 1 : 0)
                                             : (js_has_field(record, "wrapText") ? 1 : 0);
  xf->has_justify_last_line = has_explicit_justify_last_line
                                  ? (js_pull_bool(record, "hasJustifyLastLine", false) ? 1 : 0)
                                  : (js_has_field(record, "justifyLastLine") ? 1 : 0);
  const bool has_explicit_alignment = js_has_field(record, "hasAlignment");
  const bool has_supplied_alignment = js_has_field(record, "horizontalAlign") ||
                                      js_has_field(record, "verticalAlign") || js_has_field(record, "wrapText") ||
                                      js_has_field(record, "justifyLastLine") || js_has_field(record, "textRotation") ||
                                      js_has_field(record, "indent") || js_has_field(record, "relativeIndent") ||
                                      js_has_field(record, "shrinkToFit") || js_has_field(record, "readingOrder") ||
                                      js_has_field(record, "hasHorizontalAlign") ||
                                      js_has_field(record, "hasVerticalAlign") || js_has_field(record, "hasWrapText") ||
                                      js_has_field(record, "hasJustifyLastLine");
  xf->has_alignment =
      has_explicit_alignment ? (js_pull_bool(record, "hasAlignment", false) ? 1 : 0) : (has_supplied_alignment ? 1 : 0);

  if (js_has_field(record, "textRotation")) {
    xf->has_text_rotation = 1;
    xf->text_rotation = js_pull_u32(record, "textRotation", 0U);
  }
  if (js_has_field(record, "indent")) {
    xf->has_indent = 1;
    xf->indent = js_pull_u32(record, "indent", 0U);
  }
  if (js_has_field(record, "relativeIndent")) {
    xf->has_relative_indent = 1;
    xf->relative_indent = js_pull_i32(record, "relativeIndent", 0);
  }
  if (js_has_field(record, "shrinkToFit")) {
    xf->has_shrink_to_fit = 1;
    xf->shrink_to_fit = js_pull_bool(record, "shrinkToFit", false) ? 1 : 0;
  }
  if (js_has_field(record, "readingOrder")) {
    xf->has_reading_order = 1;
    xf->reading_order = js_pull_u32(record, "readingOrder", 0U);
  }
}

/// Builds the JS mirror of a font record. Shared by `getFont` and the
/// `<dxf>` font projection so both surface the same field set.
emscripten::val js_font_record(const fm_font_record& f) {
  emscripten::val o = emscripten::val::object();
  o.set("name", std::string(f.name != nullptr ? f.name : ""));
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

void js_set_cell_xf_alignment(const fm_cell_xf_ex2& xf, emscripten::val* object) {
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

}  // namespace

// ---- Per-cell xf index get/set -----------------------------------------

emscripten::val JsWorkbook::getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  uint32_t xf = 0;
  fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("xfIndex", xf);
  return o;
}

JsStatus JsWorkbook::setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf_index);
  return status_from_rc(rc);
}

// ---- Style record getters ----------------------------------------------

emscripten::val JsWorkbook::getCellXf(uint32_t xf_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_cell_xf_ex2 xf{};
  fm_status_t rc = fm_styles_get_cell_xf_ex2(handle_, xf_index, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("fontIndex", xf.base.font_index);
  o.set("fillIndex", xf.base.fill_index);
  o.set("borderIndex", xf.base.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.base.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.base.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.base.vertical_align));
  o.set("wrapText", xf.base.wrap_text != 0);
  o.set("justifyLastLine", xf.justify_last_line != 0);
  o.set("hasAlignment", xf.has_alignment != 0);
  o.set("hasHorizontalAlign", xf.has_horizontal_align != 0);
  o.set("hasVerticalAlign", xf.has_vertical_align != 0);
  o.set("hasWrapText", xf.has_wrap_text != 0);
  o.set("hasJustifyLastLine", xf.has_justify_last_line != 0);
  o.set("xfId", xf.xf_id);
  js_set_cell_xf_alignment(xf, &o);
  return o;
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_font_record f{};
  fm_status_t rc = fm_styles_get_font(handle_, font_index, &f);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o = js_font_record(f);
  o.set("status", ok_status());
  return o;
}

emscripten::val JsWorkbook::getFill(uint32_t fill_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_fill_record f{};
  fm_status_t rc = fm_styles_get_fill(handle_, fill_index, &f);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o = js_fill_record(f);
  o.set("status", ok_status());
  return o;
}

emscripten::val JsWorkbook::getBorder(uint32_t border_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_border_record b{};
  fm_status_t rc = fm_styles_get_border(handle_, border_index, &b);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o = js_border_record(b);
  o.set("status", ok_status());
  return o;
}

emscripten::val JsWorkbook::getNumFmt(uint32_t num_fmt_id) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  const char* s = nullptr;
  fm_status_t rc = fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("numFmtId", num_fmt_id);
  o.set("formatCode", std::string(s != nullptr ? s : ""));
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
    num_fmt.set("formatCode", std::string(d.num_fmt_code != nullptr ? d.num_fmt_code : ""));
    o.set("numFmt", num_fmt);
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
  fm_cell_xf_ex2 xf{};
  js_pull_cell_xf_ex2(record, &xf);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf_ex2(handle_, xf, &idx);
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

uint32_t JsWorkbook::dxfCount() const {
  if (handle_ == nullptr) {
    return 0;
  }
  uint32_t n = 0;
  if (fm_styles_get_dxf_count(handle_, &n) != 0) {
    return 0;
  }
  return n;
}
// `cellStyleCount` / `cellStyleXfCount` stay here because they have no
// N-API counterpart and are therefore not part of the cross-binding
// manifest.

uint32_t JsWorkbook::cellStyleCount() const {
  if (handle_ == nullptr) {
    return 0U;
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_style_count(handle_, &n) != 0) {
    return 0U;
  }
  return n;
}

uint32_t JsWorkbook::cellStyleXfCount() const {
  if (handle_ == nullptr) {
    return 0U;
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_style_xf_count(handle_, &n) != 0) {
    return 0U;
  }
  return n;
}

emscripten::val JsWorkbook::getCellStyle(uint32_t index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_cell_style_record_t cs{};
  fm_status_t rc = fm_styles_get_cell_style(handle_, index, &cs);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("name", std::string(cs.name != nullptr ? cs.name : ""));
  o.set("xfId", cs.xf_id);
  o.set("builtinId", cs.builtin_id);
  o.set("iLevel", cs.i_level);
  o.set("hidden", cs.hidden != 0);
  o.set("customBuiltin", cs.custom_builtin != 0);
  return o;
}

emscripten::val JsWorkbook::getCellStyleXf(uint32_t index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", error_status(7000));
    return o;
  }
  fm_cell_xf_ex2 xf{};
  fm_status_t rc = fm_styles_get_cell_style_xf_ex2(handle_, index, &xf);
  if (rc != 0) {
    o.set("status", error_status(rc));
    return o;
  }
  o.set("status", ok_status());
  o.set("fontIndex", xf.base.font_index);
  o.set("fillIndex", xf.base.fill_index);
  o.set("borderIndex", xf.base.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.base.num_fmt_id));
  o.set("horizontalAlign", static_cast<uint32_t>(xf.base.horizontal_align));
  o.set("verticalAlign", static_cast<uint32_t>(xf.base.vertical_align));
  o.set("wrapText", xf.base.wrap_text != 0);
  o.set("justifyLastLine", xf.justify_last_line != 0);
  o.set("hasAlignment", xf.has_alignment != 0);
  o.set("hasHorizontalAlign", xf.has_horizontal_align != 0);
  o.set("hasVerticalAlign", xf.has_vertical_align != 0);
  o.set("hasWrapText", xf.has_wrap_text != 0);
  o.set("hasJustifyLastLine", xf.has_justify_last_line != 0);
  js_set_cell_xf_alignment(xf, &o);
  return o;
}

JsAddStyleResult JsWorkbook::addCellStyleXf(emscripten::val record) {
  JsAddStyleResult out;
  if (handle_ == nullptr) {
    out.status = error_status(7000);
    return out;
  }
  fm_cell_xf_ex2 xf{};
  js_pull_cell_xf_ex2(record, &xf);
  uint32_t index = 0;
  const fm_status_t rc = fm_styles_add_cell_style_xf_ex2(handle_, xf, &index);
  out.status = rc == 0 ? ok_status() : error_status(rc);
  out.index = index;
  return out;
}

JsStatus JsWorkbook::setCellStyle(const std::string& name, uint32_t xfId, uint32_t builtinId) {
  if (handle_ == nullptr) {
    return error_status(7000);
  }
  return status_from_rc(fm_styles_set_cell_style(handle_, name.c_str(), xfId, builtinId));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
