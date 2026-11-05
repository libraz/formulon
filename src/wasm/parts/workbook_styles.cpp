//
// JsWorkbook styles surface: per-cell xfIndex get/set, font / fill /
// border / numFmt / xf record getters and adders, named cell-style and
// cellStyleXfs accessors, plus the matching count accessors. The four
// add* paths share the `js_pull_*` helpers in `parts/embind_common.h`
// so the embind glue is emitted once per field type rather than per
// call site.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

void js_pull_cell_xf(const emscripten::val& record, fm_cell_xf_t* xf, JsNarrowNumericReader& reader) {
  xf->font_index = reader.u32(record, "fontIndex", 0U, "xf.fontIndex");
  xf->fill_index = reader.u32(record, "fillIndex", 0U, "xf.fillIndex");
  xf->border_index = reader.u32(record, "borderIndex", 0U, "xf.borderIndex");
  xf->num_fmt_id = reader.u16(record, "numFmtId", 0U, "xf.numFmtId");
  const emscripten::val horizontal_align = reader.value(record, "horizontalAlign", "xf.horizontalAlign");
  const emscripten::val vertical_align = reader.value(record, "verticalAlign", "xf.verticalAlign");
  const emscripten::val wrap_text = reader.value(record, "wrapText", "xf.wrapText");
  const emscripten::val justify_last_line = reader.value(record, "justifyLastLine", "xf.justifyLastLine");
  const bool horizontal_align_present = js_value_present(horizontal_align);
  const bool vertical_align_present = js_value_present(vertical_align);
  const bool wrap_text_present = js_value_present(wrap_text);
  const bool justify_last_line_present = js_value_present(justify_last_line);
  xf->horizontal_align = reader.u8_value(horizontal_align, 0U, "xf.horizontalAlign");
  xf->vertical_align = reader.u8_value(vertical_align, 2U, "xf.verticalAlign");
  xf->wrap_text = reader.boolean_value(wrap_text, false) ? 1 : 0;
  xf->justify_last_line = reader.boolean_value(justify_last_line, false) ? 1 : 0;
  xf->xf_id = reader.u32(record, "xfId", 0U, "xf.xfId");
  const emscripten::val has_horizontal_align = reader.value(record, "hasHorizontalAlign", "xf.hasHorizontalAlign");
  const emscripten::val has_vertical_align = reader.value(record, "hasVerticalAlign", "xf.hasVerticalAlign");
  const emscripten::val has_wrap_text = reader.value(record, "hasWrapText", "xf.hasWrapText");
  const emscripten::val has_justify_last_line = reader.value(record, "hasJustifyLastLine", "xf.hasJustifyLastLine");
  const bool has_horizontal_align_present = js_value_present(has_horizontal_align);
  const bool has_vertical_align_present = js_value_present(has_vertical_align);
  const bool has_wrap_text_present = js_value_present(has_wrap_text);
  const bool has_justify_last_line_present = js_value_present(has_justify_last_line);
  xf->has_horizontal_align = reader.boolean_value(has_horizontal_align, horizontal_align_present) ? 1 : 0;
  xf->has_vertical_align = reader.boolean_value(has_vertical_align, vertical_align_present) ? 1 : 0;
  xf->has_wrap_text = reader.boolean_value(has_wrap_text, wrap_text_present) ? 1 : 0;
  xf->has_justify_last_line = reader.boolean_value(has_justify_last_line, justify_last_line_present) ? 1 : 0;
  const emscripten::val has_alignment = reader.value(record, "hasAlignment", "xf.hasAlignment");
  const emscripten::val text_rotation = reader.value(record, "textRotation", "xf.textRotation");
  const emscripten::val indent = reader.value(record, "indent", "xf.indent");
  const emscripten::val relative_indent = reader.value(record, "relativeIndent", "xf.relativeIndent");
  const emscripten::val shrink_to_fit = reader.value(record, "shrinkToFit", "xf.shrinkToFit");
  const emscripten::val reading_order = reader.value(record, "readingOrder", "xf.readingOrder");
  const bool text_rotation_present = js_value_present(text_rotation);
  const bool indent_present = js_value_present(indent);
  const bool relative_indent_present = js_value_present(relative_indent);
  const bool shrink_to_fit_present = js_value_present(shrink_to_fit);
  const bool reading_order_present = js_value_present(reading_order);
  xf->text_rotation = reader.u32_value(text_rotation, 0U, "xf.textRotation");
  xf->indent = reader.u32_value(indent, 0U, "xf.indent");
  xf->relative_indent = reader.i32_value(relative_indent, 0, "xf.relativeIndent");
  xf->shrink_to_fit = reader.boolean_value(shrink_to_fit, false) ? 1 : 0;
  xf->reading_order = reader.u32_value(reading_order, 0U, "xf.readingOrder");
  xf->has_text_rotation = text_rotation_present ? 1 : 0;
  xf->has_indent = indent_present ? 1 : 0;
  xf->has_relative_indent = relative_indent_present ? 1 : 0;
  xf->has_shrink_to_fit = shrink_to_fit_present ? 1 : 0;
  xf->has_reading_order = reading_order_present ? 1 : 0;
  const bool has_supplied_alignment = horizontal_align_present || vertical_align_present || wrap_text_present ||
                                      justify_last_line_present || text_rotation_present || indent_present ||
                                      relative_indent_present || shrink_to_fit_present || reading_order_present ||
                                      has_horizontal_align_present || has_vertical_align_present ||
                                      has_wrap_text_present || has_justify_last_line_present;
  xf->has_alignment = reader.boolean_value(has_alignment, has_supplied_alignment) ? 1 : 0;

  xf->apply_number_format = reader.boolean(record, "applyNumberFormat", false, "xf.applyNumberFormat") ? 1 : 0;
  xf->apply_font = reader.boolean(record, "applyFont", false, "xf.applyFont") ? 1 : 0;
  xf->apply_fill = reader.boolean(record, "applyFill", false, "xf.applyFill") ? 1 : 0;
  xf->apply_border = reader.boolean(record, "applyBorder", false, "xf.applyBorder") ? 1 : 0;
  xf->apply_alignment = reader.boolean(record, "applyAlignment", false, "xf.applyAlignment") ? 1 : 0;
  xf->apply_protection = reader.boolean(record, "applyProtection", false, "xf.applyProtection") ? 1 : 0;
  xf->quote_prefix = reader.boolean(record, "quotePrefix", false, "xf.quotePrefix") ? 1 : 0;
  const emscripten::val has_protection = reader.value(record, "hasProtection", "xf.hasProtection");
  const emscripten::val locked = reader.value(record, "locked", "xf.locked");
  const emscripten::val hidden = reader.value(record, "hidden", "xf.hidden");
  const bool locked_present = js_value_present(locked);
  const bool hidden_present = js_value_present(hidden);
  xf->has_protection = reader.boolean_value(has_protection, locked_present || hidden_present) ? 1 : 0;
  xf->locked = reader.boolean_value(locked, true) ? 1 : 0;
  xf->hidden = reader.boolean_value(hidden, false) ? 1 : 0;
}

/// Diagnostic field names for the members of one border record.
struct BorderFieldNames {
  const char* left;
  const char* right;
  const char* top;
  const char* bottom;
  const char* diagonal;
  const char* diagonal_up;
  const char* diagonal_down;
};

constexpr BorderFieldNames kBorderFields{"border.left",     "border.right",      "border.top",         "border.bottom",
                                         "border.diagonal", "border.diagonalUp", "border.diagonalDown"};
constexpr BorderFieldNames kDxfBorderFields{"dxf.border.left",        "dxf.border.right",    "dxf.border.top",
                                            "dxf.border.bottom",      "dxf.border.diagonal", "dxf.border.diagonalUp",
                                            "dxf.border.diagonalDown"};

fm_border_record_t js_pull_border_record(const emscripten::val& record, const BorderFieldNames& names,
                                         JsNarrowNumericReader& reader) {
  fm_border_record_t br{};
  br.left = js_pull_border_side(reader.value(record, "left", names.left), &reader);
  br.right = js_pull_border_side(reader.value(record, "right", names.right), &reader);
  br.top = js_pull_border_side(reader.value(record, "top", names.top), &reader);
  br.bottom = js_pull_border_side(reader.value(record, "bottom", names.bottom), &reader);
  br.diagonal = js_pull_border_side(reader.value(record, "diagonal", names.diagonal), &reader);
  br.diagonal_up = reader.boolean(record, "diagonalUp", false, names.diagonal_up) ? 1 : 0;
  br.diagonal_down = reader.boolean(record, "diagonalDown", false, names.diagonal_down) ? 1 : 0;
  return br;
}

/// Builds the JS mirror of a font record. Shared by `getFont` and the
/// `<dxf>` font projection so both surface the same field set.
emscripten::val js_font_record(const fm_font_record_t& f) {
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
  o.set("hasSize", f.has_size != 0);
  o.set("scheme", static_cast<std::uint32_t>(f.scheme));
  o.set("outline", f.outline != 0);
  o.set("shadow", f.shadow != 0);
  o.set("condense", f.condense != 0);
  o.set("extend", f.extend != 0);
  o.set("hasOutline", f.has_outline != 0);
  o.set("hasShadow", f.has_shadow != 0);
  o.set("hasCondense", f.has_condense != 0);
  o.set("hasExtend", f.has_extend != 0);
  o.set("color", js_color_spec(f.color));
  return o;
}

/// Reads a font record out of a JS object. `name_storage` owns the font
/// name for the duration of the C ABI call, which borrows the pointer.
void js_pull_font_record(const emscripten::val& record, std::string* name_storage, fm_font_record_t* out,
                         JsNarrowNumericReader& reader) {
  *name_storage = reader.string(record, "name", "font.name");
  out->name = name_storage->c_str();
  out->size = reader.number(record, "size", 11.0, "font.size");
  out->bold = reader.boolean(record, "bold", false, "font.bold") ? 1 : 0;
  out->italic = reader.boolean(record, "italic", false, "font.italic") ? 1 : 0;
  out->strike = reader.boolean(record, "strike", false, "font.strike") ? 1 : 0;
  out->has_bold = reader.boolean(record, "hasBold", false, "font.hasBold") ? 1 : 0;
  out->has_italic = reader.boolean(record, "hasItalic", false, "font.hasItalic") ? 1 : 0;
  out->has_strike = reader.boolean(record, "hasStrike", false, "font.hasStrike") ? 1 : 0;
  out->underline = reader.u8(record, "underline", 0U, "font.underline");
  out->vert_align = reader.u8(record, "vertAlign", 0U, "font.vertAlign");
  out->has_family = reader.boolean(record, "hasFamily", false, "font.hasFamily") ? 1 : 0;
  out->family = reader.u8(record, "family", 0U, "font.family");
  out->has_charset = reader.boolean(record, "hasCharset", false, "font.hasCharset") ? 1 : 0;
  out->charset = reader.u8(record, "charset", 0U, "font.charset");
  // An omitted size means "leave the size unchanged" unless hasSize says otherwise.
  out->has_size =
      reader.boolean(record, "hasSize", reader.has_own(record, "size", "font.size"), "font.hasSize") ? 1 : 0;
  out->scheme = reader.u8(record, "scheme", 0U, "font.scheme");
  out->outline = reader.boolean(record, "outline", false, "font.outline") ? 1 : 0;
  out->shadow = reader.boolean(record, "shadow", false, "font.shadow") ? 1 : 0;
  out->condense = reader.boolean(record, "condense", false, "font.condense") ? 1 : 0;
  out->extend = reader.boolean(record, "extend", false, "font.extend") ? 1 : 0;
  out->has_outline = reader.boolean(record, "hasOutline", false, "font.hasOutline") ? 1 : 0;
  out->has_shadow = reader.boolean(record, "hasShadow", false, "font.hasShadow") ? 1 : 0;
  out->has_condense = reader.boolean(record, "hasCondense", false, "font.hasCondense") ? 1 : 0;
  out->has_extend = reader.boolean(record, "hasExtend", false, "font.hasExtend") ? 1 : 0;
  out->color_argb = reader.u32(record, "colorArgb", 0xFF000000U, "font.colorArgb");
  out->color = js_pull_color_spec(record, "color", &reader);
}

emscripten::val js_fill_record(const fm_fill_record_t& f) {
  emscripten::val o = emscripten::val::object();
  o.set("pattern", static_cast<std::uint32_t>(f.pattern));
  o.set("fgArgb", f.fg_argb);
  o.set("bgArgb", f.bg_argb);
  o.set("fg", js_color_spec(f.fg));
  o.set("bg", js_color_spec(f.bg));
  if (f.has_gradient != 0) {
    emscripten::val gradient = emscripten::val::object();
    gradient.set("type", static_cast<std::int32_t>(f.gradient_type));
    gradient.set("degree", f.gradient_degree);
    gradient.set("left", f.gradient_left);
    gradient.set("right", f.gradient_right);
    gradient.set("top", f.gradient_top);
    gradient.set("bottom", f.gradient_bottom);
    emscripten::val stops = emscripten::val::array();
    for (std::size_t i = 0; i < f.gradient_stop_count; ++i) {
      emscripten::val stop = emscripten::val::object();
      stop.set("position", f.gradient_stops[i].position);
      stop.set("colorArgb", f.gradient_stops[i].color_argb);
      stop.set("color", js_color_spec(f.gradient_stops[i].color));
      stops.set(static_cast<std::uint32_t>(i), stop);
    }
    gradient.set("stops", stops);
    o.set("gradient", gradient);
  }
  return o;
}

/// Reads a fill record out of a JS object. `stops` owns the gradient stops
/// for the duration of the C ABI call, which borrows the pointer.
fm_fill_record_t js_pull_fill_record(const emscripten::val& record, std::vector<fm_gradient_stop_t>* stops,
                                     JsNarrowNumericReader& reader) {
  fm_fill_record_t fr{};
  fr.pattern = reader.u8(record, "pattern", 0U, "fill.pattern");
  fr.fg_argb = reader.u32(record, "fgArgb", 0U, "fill.fgArgb");
  fr.bg_argb = reader.u32(record, "bgArgb", 0U, "fill.bgArgb");
  fr.fg = js_pull_color_spec(record, "fg", &reader);
  fr.bg = js_pull_color_spec(record, "bg", &reader);
  const emscripten::val gradient = reader.value(record, "gradient", "fill.gradient");
  if (js_value_present(gradient)) {
    fr.has_gradient = 1;
    fr.gradient_type = static_cast<std::int32_t>(reader.u32(gradient, "type", 0U, "fill.gradient.type"));
    fr.gradient_degree = reader.number(gradient, "degree", 0.0, "fill.gradient.degree");
    fr.gradient_left = reader.number(gradient, "left", 0.0, "fill.gradient.left");
    fr.gradient_right = reader.number(gradient, "right", 0.0, "fill.gradient.right");
    fr.gradient_top = reader.number(gradient, "top", 0.0, "fill.gradient.top");
    fr.gradient_bottom = reader.number(gradient, "bottom", 0.0, "fill.gradient.bottom");
    const emscripten::val list = reader.value(gradient, "stops", "fill.gradient.stops");
    if (reader.is_array(list, "fill.gradient.stops")) {
      const uint32_t n = reader.length(list, "fill.gradient.stops");
      stops->reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        const emscripten::val item = reader.array_element(list, i, "fill.gradient.stops[]");
        fm_gradient_stop_t stop{};
        stop.position = reader.number(item, "position", 0.0, "fill.gradient.stops[].position");
        stop.color_argb = reader.u32(item, "colorArgb", 0U, "fill.gradient.stops[].colorArgb");
        stop.color = js_pull_color_spec(item, "color", &reader);
        stops->push_back(stop);
      }
    }
    fr.gradient_stops = stops->empty() ? nullptr : stops->data();
    fr.gradient_stop_count = stops->size();
  }
  return fr;
}

emscripten::val js_border_record(const fm_border_record_t& b) {
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
emscripten::val js_cell_xf_record(const fm_cell_xf_t& xf) {
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
  // Without a `<protection>` child the model default applies: locked, not
  // hidden. This also fixes the all-default record a failed read reports.
  o.set("locked", xf.has_protection == 0 || xf.locked != 0);
  o.set("hidden", xf.has_protection != 0 && xf.hidden != 0);
  return o;
}

void js_set_cell_xf_alignment(const fm_cell_xf_t& xf, emscripten::val* object) {
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
emscripten::val js_cell_xf_result(fm_status_t rc, fm_cell_xf_t xf) {
  if (rc != 0) {
    xf = fm_cell_xf_t{};
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

JsNumberResult JsWorkbook::getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
  std::size_t xf = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_cell_xf_index(handle_, sheet, row, col, &xf) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(xf));
}

JsStatus JsWorkbook::setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_set_cell_xf_index(handle_, sheet, row, col, xf_index)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::clearCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) {
  return status_from_rc(handle_ != nullptr ? fm_sheet_clear_cell_xf_index(handle_, sheet, row, col)
                                           : kBindingInvalidHandle);
}

JsStatus JsWorkbook::setRangeXfIndex(uint32_t sheet, emscripten::val range, uint32_t xf_index) {
  if (handle_ == nullptr) {
    return status_from_rc(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setRangeXfIndex");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_sheet_set_range_xf_index(handle_, sheet, &rect, xf_index));
}

// ---- Style record getters ----------------------------------------------

emscripten::val JsWorkbook::getCellXf(uint32_t xf_index) const {
  fm_cell_xf_t xf{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_cell_xf(handle_, xf_index, &xf) : kBindingInvalidHandle;
  return js_cell_xf_result(rc, xf);
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  fm_font_record_t f{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_font(handle_, font_index, &f) : kBindingInvalidHandle;
  if (rc != 0) {
    f = fm_font_record_t{};
  }
  emscripten::val o = js_font_record(f);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getFill(uint32_t fill_index) const {
  fm_fill_record_t f{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_fill(handle_, fill_index, &f) : kBindingInvalidHandle;
  if (rc != 0) {
    f = fm_fill_record_t{};
  }
  emscripten::val o = js_fill_record(f);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getBorder(uint32_t border_index) const {
  fm_border_record_t b{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_border(handle_, border_index, &b) : kBindingInvalidHandle;
  if (rc != 0) {
    b = fm_border_record_t{};
  }
  emscripten::val o = js_border_record(b);
  o.set("status", status_from_rc(rc));
  return o;
}

emscripten::val JsWorkbook::getNumFmt(uint32_t num_fmt_id) const {
  int32_t present = 0;
  const char* s = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_num_fmt_string(handle_, num_fmt_id, &present, &s) : kBindingInvalidHandle;
  return optional_result(rc, present,
                         rc == 0 && present != 0 ? emscripten::val(std::string(s)) : emscripten::val::null());
}

emscripten::val JsWorkbook::getDxf(uint32_t dxf_index) const {
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    return o;
  }
  fm_dxf_record_t d{};
  fm_status_t rc = fm_workbook_get_dxf(handle_, dxf_index, &d);
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
    r.status = invalid_handle_status();
    return r;
  }
  std::string name;
  fm_font_record_t fr{};
  JsNarrowNumericReader reader("addFont");
  js_pull_font_record(record, &name, &fr, reader);
  if (!reader.ok()) {
    r.status = reader.failure();
    return r;
  }
  std::size_t idx = 0;
  const fm_status_t rc = fm_workbook_add_font(handle_, &fr, &idx);
  return index_result(rc, idx);
}

JsStatus JsWorkbook::setFont(uint32_t font_index, emscripten::val record) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  std::string name;
  fm_font_record_t fr{};
  JsNarrowNumericReader reader("setFont");
  js_pull_font_record(record, &name, &fr, reader);
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_workbook_set_font(handle_, font_index, &fr));
}

JsStatus JsWorkbook::setDefaultFont(emscripten::val record) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  std::string name;
  fm_font_record_t fr{};
  JsNarrowNumericReader reader("setDefaultFont");
  js_pull_font_record(record, &name, &fr, reader);
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_workbook_set_default_font(handle_, &fr));
}

JsAddStyleResult JsWorkbook::addFill(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  JsNarrowNumericReader reader("addFill");
  std::vector<fm_gradient_stop_t> stops;
  const fm_fill_record_t fr = js_pull_fill_record(record, &stops, reader);
  if (!reader.ok()) {
    r.status = reader.failure();
    return r;
  }
  std::size_t idx = 0;
  const fm_status_t rc = fm_workbook_add_fill(handle_, &fr, &idx);
  return index_result(rc, idx);
}

JsAddStyleResult JsWorkbook::addBorder(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  JsNarrowNumericReader reader("addBorder");
  const fm_border_record_t br = js_pull_border_record(record, kBorderFields, reader);
  if (!reader.ok()) {
    r.status = reader.failure();
    return r;
  }
  std::size_t idx = 0;
  const fm_status_t rc = fm_workbook_add_border(handle_, &br, &idx);
  return index_result(rc, idx);
}

JsAddNumFmtResult JsWorkbook::addNumFmt(const std::string& format_code) {
  JsAddNumFmtResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }
  uint32_t id = 0;
  fm_status_t rc = fm_workbook_add_num_fmt(handle_, format_code.c_str(), &id);
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
    r.status = invalid_handle_status();
    return r;
  }
  fm_cell_xf_t xf{};
  JsNarrowNumericReader reader("addCellXf");
  js_pull_cell_xf(record, &xf, reader);
  if (!reader.ok()) {
    r.status = reader.failure();
    return r;
  }
  std::size_t idx = 0;
  const fm_status_t rc = fm_workbook_add_cell_xf(handle_, &xf, &idx);
  return index_result(rc, idx);
}

JsAddStyleResult JsWorkbook::addDxf(emscripten::val record) {
  JsAddStyleResult r;
  if (handle_ == nullptr) {
    r.status = invalid_handle_status();
    return r;
  }

  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  std::vector<fm_gradient_stop_t> fill_stops;
  fm_dxf_record_t dxf{};
  JsNarrowNumericReader reader("addDxf");

  emscripten::val font = reader.value(record, "font", "dxf.font");
  if (js_value_present(font)) {
    dxf.font_engaged = 1;
    js_pull_font_record(font, &font_name, &dxf.font, reader);
  }

  emscripten::val fill = reader.value(record, "fill", "dxf.fill");
  if (js_value_present(fill)) {
    dxf.fill_engaged = 1;
    dxf.fill = js_pull_fill_record(fill, &fill_stops, reader);
  }

  emscripten::val border = reader.value(record, "border", "dxf.border");
  if (js_value_present(border)) {
    dxf.border_engaged = 1;
    dxf.border = js_pull_border_record(border, kDxfBorderFields, reader);
  }

  emscripten::val num_fmt = reader.value(record, "numFmt", "dxf.numFmt");
  if (js_value_present(num_fmt)) {
    dxf.num_fmt_engaged = 1;
    dxf.num_fmt_id = reader.u16(num_fmt, "numFmtId", 0U, "dxf.numFmtId");
    num_fmt_code = reader.string(num_fmt, "formatCode", "dxf.numFmt.formatCode");
    dxf.num_fmt_code = num_fmt_code.c_str();
  }

  alignment_xml = reader.string(record, "alignmentXml", "dxf.alignmentXml");
  protection_xml = reader.string(record, "protectionXml", "dxf.protectionXml");
  dxf.alignment_xml = alignment_xml.c_str();
  dxf.protection_xml = protection_xml.c_str();

  if (!reader.ok()) {
    r.status = reader.failure();
    return r;
  }

  std::size_t idx = 0;
  const fm_status_t rc = fm_workbook_add_dxf(handle_, &dxf, &idx);
  return index_result(rc, idx);
}

// ---- Style count accessors ---------------------------------------------
//
// `fontCount` / `fillCount` / `borderCount` / `xfCount` are now emitted
// by the binding codegen (see `src/wasm/generated/styles_counts.cpp`).

JsNumberResult JsWorkbook::dxfCount() const {
  std::size_t n = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_dxf_count(handle_, &n) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(n));
}
// `cellStyleXfCount` stays here because it has no N-API counterpart and is
// therefore not part of the cross-binding manifest.

JsNumberResult JsWorkbook::cellStyleXfCount() const {
  std::size_t n = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_cell_style_xf_count(handle_, &n) : kBindingInvalidHandle;
  return number_result(rc, static_cast<double>(n));
}

namespace {

emscripten::val cell_style_to_val(const fm_cell_style_record_t& cs) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", cs.name);
  o.set("xfId", cs.xf_id);
  o.set("builtinId", cs.builtin_id);
  o.set("iLevel", cs.i_level);
  o.set("hidden", cs.hidden != 0);
  o.set("customBuiltin", cs.custom_builtin != 0);
  return o;
}

emscripten::val table_style_to_val(const fm_table_style_t& style) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", style.name);
  o.set("pivot", style.pivot != 0);
  o.set("table", style.table != 0);
  emscripten::val elements = emscripten::val::array();
  for (uint32_t i = 0; i < style.element_count; ++i) {
    const fm_table_style_element_t& element = style.elements[i];
    emscripten::val e = emscripten::val::object();
    js_set_cstr(e, "type", element.type);
    e.set("size", element.size);
    if (element.has_dxf_id != 0) {
      e.set("dxfId", element.dxf_id);
    }
    elements.call<void>("push", e);
  }
  o.set("elements", elements);
  return o;
}

// Style tables are workbook-scoped; the sheet argument is ignored.
fm_status_t count_cell_styles(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_cell_style_count(handle, out);
}

fm_status_t cell_style_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  fm_cell_style_record_t cs{};
  const fm_status_t rc = fm_workbook_get_cell_style_at(handle, index, &cs);
  if (rc == 0) {
    *out = cell_style_to_val(cs);
  }
  return rc;
}

fm_status_t count_table_styles(const fm_workbook_t* handle, size_t, size_t* out) {
  return fm_workbook_get_table_style_count(handle, out);
}

fm_status_t table_style_item(const fm_workbook_t* handle, size_t, size_t index, emscripten::val* out) {
  fm_table_style_t style{};
  const fm_status_t rc = fm_workbook_get_table_style_at(handle, index, &style);
  if (rc == 0) {
    *out = table_style_to_val(style);
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listCellStyles() const {
  emscripten::val arr = emscripten::val::array();
  arr.set("status", status_from_rc(js_collect_list(handle_, 0U, &count_cell_styles, &cell_style_item, arr)));
  return arr;
}

emscripten::val JsWorkbook::getCellStyle(const std::string& name) const {
  fm_cell_style_record_t cs{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_style(handle_, name.c_str(), &present, &cs) : kBindingInvalidHandle;
  return optional_result(rc, present, rc == 0 && present != 0 ? cell_style_to_val(cs) : emscripten::val::null());
}

emscripten::val JsWorkbook::getCellStyleXf(uint32_t index) const {
  fm_cell_xf_t xf{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_style_xf(handle_, index, &xf) : kBindingInvalidHandle;
  return js_cell_xf_result(rc, xf);
}

JsAddStyleResult JsWorkbook::addCellStyleXf(emscripten::val record) {
  JsAddStyleResult out;
  if (handle_ == nullptr) {
    out.status = invalid_handle_status();
    return out;
  }
  fm_cell_xf_t xf{};
  JsNarrowNumericReader reader("addCellStyleXf");
  js_pull_cell_xf(record, &xf, reader);
  if (!reader.ok()) {
    out.status = reader.failure();
    return out;
  }
  std::size_t index = 0;
  const fm_status_t rc = fm_workbook_add_cell_style_xf(handle_, &xf, &index);
  return index_result(rc, index);
}

JsStatus JsWorkbook::setCellStyle(emscripten::val record) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setCellStyle");
  const std::string name = reader.string(record, "name", "cellStyle.name");
  fm_cell_style_record_t cs{};
  cs.name = name.c_str();
  cs.xf_id = reader.u32(record, "xfId", 0U, "cellStyle.xfId");
  cs.builtin_id = reader.u32(record, "builtinId", FM_CELL_STYLE_BUILTIN_ID_NONE, "cellStyle.builtinId");
  cs.i_level = reader.u32(record, "iLevel", 0U, "cellStyle.iLevel");
  cs.hidden = reader.boolean(record, "hidden", false, "cellStyle.hidden") ? 1 : 0;
  cs.custom_builtin = reader.boolean(record, "customBuiltin", false, "cellStyle.customBuiltin") ? 1 : 0;
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_workbook_set_cell_style(handle_, &cs));
}

JsStatus JsWorkbook::removeCellStyle(const std::string& name) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_remove_cell_style(handle_, name.c_str())
                                           : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::listTableStyles() const {
  emscripten::val arr = emscripten::val::array();
  arr.set("status", status_from_rc(js_collect_list(handle_, 0U, &count_table_styles, &table_style_item, arr)));
  return arr;
}

emscripten::val JsWorkbook::getTableStyle(const std::string& name) const {
  fm_table_style_t style{};
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_table_style(handle_, name.c_str(), &present, &style) : kBindingInvalidHandle;
  return optional_result(rc, present, rc == 0 && present != 0 ? table_style_to_val(style) : emscripten::val::null());
}

JsStatus JsWorkbook::setTableStyle(emscripten::val style) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setTableStyle");
  const std::string name = reader.string(style, "name", "style.name");
  const bool pivot = reader.boolean(style, "pivot", true, "style.pivot");
  const bool table = reader.boolean(style, "table", true, "style.table");
  const emscripten::val elements = reader.value(style, "elements", "style.elements");
  std::vector<std::string> types;
  std::vector<fm_table_style_element_t> records;
  if (reader.ok() && js_value_present(elements)) {
    if (!reader.is_array(elements, "style.elements")) {
      return binding_error_status(kInvalidArgument, "setTableStyle: `elements` must be an array");
    }
    const uint32_t count = reader.length(elements, "style.elements.length");
    types.reserve(count);
    records.reserve(count);
    for (uint32_t i = 0; i < count && reader.ok(); ++i) {
      const emscripten::val element = reader.array_element(elements, i, "style.elements[]");
      types.push_back(reader.string(element, "type", "style.elements[].type"));
      fm_table_style_element_t record{};
      record.size = reader.u32(element, "size", 1U, "style.elements[].size");
      const emscripten::val dxf = reader.value(element, "dxfId", "style.elements[].dxfId");
      if (reader.ok() && js_value_present(dxf)) {
        record.has_dxf_id = 1;
        record.dxf_id = reader.u32_value(dxf, 0U, "style.elements[].dxfId");
      }
      records.push_back(record);
    }
  }
  if (!reader.ok()) {
    return reader.failure();
  }
  // `types` has stopped growing, so its `c_str()` views stay put.
  for (std::size_t i = 0; i < records.size(); ++i) {
    records[i].type = types[i].c_str();
  }
  fm_table_style_t record{};
  record.name = name.c_str();
  record.pivot = pivot ? 1 : 0;
  record.table = table ? 1 : 0;
  record.elements = records.empty() ? nullptr : records.data();
  record.element_count = static_cast<uint32_t>(records.size());
  return status_from_rc(fm_workbook_set_table_style(handle_, &record));
}

JsStatus JsWorkbook::removeTableStyle(const std::string& name) {
  return status_from_rc(handle_ != nullptr ? fm_workbook_remove_table_style(handle_, name.c_str())
                                           : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::getTableStyleDefaults() const {
  emscripten::val o = emscripten::val::object();
  const char* table = nullptr;
  const char* pivot = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_table_style_defaults(handle_, &table, &pivot) : kBindingInvalidHandle;
  o.set("status", status_from_rc(rc));
  js_set_cstr(o, "tableStyle", rc == 0 ? table : nullptr);
  js_set_cstr(o, "pivotStyle", rc == 0 ? pivot : nullptr);
  return o;
}

JsStatus JsWorkbook::setTableStyleDefaults(emscripten::val defaults) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setTableStyleDefaults");
  std::string table;
  std::string pivot;
  const char* table_ptr = reader.optional_string(defaults, "tableStyle", table, "defaults.tableStyle");
  const char* pivot_ptr = reader.optional_string(defaults, "pivotStyle", pivot, "defaults.pivotStyle");
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_workbook_set_table_style_defaults(handle_, table_ptr, pivot_ptr));
}

namespace {

emscripten::val js_resolved_color(uint32_t argb, int32_t resolution) {
  emscripten::val o = emscripten::val::object();
  o.set("argb", argb);
  o.set("resolution", resolution);
  return o;
}

/// Reads the major or minor script-font list as `{script, typeface}` records.
/// Each pair is copied before the next `_at` call reuses the read scratch.
emscripten::val theme_script_fonts(const fm_workbook_t* handle, bool major) {
  emscripten::val list = emscripten::val::array();
  size_t count = 0;
  if (handle == nullptr || fm_workbook_get_theme_script_font_count(handle, major ? 1 : 0, &count) != 0) {
    return list;
  }
  for (size_t i = 0; i < count; ++i) {
    const char* script = nullptr;
    const char* typeface = nullptr;
    if (fm_workbook_get_theme_script_font_at(handle, major ? 1 : 0, i, &script, &typeface) != 0) {
      break;
    }
    emscripten::val item = emscripten::val::object();
    js_set_cstr_fields(item, {{"script", script}, {"typeface", typeface}});
    list.set(static_cast<uint32_t>(i), item);
  }
  return list;
}

}  // namespace

emscripten::val JsWorkbook::getThemeScriptFont(bool major, const std::string& script) const {
  const char* typeface = nullptr;
  int32_t present = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_theme_script_font(handle_, major ? 1 : 0, script.c_str(),
                                                                                &present, &typeface)
                                            : kBindingInvalidHandle;
  emscripten::val item = emscripten::val::null();
  if (rc == 0 && present != 0) {
    item = emscripten::val::object();
    js_set_cstr_fields(item, {{"script", script.c_str()}, {"typeface", typeface}});
  }
  return optional_result(rc, present, item);
}

emscripten::val JsWorkbook::getTheme() const {
  fm_theme_colors_t colors{};
  fm_theme_fonts_t fonts{};
  int32_t source = 0;
  fm_status_t rc = handle_ != nullptr ? fm_workbook_get_theme_colors(handle_, &colors, &source) : kBindingInvalidHandle;
  if (rc == 0) {
    rc = fm_workbook_get_theme_fonts(handle_, &fonts);
  }
  if (rc != 0) {
    colors = fm_theme_colors_t{};
    fonts = fm_theme_fonts_t{};
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
  o.set("majorScripts", theme_script_fonts(handle_, true));
  o.set("minorScripts", theme_script_fonts(handle_, false));
  return o;
}

JsStatus JsWorkbook::setThemeColors(emscripten::val colors) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  fm_theme_colors_t tc{};
  JsNarrowNumericReader reader("setThemeColors");
  if (!reader.is_array(colors, "colors") || reader.length(colors, "colors") != std::size(tc.argb)) {
    return binding_error_status(kInvalidArgument, "setThemeColors: `colors` must be an array of 12 ARGB numbers");
  }
  for (uint32_t i = 0; i < std::size(tc.argb); ++i) {
    tc.argb[i] = reader.u32_value(reader.array_element(colors, i, "colors[]"), 0U, "colors[]");
  }
  if (!reader.ok()) {
    return reader.failure();
  }
  return status_from_rc(fm_workbook_set_theme_colors(handle_, &tc));
}

JsStatus JsWorkbook::setThemeFonts(emscripten::val fonts) {
  if (handle_ == nullptr) {
    return invalid_handle_status();
  }
  JsNarrowNumericReader reader("setThemeFonts");
  const std::string major_latin = reader.string(fonts, "majorLatin", "theme.majorLatin");
  const std::string major_ea = reader.string(fonts, "majorEastAsian", "theme.majorEastAsian");
  const std::string minor_latin = reader.string(fonts, "minorLatin", "theme.minorLatin");
  const std::string minor_ea = reader.string(fonts, "minorEastAsian", "theme.minorEastAsian");
  if (!reader.ok()) {
    return reader.failure();
  }
  fm_theme_fonts_t tf{};
  tf.major_latin = major_latin.c_str();
  tf.major_east_asian = major_ea.c_str();
  tf.minor_latin = minor_latin.c_str();
  tf.minor_east_asian = minor_ea.c_str();
  return status_from_rc(fm_workbook_set_theme_fonts(handle_, &tf));
}

JsStatus JsWorkbook::setThemeScriptFont(bool major, const std::string& script, const std::string& typeface) {
  return status_from_rc(
      handle_ != nullptr ? fm_workbook_set_theme_script_font(handle_, major ? 1 : 0, script.c_str(), typeface.c_str())
                         : kBindingInvalidHandle);
}

JsStatus JsWorkbook::resetTheme() {
  return status_from_rc(handle_ != nullptr ? fm_workbook_clear_theme(handle_) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::resolveColor(emscripten::val spec, int32_t context) const {
  uint32_t argb = 0;
  int32_t resolution = 0;
  emscripten::val o = emscripten::val::object();
  if (handle_ == nullptr) {
    o.set("status", invalid_handle_status());
    o.set("argb", 0U);
    o.set("resolution", 0);
    return o;
  }
  emscripten::val holder = emscripten::val::object();
  holder.set("c", spec);
  JsNarrowNumericReader reader("resolveColor");
  const fm_color_spec_t color = js_pull_color_spec(holder, "c", &reader);
  if (!reader.ok()) {
    o.set("status", reader.failure());
    o.set("argb", 0U);
    o.set("resolution", 0);
    return o;
  }
  const fm_status_t rc = fm_workbook_resolve_color(handle_, &color, context, &argb, &resolution);
  o.set("status", status_from_rc(rc));
  o.set("argb", rc == 0 ? argb : 0U);
  o.set("resolution", rc == 0 ? resolution : 0);
  return o;
}

emscripten::val JsWorkbook::getEffectiveStyle(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_effective_style_t e{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_effective_style(handle_, sheet, row, col, &e) : kBindingInvalidHandle;
  if (rc != 0) {
    e = fm_effective_style_t{};
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
