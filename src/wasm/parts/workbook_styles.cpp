//
// JsWorkbook styles surface: per-cell xfIndex get/set, font / fill /
// border / numFmt / xf record getters and adders, cellStyleXfs accessors,
// plus the matching count accessors. The add* paths share the `js_pull_*`
// helpers in `parts/embind_common.h` so the embind glue is emitted once
// per field type rather than per call site.

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
  xf->reading_order = reader.i32_value(reading_order, 0, "xf.readingOrder");
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
/// `<dxf>` font projection so both surface the same field set. An absent
/// element reads as `null`.
emscripten::val js_font_record(const fm_font_record_t& f) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", f.name);
  js_set_present_or_null(o, "size", f.has_size, f.size);
  js_set_present_or_null(o, "bold", f.has_bold, f.bold != 0);
  js_set_present_or_null(o, "italic", f.has_italic, f.italic != 0);
  js_set_present_or_null(o, "strike", f.has_strike, f.strike != 0);
  o.set("underline", static_cast<std::uint32_t>(f.underline));
  o.set("vertAlign", static_cast<std::uint32_t>(f.vert_align));
  js_set_present_or_null(o, "family", f.has_family, static_cast<std::uint32_t>(f.family));
  js_set_present_or_null(o, "charset", f.has_charset, static_cast<std::uint32_t>(f.charset));
  o.set("scheme", static_cast<std::uint32_t>(f.scheme));
  js_set_present_or_null(o, "outline", f.has_outline, f.outline != 0);
  js_set_present_or_null(o, "shadow", f.has_shadow, f.shadow != 0);
  js_set_present_or_null(o, "condense", f.has_condense, f.condense != 0);
  js_set_present_or_null(o, "extend", f.has_extend, f.extend != 0);
  o.set("color", js_color_spec(f.color));
  return o;
}

/// Reads a font record out of a JS object. `name_storage` owns the font
/// name for the duration of the C ABI call, which borrows the pointer. An
/// omitted or `null` element is absent from the record.
void js_pull_font_record(const emscripten::val& record, std::string* name_storage, fm_font_record_t* out,
                         JsNarrowNumericReader& reader) {
  const auto flag = [&](const char* key, const char* field, int32_t* has, int32_t* value) {
    const emscripten::val v = reader.value(record, key, field);
    *has = js_value_present(v) ? 1 : 0;
    *value = reader.boolean_value(v, false) ? 1 : 0;
  };
  *name_storage = reader.string(record, "name", "font.name");
  out->name = name_storage->c_str();
  const emscripten::val size = reader.value(record, "size", "font.size");
  out->has_size = js_value_present(size) ? 1 : 0;
  out->size = reader.number_value(size, 11.0, "font.size");
  flag("bold", "font.bold", &out->has_bold, &out->bold);
  flag("italic", "font.italic", &out->has_italic, &out->italic);
  flag("strike", "font.strike", &out->has_strike, &out->strike);
  out->underline = reader.u8(record, "underline", 0U, "font.underline");
  out->vert_align = reader.u8(record, "vertAlign", 0U, "font.vertAlign");
  const emscripten::val family = reader.value(record, "family", "font.family");
  out->has_family = js_value_present(family) ? 1 : 0;
  out->family = reader.u8_value(family, 0U, "font.family");
  const emscripten::val charset = reader.value(record, "charset", "font.charset");
  out->has_charset = js_value_present(charset) ? 1 : 0;
  out->charset = reader.u8_value(charset, 0U, "font.charset");
  out->scheme = reader.u8(record, "scheme", 0U, "font.scheme");
  flag("outline", "font.outline", &out->has_outline, &out->outline);
  flag("shadow", "font.shadow", &out->has_shadow, &out->shadow);
  flag("condense", "font.condense", &out->has_condense, &out->condense);
  flag("extend", "font.extend", &out->has_extend, &out->extend);
  out->color = js_pull_color_spec(record, "color", &reader);
}

emscripten::val js_fill_record(const fm_fill_record_t& f) {
  emscripten::val o = emscripten::val::object();
  o.set("pattern", static_cast<std::uint32_t>(f.pattern));
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
/// `getCellStyleXf`. `xfId` is set by the callers: a named-style xf is its
/// own parent, so `getCellStyleXf` always reports 0 there.
emscripten::val js_cell_xf_record(const fm_cell_xf_t& xf) {
  emscripten::val o = emscripten::val::object();
  o.set("fontIndex", xf.font_index);
  o.set("fillIndex", xf.fill_index);
  o.set("borderIndex", xf.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.num_fmt_id));
  js_set_present_or_null(o, "horizontalAlign", xf.has_horizontal_align, static_cast<uint32_t>(xf.horizontal_align));
  js_set_present_or_null(o, "verticalAlign", xf.has_vertical_align, static_cast<uint32_t>(xf.vertical_align));
  js_set_present_or_null(o, "wrapText", xf.has_wrap_text, xf.wrap_text != 0);
  js_set_present_or_null(o, "justifyLastLine", xf.has_justify_last_line, xf.justify_last_line != 0);
  o.set("applyNumberFormat", xf.apply_number_format != 0);
  o.set("applyFont", xf.apply_font != 0);
  o.set("applyFill", xf.apply_fill != 0);
  o.set("applyBorder", xf.apply_border != 0);
  o.set("applyAlignment", xf.apply_alignment != 0);
  o.set("applyProtection", xf.apply_protection != 0);
  o.set("quotePrefix", xf.quote_prefix != 0);
  // Without a `<protection>` child both read `null`: the model default
  // (locked, not hidden) applies.
  js_set_present_or_null(o, "locked", xf.has_protection, xf.locked != 0);
  js_set_present_or_null(o, "hidden", xf.has_protection, xf.hidden != 0);
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

// Builds the `getCellXf` / `getCellStyleXf` record from the C call's status
// and xf; `undefined` after recording the failure.
emscripten::val js_cell_xf_result(fm_status_t rc, const fm_cell_xf_t& xf) {
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = js_cell_xf_record(xf);
  o.set("xfId", xf.xf_id);
  js_set_cell_xf_alignment(xf, &o);
  return o;
}

}  // namespace

// ---- Per-cell xf index get/set -----------------------------------------

double JsWorkbook::getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
  std::size_t xf = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_xf_index(handle_, sheet, row, col, &xf)
                                    : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(xf);
}

void JsWorkbook::setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
  succeeded(handle_ != nullptr ? fm_sheet_set_cell_xf_index(handle_, sheet, row, col, xf_index)
                               : kBindingInvalidHandle);
}

void JsWorkbook::clearCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_cell_xf_index(handle_, sheet, row, col) : kBindingInvalidHandle);
}

void JsWorkbook::setRangeXfIndex(uint32_t sheet, emscripten::val range, uint32_t xf_index) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setRangeXfIndex");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_range_xf_index(handle_, sheet, &rect, xf_index));
}

// ---- Style record getters ----------------------------------------------

emscripten::val JsWorkbook::getCellXf(uint32_t xf_index) const {
  fm_cell_xf_t xf{};
  const fm_status_t rc = handle_ != nullptr ? fm_workbook_get_cell_xf(handle_, xf_index, &xf) : kBindingInvalidHandle;
  return js_cell_xf_result(rc, xf);
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  fm_font_record_t f{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_font(handle_, font_index, &f) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_font_record(f);
}

emscripten::val JsWorkbook::getFill(uint32_t fill_index) const {
  fm_fill_record_t f{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_fill(handle_, fill_index, &f) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_fill_record(f);
}

emscripten::val JsWorkbook::getBorder(uint32_t border_index) const {
  fm_border_record_t b{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_border(handle_, border_index, &b) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_border_record(b);
}

emscripten::val JsWorkbook::getNumFmt(uint32_t num_fmt_id) const {
  int32_t present = 0;
  const char* s = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_num_fmt_string(handle_, num_fmt_id, &present, &s) : kBindingInvalidHandle;
  return optional_or_null(rc, present,
                          rc == 0 && present != 0 ? emscripten::val(std::string(s)) : emscripten::val::null());
}

emscripten::val JsWorkbook::getDxf(uint32_t dxf_index) const {
  fm_dxf_record_t d{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_dxf(handle_, dxf_index, &d) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
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

template <typename Pull, typename Add>
uint32_t JsWorkbook::add_style_record(const char* op, const emscripten::val& record, Pull&& pull, Add&& add) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader(op);
  pull(reader, record);
  if (!reader.ok()) {
    return 0U;
  }
  std::size_t idx = 0;
  if (!succeeded(add(handle_, &idx))) {
    return 0U;
  }
  return static_cast<uint32_t>(idx);
}

template <typename Pull, typename Apply>
void JsWorkbook::set_style_record(const char* op, const emscripten::val& record, Pull&& pull, Apply&& apply) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader(op);
  pull(reader, record);
  if (!reader.ok()) {
    return;
  }
  succeeded(apply(handle_));
}

uint32_t JsWorkbook::addFont(emscripten::val record) {
  std::string name;
  fm_font_record_t fr{};
  return add_style_record(
      "addFont", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_font_record(rec, &name, &fr, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_font(handle, &fr, idx); });
}

void JsWorkbook::setFont(uint32_t font_index, emscripten::val record) {
  std::string name;
  fm_font_record_t fr{};
  set_style_record(
      "setFont", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_font_record(rec, &name, &fr, reader); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_font(handle, font_index, &fr); });
}

void JsWorkbook::setDefaultFont(emscripten::val record) {
  std::string name;
  fm_font_record_t fr{};
  set_style_record(
      "setDefaultFont", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_font_record(rec, &name, &fr, reader); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_default_font(handle, &fr); });
}

uint32_t JsWorkbook::addFill(emscripten::val record) {
  std::vector<fm_gradient_stop_t> stops;
  fm_fill_record_t fr{};
  return add_style_record(
      "addFill", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { fr = js_pull_fill_record(rec, &stops, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_fill(handle, &fr, idx); });
}

uint32_t JsWorkbook::addBorder(emscripten::val record) {
  fm_border_record_t br{};
  return add_style_record(
      "addBorder", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) {
        br = js_pull_border_record(rec, kBorderFields, reader);
      },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_border(handle, &br, idx); });
}

uint32_t JsWorkbook::addNumFmt(const std::string& format_code) {
  uint32_t id = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_num_fmt(handle_, format_code.c_str(), &id)
                                    : kBindingInvalidHandle)) {
    return 0U;
  }
  return id;
}

uint32_t JsWorkbook::addXf(emscripten::val record) {
  fm_cell_xf_t xf{};
  return add_style_record(
      "addCellXf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_cell_xf(rec, &xf, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_cell_xf(handle, &xf, idx); });
}

uint32_t JsWorkbook::addDxf(emscripten::val record) {
  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  std::vector<fm_gradient_stop_t> fill_stops;
  fm_dxf_record_t dxf{};
  return add_style_record(
      "addDxf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) {
        emscripten::val font = reader.value(rec, "font", "dxf.font");
        if (js_value_present(font)) {
          dxf.font_engaged = 1;
          js_pull_font_record(font, &font_name, &dxf.font, reader);
        }

        emscripten::val fill = reader.value(rec, "fill", "dxf.fill");
        if (js_value_present(fill)) {
          dxf.fill_engaged = 1;
          dxf.fill = js_pull_fill_record(fill, &fill_stops, reader);
        }

        emscripten::val border = reader.value(rec, "border", "dxf.border");
        if (js_value_present(border)) {
          dxf.border_engaged = 1;
          dxf.border = js_pull_border_record(border, kDxfBorderFields, reader);
        }

        emscripten::val num_fmt = reader.value(rec, "numFmt", "dxf.numFmt");
        if (js_value_present(num_fmt)) {
          dxf.num_fmt_engaged = 1;
          dxf.num_fmt_id = reader.u16(num_fmt, "numFmtId", 0U, "dxf.numFmtId");
          num_fmt_code = reader.string(num_fmt, "formatCode", "dxf.numFmt.formatCode");
          dxf.num_fmt_code = num_fmt_code.c_str();
        }

        alignment_xml = reader.string(rec, "alignmentXml", "dxf.alignmentXml");
        protection_xml = reader.string(rec, "protectionXml", "dxf.protectionXml");
        dxf.alignment_xml = alignment_xml.c_str();
        dxf.protection_xml = protection_xml.c_str();
      },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_dxf(handle, &dxf, idx); });
}

// ---- Style count accessors ---------------------------------------------
//
// `fontCount` / `fillCount` / `borderCount` / `xfCount` are now emitted
// by the binding codegen (see `src/wasm/generated/styles_counts.cpp`).

double JsWorkbook::dxfCount() const {
  std::size_t n = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_dxf_count(handle_, &n) : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(n);
}
// `cellStyleXfCount` stays here because it has no N-API counterpart and is
// therefore not part of the cross-binding manifest.

double JsWorkbook::cellStyleXfCount() const {
  std::size_t n = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_cell_style_xf_count(handle_, &n) : kBindingInvalidHandle)) {
    return 0.0;
  }
  return static_cast<double>(n);
}

emscripten::val JsWorkbook::getCellStyleXf(uint32_t index) const {
  fm_cell_xf_t xf{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_style_xf(handle_, index, &xf) : kBindingInvalidHandle;
  return js_cell_xf_result(rc, xf);
}

uint32_t JsWorkbook::addCellStyleXf(emscripten::val record) {
  fm_cell_xf_t xf{};
  return add_style_record(
      "addCellStyleXf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_cell_xf(rec, &xf, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_cell_style_xf(handle, &xf, idx); });
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
