//
// JsWorkbook styles surface: per-cell xfIndex get/set, font / fill /
// border / numFmt / xf record getters and adders, cellStyleXfs accessors,
// plus the matching count accessors. The add* paths share the `js_pull_*`
// helpers in `parts/embind_common.h` so the embind glue is emitted once
// per field type rather than per call site.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <cstdlib>
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

// ---- Differential-format alignment / protection <-> XML -----------------
//
// The C record keeps `<alignment/>` and `<protection/>` as serialized XML;
// the binding projects them to the `Alignment` / `CellProtection` records.

constexpr const char* kHorizontalNames[] = {"general", "left",    "center",           "right",
                                            "fill",    "justify", "centerContinuous", "distributed"};
constexpr const char* kVerticalNames[] = {"top", "center", "bottom", "justify", "distributed"};

/// The value of attribute `name` in the element text `xml`; false when absent.
bool xml_attribute(const std::string& xml, const std::string& name, std::string* value) {
  const std::size_t head = xml.find(' ' + name + '=');
  if (head == std::string::npos) {
    return false;
  }
  const std::size_t quote = head + name.size() + 2U;
  if (quote >= xml.size() || (xml[quote] != '"' && xml[quote] != '\'')) {
    return false;
  }
  const std::size_t end = xml.find(xml[quote], quote + 1U);
  if (end == std::string::npos) {
    return false;
  }
  *value = xml.substr(quote + 1U, end - quote - 1U);
  return true;
}

int32_t xml_int(const std::string& xml, const char* name, int32_t dflt) {
  std::string text;
  if (!xml_attribute(xml, name, &text) || text.empty()) {
    return dflt;
  }
  char* end = nullptr;
  const long parsed = std::strtol(text.c_str(), &end, 10);
  return end != text.c_str() ? static_cast<int32_t>(parsed) : dflt;
}

bool xml_flag(const std::string& xml, const char* name, bool dflt) {
  std::string text;
  if (!xml_attribute(xml, name, &text)) {
    return dflt;
  }
  return text == "1" || text == "true";
}

template <std::size_t N>
uint32_t xml_name_ordinal(const std::string& xml, const char* name, const char* const (&names)[N], uint32_t dflt) {
  std::string text;
  if (!xml_attribute(xml, name, &text)) {
    return dflt;
  }
  for (std::size_t i = 0; i < N; ++i) {
    if (text == names[i]) {
      return static_cast<uint32_t>(i);
    }
  }
  return dflt;
}

emscripten::val dxf_alignment_to_val(const std::string& xml) {
  emscripten::val o = emscripten::val::object();
  o.set("horizontal", xml_name_ordinal(xml, "horizontal", kHorizontalNames, 0U));
  o.set("vertical", xml_name_ordinal(xml, "vertical", kVerticalNames, 2U));
  o.set("wrapText", xml_flag(xml, "wrapText", false));
  o.set("shrinkToFit", xml_flag(xml, "shrinkToFit", false));
  o.set("indent", xml_int(xml, "indent", 0));
  o.set("textRotation", xml_int(xml, "textRotation", 0));
  o.set("readingOrder", xml_int(xml, "readingOrder", 0));
  o.set("justifyLastLine", xml_flag(xml, "justifyLastLine", false));
  o.set("relativeIndent", xml_int(xml, "relativeIndent", 0));
  return o;
}

emscripten::val dxf_protection_to_val(const std::string& xml) {
  emscripten::val o = emscripten::val::object();
  o.set("locked", xml_flag(xml, "locked", true));
  o.set("hidden", xml_flag(xml, "hidden", false));
  return o;
}

/// `<alignment .../>` for an `Alignment` record; attributes at their default are left out.
std::string dxf_alignment_xml(JsNarrowNumericReader& reader, const emscripten::val& a) {
  const uint8_t horizontal = reader.u8(a, "horizontal", 0U, "dxf.alignment.horizontal");
  const uint8_t vertical = reader.u8(a, "vertical", 2U, "dxf.alignment.vertical");
  if (horizontal >= std::size(kHorizontalNames)) {
    reader.fail_range("dxf.alignment.horizontal", "[0, 7]");
  }
  if (vertical >= std::size(kVerticalNames)) {
    reader.fail_range("dxf.alignment.vertical", "[0, 4]");
  }
  const bool wrap = reader.boolean(a, "wrapText", false, "dxf.alignment.wrapText");
  const bool shrink = reader.boolean(a, "shrinkToFit", false, "dxf.alignment.shrinkToFit");
  const bool justify = reader.boolean(a, "justifyLastLine", false, "dxf.alignment.justifyLastLine");
  const int32_t indent = reader.i32(a, "indent", 0, "dxf.alignment.indent");
  const int32_t rotation = reader.i32(a, "textRotation", 0, "dxf.alignment.textRotation");
  const int32_t order = reader.i32(a, "readingOrder", 0, "dxf.alignment.readingOrder");
  const int32_t relative = reader.i32(a, "relativeIndent", 0, "dxf.alignment.relativeIndent");
  if (!reader.ok()) {
    return std::string();
  }
  std::string xml = "<alignment";
  const auto text = [&](const char* name, const std::string& value) {
    xml += ' ' + std::string(name) + "=\"" + value + '"';
  };
  if (horizontal != 0U) {
    text("horizontal", kHorizontalNames[horizontal]);
  }
  if (vertical != 2U) {
    text("vertical", kVerticalNames[vertical]);
  }
  if (rotation != 0) {
    text("textRotation", std::to_string(rotation));
  }
  if (wrap) {
    text("wrapText", "1");
  }
  if (justify) {
    text("justifyLastLine", "1");
  }
  if (shrink) {
    text("shrinkToFit", "1");
  }
  if (indent != 0) {
    text("indent", std::to_string(indent));
  }
  if (relative != 0) {
    text("relativeIndent", std::to_string(relative));
  }
  if (order != 0) {
    text("readingOrder", std::to_string(order));
  }
  return xml + "/>";
}

/// `<protection .../>` for a `CellProtection` record.
std::string dxf_protection_xml(JsNarrowNumericReader& reader, const emscripten::val& p) {
  const bool locked = reader.boolean(p, "locked", true, "dxf.protection.locked");
  const bool hidden = reader.boolean(p, "hidden", false, "dxf.protection.hidden");
  std::string xml = "<protection";
  if (!locked) {
    xml += " locked=\"0\"";
  }
  if (hidden) {
    xml += " hidden=\"1\"";
  }
  return xml + "/>";
}

void js_pull_cell_xf(const emscripten::val& record, fm_cell_xf_t* xf, JsNarrowNumericReader& reader) {
  xf->font_index = reader.u32(record, "fontIndex", 0U, "xf.fontIndex");
  xf->fill_index = reader.u32(record, "fillIndex", 0U, "xf.fillIndex");
  xf->border_index = reader.u32(record, "borderIndex", 0U, "xf.borderIndex");
  xf->num_fmt_id = reader.u16(record, "numFmtId", 0U, "xf.numFmtId");
  const emscripten::val horizontal_align = reader.value(record, "horizontal", "xf.horizontal");
  const emscripten::val vertical_align = reader.value(record, "vertical", "xf.vertical");
  const emscripten::val wrap_text = reader.value(record, "wrapText", "xf.wrapText");
  const emscripten::val justify_last_line = reader.value(record, "justifyLastLine", "xf.justifyLastLine");
  const bool horizontal_align_present = js_value_present(horizontal_align);
  const bool vertical_align_present = js_value_present(vertical_align);
  const bool wrap_text_present = js_value_present(wrap_text);
  const bool justify_last_line_present = js_value_present(justify_last_line);
  xf->horizontal = reader.u8_value(horizontal_align, 0U, "xf.horizontal");
  xf->vertical = reader.u8_value(vertical_align, 2U, "xf.vertical");
  xf->wrap_text = reader.boolean_value(wrap_text, false) ? 1 : 0;
  xf->justify_last_line = reader.boolean_value(justify_last_line, false) ? 1 : 0;
  xf->style_xf_index = reader.u32(record, "styleXfIndex", 0U, "xf.styleXfIndex");
  xf->has_horizontal = horizontal_align_present ? 1 : 0;
  xf->has_vertical = vertical_align_present ? 1 : 0;
  xf->has_wrap_text = wrap_text_present ? 1 : 0;
  xf->has_justify_last_line = justify_last_line_present ? 1 : 0;
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
                                      relative_indent_present || shrink_to_fit_present || reading_order_present;
  xf->has_alignment = has_supplied_alignment ? 1 : 0;

  xf->apply_number_format = reader.boolean(record, "applyNumberFormat", false, "xf.applyNumberFormat") ? 1 : 0;
  xf->apply_font = reader.boolean(record, "applyFont", false, "xf.applyFont") ? 1 : 0;
  xf->apply_fill = reader.boolean(record, "applyFill", false, "xf.applyFill") ? 1 : 0;
  xf->apply_border = reader.boolean(record, "applyBorder", false, "xf.applyBorder") ? 1 : 0;
  xf->apply_alignment = reader.boolean(record, "applyAlignment", false, "xf.applyAlignment") ? 1 : 0;
  xf->apply_protection = reader.boolean(record, "applyProtection", false, "xf.applyProtection") ? 1 : 0;
  xf->quote_prefix = reader.boolean(record, "quotePrefix", false, "xf.quotePrefix") ? 1 : 0;
  const emscripten::val locked = reader.value(record, "locked", "xf.locked");
  const emscripten::val hidden = reader.value(record, "hidden", "xf.hidden");
  const bool locked_present = js_value_present(locked);
  const bool hidden_present = js_value_present(hidden);
  xf->has_protection = (locked_present || hidden_present) ? 1 : 0;
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

fm_border_t js_pull_border_record(const emscripten::val& record, const BorderFieldNames& names,
                                  JsNarrowNumericReader& reader) {
  fm_border_t br{};
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
emscripten::val js_font_record(const fm_font_t& f) {
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", f.name);
  js_set_present_or_null(o, "sizePt", f.has_size_pt, f.size_pt);
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
void js_pull_font_record(const emscripten::val& record, std::string* name_storage, fm_font_t* out,
                         JsNarrowNumericReader& reader) {
  const auto flag = [&](const char* key, const char* field, int32_t* has, int32_t* value) {
    const emscripten::val v = reader.value(record, key, field);
    *has = js_value_present(v) ? 1 : 0;
    *value = reader.boolean_value(v, false) ? 1 : 0;
  };
  *name_storage = reader.string(record, "name", "font.name");
  out->name = name_storage->c_str();
  const emscripten::val size = reader.value(record, "sizePt", "font.sizePt");
  out->has_size_pt = js_value_present(size) ? 1 : 0;
  out->size_pt = reader.number_value(size, 11.0, "font.sizePt");
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

emscripten::val js_fill_record(const fm_fill_t& f) {
  emscripten::val o = emscripten::val::object();
  o.set("pattern", static_cast<std::uint32_t>(f.pattern));
  o.set("foreground", js_color_spec(f.foreground));
  o.set("background", js_color_spec(f.background));
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
  } else {
    o.set("gradient", emscripten::val::null());
  }
  return o;
}

/// Reads a fill record out of a JS object. `stops` owns the gradient stops
/// for the duration of the C ABI call, which borrows the pointer.
fm_fill_t js_pull_fill_record(const emscripten::val& record, std::vector<fm_gradient_stop_t>* stops,
                              JsNarrowNumericReader& reader) {
  fm_fill_t fr{};
  fr.pattern = reader.u8(record, "pattern", 0U, "fill.pattern");
  fr.foreground = js_pull_color_spec(record, "foreground", &reader);
  fr.background = js_pull_color_spec(record, "background", &reader);
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

emscripten::val js_border_record(const fm_border_t& b) {
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
/// `getCellStyleXf`. `styleXfIndex` is set by the callers: a named-style xf is its
/// own parent, so `getCellStyleXf` always reports 0 there.
emscripten::val js_cell_xf_record(const fm_cell_xf_t& xf) {
  emscripten::val o = emscripten::val::object();
  o.set("fontIndex", xf.font_index);
  o.set("fillIndex", xf.fill_index);
  o.set("borderIndex", xf.border_index);
  o.set("numFmtId", static_cast<uint32_t>(xf.num_fmt_id));
  js_set_present_or_null(o, "horizontal", xf.has_horizontal, static_cast<uint32_t>(xf.horizontal));
  js_set_present_or_null(o, "vertical", xf.has_vertical, static_cast<uint32_t>(xf.vertical));
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
  js_set_present_or_null(*object, "textRotation", xf.has_text_rotation, xf.text_rotation);
  js_set_present_or_null(*object, "indent", xf.has_indent, xf.indent);
  js_set_present_or_null(*object, "relativeIndent", xf.has_relative_indent, xf.relative_indent);
  js_set_present_or_null(*object, "shrinkToFit", xf.has_shrink_to_fit, xf.shrink_to_fit != 0);
  js_set_present_or_null(*object, "readingOrder", xf.has_reading_order, xf.reading_order);
}

// Builds the `getCellXf` / `getCellStyleXf` record from the C call's status
// and xf; `undefined` after recording the failure.
emscripten::val js_cell_xf_result(fm_status_t rc, const fm_cell_xf_t& xf) {
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = js_cell_xf_record(xf);
  o.set("styleXfIndex", xf.style_xf_index);
  js_set_cell_xf_alignment(xf, &o);
  return o;
}

}  // namespace

// ---- Per-cell xf index get/set -----------------------------------------

emscripten::val JsWorkbook::getCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) const {
  int32_t present = 0;
  uint32_t xf = 0;
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_cell_xf_index(handle_, sheet, row, col, &present, &xf)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return present != 0 ? emscripten::val(xf) : emscripten::val::null();
}

void JsWorkbook::setCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col, uint32_t xf_index) {
  succeeded(handle_ != nullptr ? fm_sheet_set_cell_xf_index(handle_, sheet, row, col, xf_index)
                               : kBindingInvalidHandle);
}

void JsWorkbook::clearCellXfIndex(uint32_t sheet, uint32_t row, uint32_t col) {
  succeeded(handle_ != nullptr ? fm_sheet_reset_cell_xf_index(handle_, sheet, row, col) : kBindingInvalidHandle);
}

void JsWorkbook::setRangeXfIndex(uint32_t sheet, emscripten::val range, uint32_t xf_index) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.setRangeXfIndex");
  const fm_cell_rect_t rect = js_pull_range(range, &reader);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_range_xf_index(handle_, sheet, &rect, xf_index));
}

// ---- Style record getters ----------------------------------------------

emscripten::val JsWorkbook::getCellXf(uint32_t xf_index) const {
  fm_cell_xf_t xf{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_cell_xf_at(handle_, xf_index, &xf) : kBindingInvalidHandle;
  return js_cell_xf_result(rc, xf);
}

emscripten::val JsWorkbook::getFont(uint32_t font_index) const {
  fm_font_t f{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_font_at(handle_, font_index, &f) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_font_record(f);
}

emscripten::val JsWorkbook::getFill(uint32_t fill_index) const {
  fm_fill_t f{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_fill_at(handle_, fill_index, &f) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_fill_record(f);
}

emscripten::val JsWorkbook::getDefaultFont() const {
  return getFont(0U);
}

emscripten::val JsWorkbook::getBorder(uint32_t border_index) const {
  fm_border_t b{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_border_at(handle_, border_index, &b) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  return js_border_record(b);
}

emscripten::val JsWorkbook::getNumFmt(uint32_t num_fmt_id) const {
  int32_t present = 0;
  const char* s = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_workbook_get_number_format(handle_, num_fmt_id, &present, &s) : kBindingInvalidHandle;
  return optional_or_null(rc, present,
                          rc == 0 && present != 0 ? emscripten::val(std::string(s)) : emscripten::val::null());
}

emscripten::val JsWorkbook::getDxf(uint32_t dxf_index) const {
  fm_dxf_t d{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_dxf_at(handle_, dxf_index, &d) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("font", d.has_font != 0 ? js_font_record(d.font) : emscripten::val::null());
  o.set("fill", d.has_fill != 0 ? js_fill_record(d.fill) : emscripten::val::null());
  o.set("border", d.has_border != 0 ? js_border_record(d.border) : emscripten::val::null());
  o.set("numberFormat",
        d.number_format != nullptr ? emscripten::val(std::string(d.number_format)) : emscripten::val::null());
  o.set("alignment", d.alignment_xml != nullptr && d.alignment_xml[0] != '\0' ? dxf_alignment_to_val(d.alignment_xml)
                                                                              : emscripten::val::null());
  o.set("protection", d.protection_xml != nullptr && d.protection_xml[0] != '\0'
                          ? dxf_protection_to_val(d.protection_xml)
                          : emscripten::val::null());
  return o;
}

// ---- Style record adders -----------------------------------------------

namespace {

/// Calls an adder whose out-parameter is `uint32_t` and widens the index to the caller's `size_t`.
template <typename Fn, typename Record>
fm_status_t add_indexed(Fn fn, fm_workbook_t* handle, const Record* record, std::size_t* idx) {
  uint32_t out = 0;
  const fm_status_t rc = fn(handle, record, &out);
  *idx = out;
  return rc;
}

}  // namespace

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
  fm_font_t fr{};
  return add_style_record(
      "addFont", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_font_record(rec, &name, &fr, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_font(handle, &fr, idx); });
}

void JsWorkbook::setFont(uint32_t font_index, emscripten::val record) {
  std::string name;
  fm_font_t fr{};
  set_style_record(
      "setFont", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_font_record(rec, &name, &fr, reader); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_font(handle, font_index, &fr); });
}

void JsWorkbook::setDefaultFont(emscripten::val record) {
  std::string name;
  fm_font_t fr{};
  set_style_record(
      "setDefaultFont", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_font_record(rec, &name, &fr, reader); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_default_font(handle, &fr); });
}

uint32_t JsWorkbook::addFill(emscripten::val record) {
  std::vector<fm_gradient_stop_t> stops;
  fm_fill_t fr{};
  return add_style_record(
      "addFill", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { fr = js_pull_fill_record(rec, &stops, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_fill(handle, &fr, idx); });
}

void JsWorkbook::setFill(uint32_t fill_index, emscripten::val record) {
  std::vector<fm_gradient_stop_t> stops;
  fm_fill_t fr{};
  set_style_record(
      "setFill", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { fr = js_pull_fill_record(rec, &stops, reader); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_fill(handle, fill_index, &fr); });
}

uint32_t JsWorkbook::addBorder(emscripten::val record) {
  fm_border_t br{};
  return add_style_record(
      "addBorder", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) {
        br = js_pull_border_record(rec, kBorderFields, reader);
      },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_border(handle, &br, idx); });
}

void JsWorkbook::setBorder(uint32_t border_index, emscripten::val record) {
  fm_border_t br{};
  set_style_record(
      "setBorder", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) {
        br = js_pull_border_record(rec, kBorderFields, reader);
      },
      [&](fm_workbook_t* handle) { return fm_workbook_set_border(handle, border_index, &br); });
}

uint32_t JsWorkbook::addNumFmt(const std::string& format_code) {
  uint32_t id = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_add_number_format(handle_, format_code.c_str(), &id)
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
      [&](fm_workbook_t* handle, std::size_t* idx) { return add_indexed(fm_workbook_add_cell_xf, handle, &xf, idx); });
}

void JsWorkbook::setXf(uint32_t xf_index, emscripten::val record) {
  fm_cell_xf_t xf{};
  set_style_record(
      "setCellXf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_cell_xf(rec, &xf, reader); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_cell_xf(handle, xf_index, &xf); });
}

namespace {

// Owns the strings and gradient stops a pulled `fm_dxf_t` borrows.
struct DxfStore {
  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  std::vector<fm_gradient_stop_t> fill_stops;
};

void pull_dxf(JsNarrowNumericReader& reader, const emscripten::val& rec, DxfStore& store, fm_dxf_t& dxf) {
  emscripten::val font = reader.value(rec, "font", "dxf.font");
  if (js_value_present(font)) {
    dxf.has_font = 1;
    js_pull_font_record(font, &store.font_name, &dxf.font, reader);
  }

  emscripten::val fill = reader.value(rec, "fill", "dxf.fill");
  if (js_value_present(fill)) {
    dxf.has_fill = 1;
    dxf.fill = js_pull_fill_record(fill, &store.fill_stops, reader);
  }

  emscripten::val border = reader.value(rec, "border", "dxf.border");
  if (js_value_present(border)) {
    dxf.has_border = 1;
    dxf.border = js_pull_border_record(border, kDxfBorderFields, reader);
  }

  dxf.number_format = reader.optional_string(rec, "numberFormat", store.num_fmt_code, "dxf.numberFormat");

  const emscripten::val alignment = reader.value(rec, "alignment", "dxf.alignment");
  if (js_value_present(alignment)) {
    store.alignment_xml = dxf_alignment_xml(reader, alignment);
  }
  const emscripten::val protection = reader.value(rec, "protection", "dxf.protection");
  if (js_value_present(protection)) {
    store.protection_xml = dxf_protection_xml(reader, protection);
  }
  dxf.alignment_xml = store.alignment_xml.c_str();
  dxf.protection_xml = store.protection_xml.c_str();
}

}  // namespace

uint32_t JsWorkbook::addDxf(emscripten::val record) {
  DxfStore store;
  fm_dxf_t dxf{};
  return add_style_record(
      "addDxf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { pull_dxf(reader, rec, store, dxf); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return add_indexed(fm_workbook_add_dxf, handle, &dxf, idx); });
}

void JsWorkbook::setDxf(uint32_t dxf_index, emscripten::val record) {
  DxfStore store;
  fm_dxf_t dxf{};
  set_style_record(
      "setDxf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { pull_dxf(reader, rec, store, dxf); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_dxf(handle, dxf_index, &dxf); });
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
      handle_ != nullptr ? fm_workbook_get_cell_style_xf_at(handle_, index, &xf) : kBindingInvalidHandle;
  return js_cell_xf_result(rc, xf);
}

uint32_t JsWorkbook::addCellStyleXf(emscripten::val record) {
  fm_cell_xf_t xf{};
  return add_style_record(
      "addCellStyleXf", record,
      [&](JsNarrowNumericReader& reader, const emscripten::val& rec) { js_pull_cell_xf(rec, &xf, reader); },
      [&](fm_workbook_t* handle, std::size_t* idx) {
        return add_indexed(fm_workbook_add_cell_style_xf, handle, &xf, idx);
      });
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
