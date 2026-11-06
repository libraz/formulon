//
// JsWorkbook theme surface: theme colors / fonts / script fonts, color
// resolution and the effective-style readback.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

emscripten::val js_resolved_color(uint32_t argb, int32_t resolution) {
  emscripten::val o = emscripten::val::object();
  o.set("argb", argb);
  o.set("resolution", resolution);
  return o;
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
    item.set("major", major);
    js_set_cstr_fields(item, {{"script", script.c_str()}, {"typeface", typeface}});
  }
  return optional_or_null(rc, present, item);
}

emscripten::val JsWorkbook::getThemeColors() const {
  fm_theme_colors_t colors{};
  int32_t source = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_theme_colors(handle_, &colors, &source)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val list = emscripten::val::array();
  for (uint32_t i = 0; i < 12U; ++i) {
    list.set(i, colors.argb[i]);
  }
  emscripten::val o = emscripten::val::object();
  o.set("source", source);
  o.set("colors", list);
  return o;
}

emscripten::val JsWorkbook::getThemeFonts() const {
  fm_theme_fonts_t fonts{};
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_theme_fonts(handle_, &fonts) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val f = emscripten::val::object();
  js_set_cstr_fields(f, {{"majorLatin", fonts.major_latin},
                         {"majorEastAsian", fonts.major_east_asian},
                         {"minorLatin", fonts.minor_latin},
                         {"minorEastAsian", fonts.minor_east_asian}});
  return f;
}

void JsWorkbook::setThemeColors(emscripten::val colors) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  fm_theme_colors_t tc{};
  JsNarrowNumericReader reader("setThemeColors");
  if (!reader.is_array(colors, "colors") || reader.length(colors, "colors") != std::size(tc.argb)) {
    record_type_error("setThemeColors: `colors` must be an array of 12 ARGB numbers");
    return;
  }
  for (uint32_t i = 0; i < std::size(tc.argb); ++i) {
    tc.argb[i] = reader.u32_value(reader.array_element(colors, i, "colors[]"), 0U, "colors[]");
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_workbook_set_theme_colors(handle_, &tc));
}

void JsWorkbook::setThemeFonts(emscripten::val fonts) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setThemeFonts");
  const std::string major_latin = reader.string(fonts, "majorLatin", "theme.majorLatin");
  const std::string major_ea = reader.string(fonts, "majorEastAsian", "theme.majorEastAsian");
  const std::string minor_latin = reader.string(fonts, "minorLatin", "theme.minorLatin");
  const std::string minor_ea = reader.string(fonts, "minorEastAsian", "theme.minorEastAsian");
  if (!reader.ok()) {
    return;
  }
  fm_theme_fonts_t tf{};
  tf.major_latin = major_latin.c_str();
  tf.major_east_asian = major_ea.c_str();
  tf.minor_latin = minor_latin.c_str();
  tf.minor_east_asian = minor_ea.c_str();
  succeeded(fm_workbook_set_theme_fonts(handle_, &tf));
}

void JsWorkbook::setThemeScriptFont(bool major, const std::string& script, const std::string& typeface) {
  succeeded(handle_ != nullptr
                ? fm_workbook_set_theme_script_font(handle_, major ? 1 : 0, script.c_str(), typeface.c_str())
                : kBindingInvalidHandle);
}

void JsWorkbook::resetTheme() {
  succeeded(handle_ != nullptr ? fm_workbook_clear_theme(handle_) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::resolveColor(emscripten::val spec, int32_t context) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  emscripten::val holder = emscripten::val::object();
  holder.set("c", spec);
  JsNarrowNumericReader reader("resolveColor");
  const fm_color_spec_t color = js_pull_color_spec(holder, "c", &reader);
  if (!reader.ok()) {
    return emscripten::val::undefined();
  }
  uint32_t argb = 0;
  int32_t resolution = 0;
  if (!succeeded(fm_workbook_resolve_color(handle_, &color, context, &argb, &resolution))) {
    return emscripten::val::undefined();
  }
  return js_resolved_color(argb, resolution);
}

emscripten::val JsWorkbook::getEffectiveStyle(uint32_t sheet, uint32_t row, uint32_t col) const {
  fm_effective_style_t e{};
  if (!succeeded(handle_ != nullptr ? fm_sheet_get_effective_style(handle_, sheet, row, col, &e)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
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
