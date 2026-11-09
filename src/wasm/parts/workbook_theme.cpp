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

emscripten::val JsWorkbook::getThemeScriptFont(int32_t scheme, const std::string& script) const {
  const char* typeface = nullptr;
  int32_t present = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_theme_script_font(handle_, scheme, script.c_str(), &present, &typeface)
                             : kBindingInvalidHandle;
  emscripten::val item = emscripten::val::null();
  if (rc == 0 && present != 0) {
    item = emscripten::val::object();
    item.set("scheme", scheme);
    js_set_cstr_fields(item, {{"script", script.c_str()}, {"typeface", typeface}});
  }
  return optional_or_null(rc, present, item);
}

emscripten::val JsWorkbook::listThemeScriptFonts(int32_t scheme) const {
  std::size_t count = 0;
  if (!succeeded(handle_ != nullptr ? fm_workbook_get_theme_script_font_count(handle_, scheme, &count)
                                    : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  emscripten::val list = emscripten::val::array();
  for (std::size_t i = 0; i < count; ++i) {
    const char* script = nullptr;
    const char* typeface = nullptr;
    if (!succeeded(fm_workbook_get_theme_script_font_at(handle_, scheme, i, &script, &typeface))) {
      return emscripten::val::undefined();
    }
    emscripten::val item = emscripten::val::object();
    item.set("scheme", scheme);
    js_set_cstr_fields(item, {{"script", script}, {"typeface", typeface}});
    list.set(static_cast<uint32_t>(i), item);
  }
  return list;
}

emscripten::val JsWorkbook::getThemeColors() const {
  fm_theme_colors_t colors = FM_STRUCT_INIT(fm_theme_colors_t);
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
  fm_theme_fonts_t fonts = FM_STRUCT_INIT(fm_theme_fonts_t);
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
  fm_theme_colors_t tc = FM_STRUCT_INIT(fm_theme_colors_t);
  JsNarrowNumericReader reader("Workbook.setThemeColors");
  if (!reader.is_array(colors, "colors") || reader.length(colors, "colors") != std::size(tc.argb)) {
    record_type_error("Workbook.setThemeColors: `colors` must be an array of 12 RGB colours");
    return;
  }
  for (uint32_t i = 0; i < std::size(tc.argb); ++i) {
    const fm_color_t color = js_pull_color_spec_value(reader.array_element(colors, i, "colors[]"), "colors[]", &reader);
    if (!reader.ok()) {
      return;
    }
    if (color.kind != FM_COLOR_KIND_RGB) {
      record_type_error("Workbook.setThemeColors: each colour must be an RGB colour");
      return;
    }
    tc.argb[i] = color.argb;
  }
  succeeded(fm_workbook_set_theme_colors(handle_, &tc));
}

void JsWorkbook::setThemeFonts(emscripten::val fonts) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Workbook.setThemeFonts");
  const std::string major_latin = reader.required_string(fonts, "majorLatin", "theme.majorLatin");
  const std::string major_ea = reader.required_string(fonts, "majorEastAsian", "theme.majorEastAsian");
  const std::string minor_latin = reader.required_string(fonts, "minorLatin", "theme.minorLatin");
  const std::string minor_ea = reader.required_string(fonts, "minorEastAsian", "theme.minorEastAsian");
  if (!reader.ok()) {
    return;
  }
  fm_theme_fonts_t tf = FM_STRUCT_INIT(fm_theme_fonts_t);
  tf.major_latin = major_latin.c_str();
  tf.major_east_asian = major_ea.c_str();
  tf.minor_latin = minor_latin.c_str();
  tf.minor_east_asian = minor_ea.c_str();
  succeeded(fm_workbook_set_theme_fonts(handle_, &tf));
}

void JsWorkbook::setThemeScriptFont(int32_t scheme, const std::string& script, const std::string& typeface) {
  succeeded(handle_ != nullptr ? fm_workbook_set_theme_script_font(handle_, scheme, script.c_str(), typeface.c_str())
                               : kBindingInvalidHandle);
}

void JsWorkbook::resetTheme() {
  succeeded(handle_ != nullptr ? fm_workbook_reset_theme(handle_) : kBindingInvalidHandle);
}

emscripten::val JsWorkbook::resolveColor(emscripten::val spec, int32_t context) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return emscripten::val::undefined();
  }
  emscripten::val holder = emscripten::val::object();
  holder.set("c", spec);
  JsNarrowNumericReader reader("Workbook.resolveColor");
  const fm_color_t color = js_pull_color_spec(holder, "c", &reader);
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
  fm_effective_style_info_t e = FM_STRUCT_INIT(fm_effective_style_info_t);
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
  o.set("fontColor", js_resolved_color(e.font_argb, e.font_resolution));
  o.set("fillForeground", js_resolved_color(e.fill_fg_argb, e.fill_fg_resolution));
  o.set("fillBackground", js_resolved_color(e.fill_bg_argb, e.fill_bg_resolution));
  static const char* const kSides[5] = {"left", "right", "top", "bottom", "diagonal"};
  emscripten::val borders = emscripten::val::object();
  for (int i = 0; i < 5; ++i) {
    borders.set(kSides[i], js_resolved_color(e.border_argb[i], e.border_resolution[i]));
  }
  o.set("borders", borders);
  emscripten::val protection = emscripten::val::object();
  protection.set("locked", e.locked != 0);
  protection.set("hidden", e.hidden != 0);
  o.set("protection", protection);
  js_set_cstr(o, "numberFormat", e.num_fmt_code);
  return o;
}
}  // namespace parts
}  // namespace wasm
}  // namespace formulon
