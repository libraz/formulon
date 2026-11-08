// Theme, colour resolution and effective-style bindings.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

// ---- Theme, colour resolution, effective style ----------------------

Napi::Value Workbook::GetThemeColors(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_theme_colors_t colors = FM_STRUCT_INIT(fm_theme_colors_t);
  int32_t source = 0;
  const fm_status_t rc = fm_workbook_get_theme_colors(handle_, &colors, &source);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array arr = Napi::Array::New(env, 12);
  for (uint32_t i = 0; i < 12; ++i) {
    arr.Set(i, Napi::Number::New(env, colors.argb[i]));
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("source", Napi::Number::New(env, source));
  out.Set("colors", arr);
  return out;
}

Napi::Value Workbook::GetThemeFonts(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_theme_fonts_t fonts = FM_STRUCT_INIT(fm_theme_fonts_t);
  const fm_status_t rc = fm_workbook_get_theme_fonts(handle_, &fonts);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("majorLatin", JsString(env, fonts.major_latin));
  out.Set("majorEastAsian", JsString(env, fonts.major_east_asian));
  out.Set("minorLatin", JsString(env, fonts.minor_latin));
  out.Set("minorEastAsian", JsString(env, fonts.minor_east_asian));
  return out;
}

Napi::Value Workbook::SetThemeColors(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr;
  if (!ArrayArg(info, 0, &arr) || arr.Length() != 12) {
    Napi::TypeError::New(env, "setThemeColors expects an array of 12 colors").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  // Each scheme colour is a `Color` record; only its `argb` is stored.
  CheckedSpecReader reader(env);
  fm_theme_colors_t colors = FM_STRUCT_INIT(fm_theme_colors_t);
  for (uint32_t i = 0; i < 12 && reader.ok(); ++i) {
    Napi::Value element;
    if (reader.ArrayElement(arr, i, &element)) {
      colors.argb[i] = ReadColorSpec(reader, element, "colors[]").argb;
    }
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_workbook_set_theme_colors(handle_, &colors), env.Undefined());
}

Napi::Value Workbook::SetThemeFonts(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object record = ArgObject(info, 0);
  CheckedSpecReader reader(env);
  std::string major_latin;
  std::string major_ea;
  std::string minor_latin;
  std::string minor_ea;
  major_latin = reader.RequiredString(record, "majorLatin");
  major_ea = reader.RequiredString(record, "majorEastAsian");
  minor_latin = reader.RequiredString(record, "minorLatin");
  minor_ea = reader.RequiredString(record, "minorEastAsian");
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_theme_fonts_t fonts = FM_STRUCT_INIT(fm_theme_fonts_t);
  fonts.major_latin = major_latin.c_str();
  fonts.major_east_asian = major_ea.c_str();
  fonts.minor_latin = minor_latin.c_str();
  fonts.minor_east_asian = minor_ea.c_str();
  return ValueOrThrow(env, fm_workbook_set_theme_fonts(handle_, &fonts), env.Undefined());
}

Napi::Value Workbook::GetThemeScriptFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const int32_t scheme = ArgI32(info, 0);
  const std::string script = ArgString(info, 1);
  const char* typeface = nullptr;
  int32_t present = 0;
  const fm_status_t rc = fm_workbook_get_theme_script_font(handle_, scheme, script.c_str(), &present, &typeface);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object item = Napi::Object::New(env);
  item.Set("scheme", Napi::Number::New(env, scheme));
  item.Set("script", JsString(env, script.c_str()));
  item.Set("typeface", JsString(env, typeface));
  return OptionalOrNull(env, rc, present, item);
}

Napi::Value Workbook::SetThemeScriptFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string script = ArgString(info, 1);
  const std::string typeface = ArgString(info, 2);
  const fm_status_t rc =
      fm_workbook_set_theme_script_font(handle_, ArgI32(info, 0), script.c_str(), typeface.c_str());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::ResetTheme(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(env, fm_workbook_reset_theme(handle_), env.Undefined());
}

Napi::Value Workbook::ResolveColor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  fm_color_spec_t spec{};
  if (info.Length() > 0 && info[0].IsObject()) {
    Napi::Object holder = Napi::Object::New(env);
    holder.Set("spec", info[0]);
    spec = PullColorSpec(reader, holder, "spec");
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t argb = 0;
  int32_t resolution = 0;
  const fm_status_t rc = fm_workbook_resolve_color(handle_, &spec, ArgI32(info, 1), &argb, &resolution);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("argb", Napi::Number::New(env, argb));
  out.Set("resolution", Napi::Number::New(env, resolution));
  return out;
}

Napi::Value Workbook::GetEffectiveStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_effective_style_t es = FM_STRUCT_INIT(fm_effective_style_t);
  const fm_status_t rc = fm_sheet_get_effective_style(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &es);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  const auto colorOf = [&env](uint32_t argb, int32_t resolution) {
    Napi::Object c = Napi::Object::New(env);
    c.Set("argb", Napi::Number::New(env, argb));
    c.Set("resolution", Napi::Number::New(env, resolution));
    return c;
  };
  static const char* const kSides[5] = {"left", "right", "top", "bottom", "diagonal"};
  Napi::Object borders = Napi::Object::New(env);
  for (size_t i = 0; i < 5; ++i) {
    borders.Set(kSides[i], colorOf(es.border_argb[i], es.border_resolution[i]));
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("xfIndex", Napi::Number::New(env, es.xf_index));
  out.Set("source", Napi::Number::New(env, es.source));
  out.Set("fontIndex", Napi::Number::New(env, es.font_index));
  out.Set("fillIndex", Napi::Number::New(env, es.fill_index));
  out.Set("borderIndex", Napi::Number::New(env, es.border_index));
  out.Set("fontColor", colorOf(es.font_argb, es.font_resolution));
  out.Set("fillForeground", colorOf(es.fill_fg_argb, es.fill_fg_resolution));
  out.Set("fillBackground", colorOf(es.fill_bg_argb, es.fill_bg_resolution));
  out.Set("borders", borders);
  Napi::Object protection = Napi::Object::New(env);
  protection.Set("locked", Napi::Boolean::New(env, es.locked != 0));
  protection.Set("hidden", Napi::Boolean::New(env, es.hidden != 0));
  out.Set("protection", protection);
  out.Set("numberFormat", JsString(env, es.num_fmt_code));
  return out;
}

}  // namespace formulon_node
