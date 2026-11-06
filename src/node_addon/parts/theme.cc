// Theme, colour resolution and effective-style bindings.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

/// Reads the major or minor script-font list as `{script, typeface}` records.
/// Each pair is copied before the next `_at` call reuses the read scratch.
Napi::Array ThemeScriptFonts(Napi::Env env, const fm_workbook_t* handle, bool major) {
  Napi::Array list = Napi::Array::New(env);
  size_t count = 0;
  if (fm_workbook_get_theme_script_font_count(handle, major ? 1 : 0, &count) != 0) {
    return list;
  }
  for (size_t i = 0; i < count; ++i) {
    const char* script = nullptr;
    const char* typeface = nullptr;
    if (fm_workbook_get_theme_script_font_at(handle, major ? 1 : 0, i, &script, &typeface) != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("script", JsString(env, script));
    item.Set("typeface", JsString(env, typeface));
    list.Set(static_cast<uint32_t>(i), item);
  }
  return list;
}

}  // namespace

// ---- Theme, colour resolution, effective style ----------------------

Napi::Value Workbook::GetTheme(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_theme_colors_t colors{};
  fm_theme_fonts_t fonts{};
  int32_t source = 0;
  fm_status_t rc = fm_workbook_get_theme_colors(handle_, &colors, &source);
  if (rc == 0) {
    rc = fm_workbook_get_theme_fonts(handle_, &fonts);
  }
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array arr = Napi::Array::New(env, 12);
  for (uint32_t i = 0; i < 12; ++i) {
    arr.Set(i, Napi::Number::New(env, colors.argb[i]));
  }
  const auto face = [&env](const char* s) { return JsString(env, s); };
  Napi::Object fontsOut = Napi::Object::New(env);
  fontsOut.Set("majorLatin", face(fonts.major_latin));
  fontsOut.Set("majorEastAsian", face(fonts.major_east_asian));
  fontsOut.Set("minorLatin", face(fonts.minor_latin));
  fontsOut.Set("minorEastAsian", face(fonts.minor_east_asian));
  Napi::Object out = Napi::Object::New(env);
  out.Set("source", Napi::Number::New(env, source));
  out.Set("colors", arr);
  out.Set("fonts", fontsOut);
  out.Set("majorScripts", ThemeScriptFonts(env, handle_, true));
  out.Set("minorScripts", ThemeScriptFonts(env, handle_, false));
  return out;
}

Napi::Value Workbook::SetThemeColors(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr;
  if (!ArrayArg(info, 0, &arr) || arr.Length() != 12) {
    Napi::TypeError::New(env, "setThemeColors expects an array of 12 ARGB numbers").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_theme_colors_t colors{};
  for (uint32_t i = 0; i < 12; ++i) {
    const std::string key = std::to_string(i);
    colors.argb[i] = reader.U32(arr, key.c_str(), 0U);
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
  reader.String(record, "majorLatin", &major_latin);
  reader.String(record, "majorEastAsian", &major_ea);
  reader.String(record, "minorLatin", &minor_latin);
  reader.String(record, "minorEastAsian", &minor_ea);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_theme_fonts_t fonts{};
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
  const std::string script = ArgString(info, 1);
  const char* typeface = nullptr;
  int32_t present = 0;
  const fm_status_t rc =
      fm_workbook_get_theme_script_font(handle_, ArgBool(info, 0) ? 1 : 0, script.c_str(), &present, &typeface);
  if (rc != 0 || present == 0) {
    return OptionalOrNull(env, rc, present, env.Null());
  }
  Napi::Object item = Napi::Object::New(env);
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
      fm_workbook_set_theme_script_font(handle_, ArgBool(info, 0) ? 1 : 0, script.c_str(), typeface.c_str());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::ResetTheme(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(env, fm_workbook_clear_theme(handle_), env.Undefined());
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
  fm_effective_style_t es{};
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
  out.Set("font", colorOf(es.font_argb, es.font_resolution));
  out.Set("fillForeground", colorOf(es.fill_fg_argb, es.fill_fg_resolution));
  out.Set("fillBackground", colorOf(es.fill_bg_argb, es.fill_bg_resolution));
  out.Set("borders", borders);
  out.Set("locked", Napi::Boolean::New(env, es.locked != 0));
  out.Set("hidden", Napi::Boolean::New(env, es.hidden != 0));
  out.Set("numFmtCode", JsString(env, es.num_fmt_code));
  return out;
}

}  // namespace formulon_node
