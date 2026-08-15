// Style-table bindings: read / write entries in the xf / font / fill /
// border / numFmt pools plus the per-cell xf-index getters, the named
// cell-style accessors, the conditional-formatting read / mutate
// surface, and the `evaluateCfRange` evaluator. The cf translator
// lives here because it shares the styles compilation dependency on
// `fm_styles_*` typedefs declared in the same C ABI header block.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {
namespace {

fm_cf_color_t PullCfColor(const Napi::Object& spec) {
  fm_cf_color_t out{};
  out.r = static_cast<uint8_t>(SpecPullU32(spec, "r", 0U) & 0xFFU);
  out.g = static_cast<uint8_t>(SpecPullU32(spec, "g", 0U) & 0xFFU);
  out.b = static_cast<uint8_t>(SpecPullU32(spec, "b", 0U) & 0xFFU);
  out.a = static_cast<uint8_t>(SpecPullU32(spec, "a", 255U) & 0xFFU);
  return out;
}

fm_cfvo_t PullCfvo(const Napi::Object& spec, std::vector<std::string>* strings) {
  fm_cfvo_t out{};
  out.type = static_cast<uint8_t>(SpecPullU32(spec, "type", 0U) & 0xFFU);
  out.gte = SpecPullBool(spec, "gte", true) ? 1 : 0;
  if (SpecHas(spec, "value")) {
    strings->push_back(spec.Get("value").ToString().Utf8Value());
    out.value = strings->back().c_str();
  }
  return out;
}

Napi::Object CfColorToJs(Napi::Env env, fm_cf_color_t color) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("r", Napi::Number::New(env, color.r));
  out.Set("g", Napi::Number::New(env, color.g));
  out.Set("b", Napi::Number::New(env, color.b));
  out.Set("a", Napi::Number::New(env, color.a));
  return out;
}

Napi::Object CfvoToJs(Napi::Env env, const fm_cfvo_t& cfvo) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("type", Napi::Number::New(env, static_cast<uint32_t>(cfvo.type)));
  out.Set("gte", Napi::Boolean::New(env, cfvo.gte != 0));
  if (cfvo.value != nullptr) {
    out.Set("value", Napi::String::New(env, cfvo.value));
  }
  return out;
}

/// Reads the `{kind, rgb, theme, tint, indexed}` colour specification out
/// of `owner[key]`. An absent object leaves `kind` at `kFmColorNone`, so
/// the writer emits the sibling `*Argb` as literal `rgb`. A supplied
/// selector is authoritative; this binding does not resolve theme/indexed /
/// auto colours.
fm_color_spec PullColorSpec(const Napi::Object& owner, const char* key) {
  fm_color_spec spec{};
  if (!SpecHas(owner, key) || !owner.Get(key).IsObject()) {
    return spec;
  }
  Napi::Object v = owner.Get(key).As<Napi::Object>();
  spec.kind = static_cast<uint8_t>(SpecPullU32(v, "kind", 0U) & 0xFFU);
  spec.rgb = SpecPullU32(v, "rgb", 0U);
  spec.theme = SpecPullU32(v, "theme", 0U);
  spec.tint = SpecPullDouble(v, "tint", 0.0);
  spec.indexed = SpecPullU32(v, "indexed", 0U);
  return spec;
}

Napi::Object ColorSpecToJs(Napi::Env env, const fm_color_spec& spec) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("kind", Napi::Number::New(env, static_cast<uint32_t>(spec.kind)));
  out.Set("rgb", Napi::Number::New(env, spec.rgb));
  out.Set("theme", Napi::Number::New(env, spec.theme));
  out.Set("tint", Napi::Number::New(env, spec.tint));
  out.Set("indexed", Napi::Number::New(env, spec.indexed));
  return out;
}

/// Builds the JS mirror of a font record. Shared by `getFont` and the
/// `<dxf>` font projection so both surface the same field set.
Napi::Object FontRecordToJs(Napi::Env env, const fm_font_record& f) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", Napi::String::New(env, f.name != nullptr ? f.name : ""));
  out.Set("size", Napi::Number::New(env, f.size));
  out.Set("colorArgb", Napi::Number::New(env, f.color_argb));
  out.Set("bold", Napi::Boolean::New(env, f.bold != 0));
  out.Set("italic", Napi::Boolean::New(env, f.italic != 0));
  out.Set("strike", Napi::Boolean::New(env, f.strike != 0));
  out.Set("hasBold", Napi::Boolean::New(env, f.has_bold != 0));
  out.Set("hasItalic", Napi::Boolean::New(env, f.has_italic != 0));
  out.Set("hasStrike", Napi::Boolean::New(env, f.has_strike != 0));
  out.Set("underline", Napi::Number::New(env, static_cast<uint32_t>(f.underline)));
  out.Set("vertAlign", Napi::Number::New(env, static_cast<uint32_t>(f.vert_align)));
  out.Set("hasFamily", Napi::Boolean::New(env, f.has_family != 0));
  out.Set("family", Napi::Number::New(env, static_cast<uint32_t>(f.family)));
  out.Set("hasCharset", Napi::Boolean::New(env, f.has_charset != 0));
  out.Set("charset", Napi::Number::New(env, static_cast<uint32_t>(f.charset)));
  out.Set("color", ColorSpecToJs(env, f.color));
  return out;
}

/// Reads a font record out of a JS object. `name_storage` owns the font
/// name for the duration of the C ABI call, which borrows the pointer.
void PullFontRecord(const Napi::Object& record, std::string* name_storage, fm_font_record* out) {
  if (SpecHas(record, "name")) {
    *name_storage = record.Get("name").ToString().Utf8Value();
  } else {
    name_storage->clear();
  }
  out->name = name_storage->c_str();
  out->size = SpecPullDouble(record, "size", 11.0);
  out->bold = SpecPullBool(record, "bold", false) ? 1 : 0;
  out->italic = SpecPullBool(record, "italic", false) ? 1 : 0;
  out->strike = SpecPullBool(record, "strike", false) ? 1 : 0;
  out->has_bold = SpecPullBool(record, "hasBold", false) ? 1 : 0;
  out->has_italic = SpecPullBool(record, "hasItalic", false) ? 1 : 0;
  out->has_strike = SpecPullBool(record, "hasStrike", false) ? 1 : 0;
  out->underline = static_cast<uint8_t>(SpecPullU32(record, "underline", 0U) & 0xFFU);
  out->vert_align = static_cast<uint8_t>(SpecPullU32(record, "vertAlign", 0U) & 0xFFU);
  out->has_family = SpecPullBool(record, "hasFamily", false) ? 1 : 0;
  out->family = static_cast<uint8_t>(SpecPullU32(record, "family", 0U) & 0xFFU);
  out->has_charset = SpecPullBool(record, "hasCharset", false) ? 1 : 0;
  out->charset = static_cast<uint8_t>(SpecPullU32(record, "charset", 0U) & 0xFFU);
  out->color_argb = SpecPullU32(record, "colorArgb", 0xFF000000U);
  out->color = PullColorSpec(record, "color");
}

Napi::Object FillRecordToJs(Napi::Env env, const fm_fill_record& f) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("pattern", Napi::Number::New(env, static_cast<uint32_t>(f.pattern)));
  out.Set("fgArgb", Napi::Number::New(env, f.fg_argb));
  out.Set("bgArgb", Napi::Number::New(env, f.bg_argb));
  out.Set("fg", ColorSpecToJs(env, f.fg));
  out.Set("bg", ColorSpecToJs(env, f.bg));
  return out;
}

fm_fill_record PullFillRecord(const Napi::Object& record) {
  fm_fill_record out{};
  out.pattern = static_cast<uint8_t>(SpecPullU32(record, "pattern", 0U) & 0xFFU);
  out.fg_argb = SpecPullU32(record, "fgArgb", 0U);
  out.bg_argb = SpecPullU32(record, "bgArgb", 0U);
  out.fg = PullColorSpec(record, "fg");
  out.bg = PullColorSpec(record, "bg");
  return out;
}

Napi::Object BorderSideToJs(Napi::Env env, const fm_border_side& s) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("style", Napi::Number::New(env, static_cast<uint32_t>(s.style)));
  out.Set("colorArgb", Napi::Number::New(env, s.color_argb));
  out.Set("color", ColorSpecToJs(env, s.color));
  return out;
}

Napi::Object BorderRecordToJs(Napi::Env env, const fm_border_record& b) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("left", BorderSideToJs(env, b.left));
  out.Set("right", BorderSideToJs(env, b.right));
  out.Set("top", BorderSideToJs(env, b.top));
  out.Set("bottom", BorderSideToJs(env, b.bottom));
  out.Set("diagonal", BorderSideToJs(env, b.diagonal));
  out.Set("diagonalUp", Napi::Boolean::New(env, b.diagonal_up != 0));
  out.Set("diagonalDown", Napi::Boolean::New(env, b.diagonal_down != 0));
  return out;
}

fm_border_side PullBorderSide(const Napi::Object& owner, const char* key) {
  fm_border_side s{};
  if (!SpecHas(owner, key) || !owner.Get(key).IsObject()) {
    return s;
  }
  Napi::Object v = owner.Get(key).As<Napi::Object>();
  s.style = static_cast<uint8_t>(SpecPullU32(v, "style", 0U) & 0xFFU);
  s.color_argb = SpecPullU32(v, "colorArgb", 0U);
  s.color = PullColorSpec(v, "color");
  return s;
}

fm_border_record PullBorderRecord(const Napi::Object& record) {
  fm_border_record out{};
  out.left = PullBorderSide(record, "left");
  out.right = PullBorderSide(record, "right");
  out.top = PullBorderSide(record, "top");
  out.bottom = PullBorderSide(record, "bottom");
  out.diagonal = PullBorderSide(record, "diagonal");
  out.diagonal_up = SpecPullBool(record, "diagonalUp", false) ? 1 : 0;
  out.diagonal_down = SpecPullBool(record, "diagonalDown", false) ? 1 : 0;
  return out;
}

}  // namespace

// ---- Cell-XF index --------------------------------------------------

Napi::Value Workbook::GetCellXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "xfIndex", 0);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  uint32_t xf = 0;
  fm_status_t rc = fm_cell_get_xf_index(handle_, sheet, row, col, &xf);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "xfIndex", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "xfIndex", xf);
}

Napi::Value Workbook::SetCellXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t row = ArgU32(info, 1);
  const uint32_t col = ArgU32(info, 2);
  const uint32_t xf = ArgU32(info, 3);
  fm_status_t rc = fm_cell_set_xf_index(handle_, sheet, row, col, xf);
  return MakeStatus(env, rc);
}

// ---- Style getters --------------------------------------------------

Napi::Value Workbook::GetCellXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t xf_index = ArgU32(info, 0);
  fm_cell_xf xf{};
  fm_status_t rc = fm_styles_get_cell_xf(handle_, xf_index, &xf);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("fontIndex", Napi::Number::New(env, xf.font_index));
  out.Set("fillIndex", Napi::Number::New(env, xf.fill_index));
  out.Set("borderIndex", Napi::Number::New(env, xf.border_index));
  out.Set("numFmtId", Napi::Number::New(env, static_cast<uint32_t>(xf.num_fmt_id)));
  out.Set("horizontalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.horizontal_align)));
  out.Set("verticalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.vertical_align)));
  out.Set("wrapText", Napi::Boolean::New(env, xf.wrap_text != 0));
  out.Set("justifyLastLine", Napi::Boolean::New(env, xf.justify_last_line != 0));
  out.Set("hasAlignment", Napi::Boolean::New(env, xf.has_alignment != 0));
  out.Set("hasHorizontalAlign", Napi::Boolean::New(env, xf.has_horizontal_align != 0));
  out.Set("hasVerticalAlign", Napi::Boolean::New(env, xf.has_vertical_align != 0));
  out.Set("hasWrapText", Napi::Boolean::New(env, xf.has_wrap_text != 0));
  out.Set("hasJustifyLastLine", Napi::Boolean::New(env, xf.has_justify_last_line != 0));
  out.Set("xfId", Napi::Number::New(env, xf.xf_id));
  if (xf.has_text_rotation != 0) {
    out.Set("textRotation", Napi::Number::New(env, xf.text_rotation));
  }
  if (xf.has_indent != 0) {
    out.Set("indent", Napi::Number::New(env, xf.indent));
  }
  if (xf.has_relative_indent != 0) {
    out.Set("relativeIndent", Napi::Number::New(env, xf.relative_indent));
  }
  if (xf.has_shrink_to_fit != 0) {
    out.Set("shrinkToFit", Napi::Boolean::New(env, xf.shrink_to_fit != 0));
  }
  if (xf.has_reading_order != 0) {
    out.Set("readingOrder", Napi::Number::New(env, xf.reading_order));
  }
  return out;
}

Napi::Value Workbook::GetFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t font_index = ArgU32(info, 0);
  fm_font_record f{};
  fm_status_t rc = fm_styles_get_font(handle_, font_index, &f);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out = FontRecordToJs(env, f);
  out.Set("status", MakeOkStatus(env));
  return out;
}

Napi::Value Workbook::GetFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t fill_index = ArgU32(info, 0);
  fm_fill_record f{};
  fm_status_t rc = fm_styles_get_fill(handle_, fill_index, &f);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out = FillRecordToJs(env, f);
  out.Set("status", MakeOkStatus(env));
  return out;
}

Napi::Value Workbook::GetBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t border_index = ArgU32(info, 0);
  fm_border_record b{};
  fm_status_t rc = fm_styles_get_border(handle_, border_index, &b);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out = BorderRecordToJs(env, b);
  out.Set("status", MakeOkStatus(env));
  return out;
}

Napi::Value Workbook::GetNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t num_fmt_id = ArgU32(info, 0);
  const char* s = nullptr;
  fm_status_t rc = fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("numFmtId", Napi::Number::New(env, num_fmt_id));
  out.Set("formatCode", Napi::String::New(env, s != nullptr ? s : ""));
  return out;
}

Napi::Value Workbook::GetDxf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t dxf_index = ArgU32(info, 0);
  fm_dxf_record d{};
  fm_status_t rc = fm_styles_get_dxf(handle_, dxf_index, &d);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  if (d.font_engaged != 0) {
    out.Set("font", FontRecordToJs(env, d.font));
  }
  if (d.fill_engaged != 0) {
    out.Set("fill", FillRecordToJs(env, d.fill));
  }
  if (d.border_engaged != 0) {
    out.Set("border", BorderRecordToJs(env, d.border));
  }
  if (d.num_fmt_engaged != 0) {
    Napi::Object num_fmt = Napi::Object::New(env);
    num_fmt.Set("numFmtId", Napi::Number::New(env, static_cast<uint32_t>(d.num_fmt_id)));
    num_fmt.Set("formatCode", Napi::String::New(env, d.num_fmt_code != nullptr ? d.num_fmt_code : ""));
    out.Set("numFmt", num_fmt);
  }
  if (d.alignment_xml != nullptr && d.alignment_xml[0] != '\0') {
    out.Set("alignmentXml", Napi::String::New(env, d.alignment_xml));
  }
  if (d.protection_xml != nullptr && d.protection_xml[0] != '\0') {
    out.Set("protectionXml", Napi::String::New(env, d.protection_xml));
  }
  return out;
}

// ---- Style adders ---------------------------------------------------

Napi::Value Workbook::AddFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  std::string name;
  fm_font_record fr{};
  PullFontRecord(record, &name, &fr);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  const fm_fill_record fr = PullFillRecord(record);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  const fm_border_record br = PullBorderRecord(record);
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "numFmtId", 0);
  }
  const std::string code = ArgString(info, 0);
  uint16_t id = 0;
  fm_status_t rc = fm_styles_add_num_fmt(handle_, code.c_str(), &id);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "numFmtId", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "numFmtId", static_cast<uint32_t>(id));
}

Napi::Value Workbook::AddXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);
  fm_cell_xf xf{};
  xf.font_index = record.Has("fontIndex") ? record.Get("fontIndex").ToNumber().Uint32Value() : 0U;
  xf.fill_index = record.Has("fillIndex") ? record.Get("fillIndex").ToNumber().Uint32Value() : 0U;
  xf.border_index = record.Has("borderIndex") ? record.Get("borderIndex").ToNumber().Uint32Value() : 0U;
  xf.num_fmt_id = record.Has("numFmtId") ? static_cast<uint16_t>(record.Get("numFmtId").ToNumber().Uint32Value()) : 0U;
  xf.horizontal_align =
      record.Has("horizontalAlign") ? static_cast<uint8_t>(record.Get("horizontalAlign").ToNumber().Uint32Value()) : 0U;
  xf.vertical_align =
      record.Has("verticalAlign") ? static_cast<uint8_t>(record.Get("verticalAlign").ToNumber().Uint32Value()) : 2U;
  xf.wrap_text = (record.Has("wrapText") && record.Get("wrapText").ToBoolean().Value()) ? 1 : 0;
  xf.justify_last_line = SpecPullBool(record, "justifyLastLine", false) ? 1 : 0;
  xf.xf_id = SpecPullU32(record, "xfId", 0U);
  xf.has_horizontal_align = SpecHas(record, "hasHorizontalAlign")
                                ? (SpecPullBool(record, "hasHorizontalAlign", false) ? 1 : 0)
                                : (SpecHas(record, "horizontalAlign") ? 1 : 0);
  xf.has_vertical_align = SpecHas(record, "hasVerticalAlign")
                              ? (SpecPullBool(record, "hasVerticalAlign", false) ? 1 : 0)
                              : (SpecHas(record, "verticalAlign") ? 1 : 0);
  xf.has_wrap_text = SpecHas(record, "hasWrapText") ? (SpecPullBool(record, "hasWrapText", false) ? 1 : 0)
                                                    : (SpecHas(record, "wrapText") ? 1 : 0);
  xf.has_justify_last_line = SpecHas(record, "hasJustifyLastLine")
                                 ? (SpecPullBool(record, "hasJustifyLastLine", false) ? 1 : 0)
                                 : (SpecHas(record, "justifyLastLine") ? 1 : 0);
  const bool has_explicit_alignment = SpecHas(record, "hasAlignment");
  const bool has_supplied_alignment =
      SpecHas(record, "horizontalAlign") || SpecHas(record, "verticalAlign") || SpecHas(record, "wrapText") ||
      SpecHas(record, "justifyLastLine") || SpecHas(record, "textRotation") || SpecHas(record, "indent") ||
      SpecHas(record, "relativeIndent") || SpecHas(record, "shrinkToFit") || SpecHas(record, "readingOrder") ||
      SpecHas(record, "hasHorizontalAlign") || SpecHas(record, "hasVerticalAlign") || SpecHas(record, "hasWrapText") ||
      SpecHas(record, "hasJustifyLastLine");
  xf.has_alignment =
      has_explicit_alignment ? (SpecPullBool(record, "hasAlignment", false) ? 1 : 0) : (has_supplied_alignment ? 1 : 0);
  if (SpecHas(record, "textRotation")) {
    xf.has_text_rotation = 1;
    xf.text_rotation = SpecPullU32(record, "textRotation", 0U);
  }
  if (SpecHas(record, "indent")) {
    xf.has_indent = 1;
    xf.indent = SpecPullU32(record, "indent", 0U);
  }
  if (SpecHas(record, "relativeIndent")) {
    xf.has_relative_indent = 1;
    xf.relative_indent = SpecPullInt32(record, "relativeIndent", 0);
  }
  if (SpecHas(record, "shrinkToFit")) {
    xf.has_shrink_to_fit = 1;
    xf.shrink_to_fit = SpecPullBool(record, "shrinkToFit", false) ? 1 : 0;
  }
  if (SpecHas(record, "readingOrder")) {
    xf.has_reading_order = 1;
    xf.reading_order = SpecPullU32(record, "readingOrder", 0U);
  }
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

Napi::Value Workbook::AddDxf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = (info.Length() > 0 && info[0].IsObject()) ? info[0].As<Napi::Object>() : Napi::Object::New(env);

  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  fm_dxf_record dxf{};

  if (record.Has("font") && record.Get("font").IsObject()) {
    dxf.font_engaged = 1;
    PullFontRecord(record.Get("font").As<Napi::Object>(), &font_name, &dxf.font);
  }

  if (record.Has("fill") && record.Get("fill").IsObject()) {
    dxf.fill_engaged = 1;
    dxf.fill = PullFillRecord(record.Get("fill").As<Napi::Object>());
  }

  if (record.Has("border") && record.Get("border").IsObject()) {
    dxf.border_engaged = 1;
    dxf.border = PullBorderRecord(record.Get("border").As<Napi::Object>());
  }

  if (record.Has("numFmt") && record.Get("numFmt").IsObject()) {
    Napi::Object num_fmt = record.Get("numFmt").As<Napi::Object>();
    dxf.num_fmt_engaged = 1;
    dxf.num_fmt_id = static_cast<uint16_t>(SpecPullU32(num_fmt, "numFmtId", 0U) & 0xFFFFU);
    if (num_fmt.Has("formatCode") && !num_fmt.Get("formatCode").IsUndefined() && !num_fmt.Get("formatCode").IsNull()) {
      num_fmt_code = num_fmt.Get("formatCode").ToString().Utf8Value();
    }
    dxf.num_fmt_code = num_fmt_code.c_str();
  }

  if (record.Has("alignmentXml") && !record.Get("alignmentXml").IsUndefined() && !record.Get("alignmentXml").IsNull()) {
    alignment_xml = record.Get("alignmentXml").ToString().Utf8Value();
  }
  if (record.Has("protectionXml") && !record.Get("protectionXml").IsUndefined() &&
      !record.Get("protectionXml").IsNull()) {
    protection_xml = record.Get("protectionXml").ToString().Utf8Value();
  }
  dxf.alignment_xml = alignment_xml.c_str();
  dxf.protection_xml = protection_xml.c_str();

  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_dxf(handle_, dxf, &idx);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", idx);
}

// ---- Style pool counts ----------------------------------------------
//
// `FontCount` / `FillCount` / `BorderCount` / `XfCount` are now emitted
// by the binding codegen (see `src/node_addon/generated/styles_counts.cc`).

Napi::Value Workbook::DxfCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_dxf_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

// ---- Conditional formatting -----------------------------------------

Napi::Value Workbook::EvaluateCfRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("cells", Napi::Array::New(env));
    return out;
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t first_row = ArgU32(info, 1);
  const uint32_t first_col = ArgU32(info, 2);
  const uint32_t last_row = ArgU32(info, 3);
  const uint32_t last_col = ArgU32(info, 4);
  // A missing `todaySerial` must disable `timePeriod` rules, matching the
  // Python binding and the C ABI contract. `ArgDouble`'s 0.0 fallback is a
  // valid Excel serial (1899-12-30), so it would silently evaluate those
  // rules against that date instead of skipping them.
  const double today_serial =
      info.Length() > 5 ? info[5].ToNumber().DoubleValue() : std::numeric_limits<double>::quiet_NaN();
  fm_cf_results_t* results = nullptr;
  fm_status_t rc =
      fm_workbook_cf_evaluate_range(handle_, sheet, first_row, first_col, last_row, last_col, today_serial, &results);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("cells", Napi::Array::New(env));
    return out;
  }
  const std::size_t cell_count = fm_cf_results_cell_count(results);
  Napi::Array cells = Napi::Array::New(env, cell_count);
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < cell_count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    std::size_t match_count = 0;
    if (fm_cf_results_cell_at(results, i, &row, &col, &match_count) != 0) {
      // Defensive: skip entries the C ABI declines to materialise.
      continue;
    }
    Napi::Array matches = Napi::Array::New(env, match_count);
    std::size_t mj = 0;
    for (std::size_t j = 0; j < match_count; ++j) {
      fm_cf_match_t m{};
      if (fm_cf_results_match_at(results, i, j, &m) != 0) {
        continue;
      }
      matches.Set(static_cast<uint32_t>(mj), TranslateCfMatch(env, m));
      ++mj;
    }
    Napi::Object cell = Napi::Object::New(env);
    cell.Set("row", Napi::Number::New(env, row));
    cell.Set("col", Napi::Number::New(env, col));
    cell.Set("matches", matches);
    cells.Set(static_cast<uint32_t>(emitted), cell);
    ++emitted;
  }
  fm_cf_results_destroy(results);
  out.Set("status", MakeOkStatus(env));
  out.Set("cells", cells);
  return out;
}

// ---- Named cell styles ----------------------------------------------

Napi::Value Workbook::CellStyleCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_style_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

Napi::Value Workbook::CellStyleXfCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  uint32_t n = 0;
  if (fm_styles_get_cell_style_xf_count(handle_, &n) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, n);
}

Napi::Value Workbook::GetCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t index = ArgU32(info, 0);
  fm_cell_style_record_t cs{};
  fm_status_t rc = fm_styles_get_cell_style(handle_, index, &cs);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("name", Napi::String::New(env, cs.name != nullptr ? cs.name : ""));
  out.Set("xfId", Napi::Number::New(env, cs.xf_id));
  out.Set("builtinId", Napi::Number::New(env, cs.builtin_id));
  out.Set("iLevel", Napi::Number::New(env, cs.i_level));
  out.Set("hidden", Napi::Boolean::New(env, cs.hidden != 0));
  out.Set("customBuiltin", Napi::Boolean::New(env, cs.custom_builtin != 0));
  return out;
}

Napi::Value Workbook::GetCellStyleXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint32_t index = ArgU32(info, 0);
  fm_cell_xf xf{};
  fm_status_t rc = fm_styles_get_cell_style_xf(handle_, index, &xf);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("fontIndex", Napi::Number::New(env, xf.font_index));
  out.Set("fillIndex", Napi::Number::New(env, xf.fill_index));
  out.Set("borderIndex", Napi::Number::New(env, xf.border_index));
  out.Set("numFmtId", Napi::Number::New(env, static_cast<uint32_t>(xf.num_fmt_id)));
  out.Set("horizontalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.horizontal_align)));
  out.Set("verticalAlign", Napi::Number::New(env, static_cast<uint32_t>(xf.vertical_align)));
  out.Set("wrapText", Napi::Boolean::New(env, xf.wrap_text != 0));
  out.Set("justifyLastLine", Napi::Boolean::New(env, xf.justify_last_line != 0));
  out.Set("hasAlignment", Napi::Boolean::New(env, xf.has_alignment != 0));
  out.Set("hasHorizontalAlign", Napi::Boolean::New(env, xf.has_horizontal_align != 0));
  out.Set("hasVerticalAlign", Napi::Boolean::New(env, xf.has_vertical_align != 0));
  out.Set("hasWrapText", Napi::Boolean::New(env, xf.has_wrap_text != 0));
  out.Set("hasJustifyLastLine", Napi::Boolean::New(env, xf.has_justify_last_line != 0));
  if (xf.has_text_rotation != 0) {
    out.Set("textRotation", Napi::Number::New(env, xf.text_rotation));
  }
  if (xf.has_indent != 0) {
    out.Set("indent", Napi::Number::New(env, xf.indent));
  }
  if (xf.has_relative_indent != 0) {
    out.Set("relativeIndent", Napi::Number::New(env, xf.relative_indent));
  }
  if (xf.has_shrink_to_fit != 0) {
    out.Set("shrinkToFit", Napi::Boolean::New(env, xf.shrink_to_fit != 0));
  }
  if (xf.has_reading_order != 0) {
    out.Set("readingOrder", Napi::Number::New(env, xf.reading_order));
  }
  return out;
}

// ---- Conditional formatting (read / mutate) -------------------------

Napi::Value Workbook::GetConditionalFormats(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return arr;
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  std::size_t count = 0;
  if (fm_sheet_cf_count(handle_, sheet, &count) != 0) {
    return arr;
  }
  std::size_t emitted = 0;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cf_rule_t rule{};
    if (fm_sheet_cf_get_at(handle_, sheet, i, &rule) != 0) {
      continue;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("id", Napi::String::New(env, rule.id != nullptr ? rule.id : ""));
    item.Set("type", Napi::Number::New(env, static_cast<uint32_t>(rule.type)));
    item.Set("priority", Napi::Number::New(env, rule.priority));
    item.Set("stopIfTrue", Napi::Boolean::New(env, rule.stop_if_true != 0));
    Napi::Array sqref = Napi::Array::New(env, rule.sqref_count);
    for (uint32_t r = 0; r < rule.sqref_count; ++r) {
      Napi::Object rng = Napi::Object::New(env);
      rng.Set("firstRow", Napi::Number::New(env, rule.sqref[r].first_row));
      rng.Set("firstCol", Napi::Number::New(env, rule.sqref[r].first_col));
      rng.Set("lastRow", Napi::Number::New(env, rule.sqref[r].last_row));
      rng.Set("lastCol", Napi::Number::New(env, rule.sqref[r].last_col));
      sqref.Set(r, rng);
    }
    item.Set("sqref", sqref);
    if (rule.dxf_id_engaged != 0) {
      item.Set("dxfId", Napi::Number::New(env, rule.dxf_id));
    }
    if (rule.formula1 != nullptr) {
      item.Set("formula1", Napi::String::New(env, rule.formula1));
    }
    if (rule.formula2 != nullptr) {
      item.Set("formula2", Napi::String::New(env, rule.formula2));
    }
    if (rule.op_engaged != 0) {
      item.Set("op", Napi::Number::New(env, static_cast<uint32_t>(rule.op)));
    }
    if (rule.rank_engaged != 0) {
      item.Set("rank", Napi::Number::New(env, rule.rank));
      item.Set("percent", Napi::Boolean::New(env, rule.percent != 0));
      item.Set("bottom", Napi::Boolean::New(env, rule.bottom != 0));
    }
    // aboveAverage flags are always engineered but only meaningful for
    // the AboveAverage rule type; surface them only there to mirror the
    // embind shape.
    if (rule.type == 6 /* AboveAverage */) {
      item.Set("aboveAverage", Napi::Boolean::New(env, rule.above_average != 0));
      item.Set("equalAverage", Napi::Boolean::New(env, rule.equal_average != 0));
      if (rule.std_dev_engaged != 0) {
        item.Set("stdDev", Napi::Number::New(env, rule.std_dev));
      }
    }
    if (rule.text != nullptr) {
      item.Set("text", Napi::String::New(env, rule.text));
    }
    if (rule.time_period_engaged != 0) {
      item.Set("timePeriod", Napi::Number::New(env, static_cast<uint32_t>(rule.time_period)));
    }
    if (rule.color_scale_count > 0 && rule.color_scale_thresholds != nullptr && rule.color_scale_colors != nullptr) {
      Napi::Object color_scale = Napi::Object::New(env);
      Napi::Array thresholds = Napi::Array::New(env, rule.color_scale_count);
      Napi::Array colors = Napi::Array::New(env, rule.color_scale_count);
      for (uint32_t j = 0; j < rule.color_scale_count; ++j) {
        thresholds.Set(j, CfvoToJs(env, rule.color_scale_thresholds[j]));
        colors.Set(j, CfColorToJs(env, rule.color_scale_colors[j]));
      }
      color_scale.Set("thresholds", thresholds);
      color_scale.Set("colors", colors);
      item.Set("colorScale", color_scale);
    }
    if (rule.data_bar_engaged != 0) {
      Napi::Object data_bar = Napi::Object::New(env);
      data_bar.Set("min", CfvoToJs(env, rule.data_bar_min));
      data_bar.Set("max", CfvoToJs(env, rule.data_bar_max));
      data_bar.Set("fill", CfColorToJs(env, rule.data_bar_fill));
      data_bar.Set("showValue", Napi::Boolean::New(env, rule.data_bar_show_value != 0));
      data_bar.Set("minLengthPct", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_min_length_pct)));
      data_bar.Set("maxLengthPct", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_max_length_pct)));
      // `x14` extension payload. Each key appears only when the C ABI
      // reports the matching `*_engaged` flag, so an absent key means the
      // rule keeps the model default and the object round-trips through
      // `addConditionalFormat` unchanged. The getter engages all six.
      if (rule.data_bar_gradient_engaged != 0) {
        data_bar.Set("gradient", Napi::Boolean::New(env, rule.data_bar_gradient != 0));
      }
      if (rule.data_bar_axis_position_engaged != 0) {
        data_bar.Set("axisPosition", Napi::Number::New(env, static_cast<uint32_t>(rule.data_bar_axis_position)));
      }
      if (rule.data_bar_negative_fill_engaged != 0) {
        data_bar.Set("negativeFill", CfColorToJs(env, rule.data_bar_negative_fill));
      }
      if (rule.data_bar_border_engaged != 0) {
        data_bar.Set("border", CfColorToJs(env, rule.data_bar_border));
      }
      if (rule.data_bar_negative_border_engaged != 0) {
        data_bar.Set("negativeBorder", CfColorToJs(env, rule.data_bar_negative_border));
      }
      if (rule.data_bar_axis_color_engaged != 0) {
        data_bar.Set("axisColor", CfColorToJs(env, rule.data_bar_axis_color));
      }
      item.Set("dataBar", data_bar);
    }
    if (rule.icon_set_engaged != 0) {
      Napi::Object icon_set = Napi::Object::New(env);
      icon_set.Set("name", Napi::Number::New(env, static_cast<uint32_t>(rule.icon_set_name)));
      Napi::Array thresholds = Napi::Array::New(env, rule.icon_set_threshold_count);
      for (uint32_t j = 0; j < rule.icon_set_threshold_count; ++j) {
        thresholds.Set(j, CfvoToJs(env, rule.icon_set_thresholds[j]));
      }
      icon_set.Set("thresholds", thresholds);
      icon_set.Set("reverse", Napi::Boolean::New(env, rule.icon_set_reverse != 0));
      icon_set.Set("showValue", Napi::Boolean::New(env, rule.icon_set_show_value != 0));
      icon_set.Set("percent", Napi::Boolean::New(env, rule.icon_set_percent != 0));
      item.Set("iconSet", icon_set);
    }
    arr.Set(static_cast<uint32_t>(emitted), item);
    ++emitted;
  }
  return arr;
}

Napi::Value Workbook::AddConditionalFormat(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  if (info.Length() < 2 || !info[0].IsNumber() || !info[1].IsObject()) {
    Napi::TypeError::New(env, "addConditionalFormat expects (sheet:number, rule:object)").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  Napi::Object v = info[1].As<Napi::Object>();

  // Pull every JS field into local storage first; the C ABI receives
  // borrowed `const char*` views that must stay valid until
  // `fm_sheet_cf_add_rule` returns.
  std::vector<fm_cf_cell_range_t> ranges_buf;
  std::vector<fm_cfvo_t> color_scale_thresholds;
  std::vector<fm_cf_color_t> color_scale_colors;
  std::vector<fm_cfvo_t> icon_set_thresholds;
  std::vector<std::string> cfvo_strings;
  if (v.Has("sqref")) {
    Napi::Value sqref_js = v.Get("sqref");
    if (sqref_js.IsArray()) {
      Napi::Array sqref_arr = sqref_js.As<Napi::Array>();
      const uint32_t n = sqref_arr.Length();
      ranges_buf.reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        Napi::Value rng_v = sqref_arr.Get(i);
        if (!rng_v.IsObject()) {
          continue;
        }
        Napi::Object rng = rng_v.As<Napi::Object>();
        fm_cf_cell_range_t r{};
        r.first_row = rng.Get("firstRow").ToNumber().Uint32Value();
        r.first_col = rng.Get("firstCol").ToNumber().Uint32Value();
        r.last_row = rng.Get("lastRow").ToNumber().Uint32Value();
        r.last_col = rng.Get("lastCol").ToNumber().Uint32Value();
        ranges_buf.push_back(r);
      }
    }
  }
  const std::string id = SpecHas(v, "id") ? v.Get("id").ToString().Utf8Value() : std::string();
  const std::string formula1 = SpecHas(v, "formula1") ? v.Get("formula1").ToString().Utf8Value() : std::string();
  const std::string formula2 = SpecHas(v, "formula2") ? v.Get("formula2").ToString().Utf8Value() : std::string();
  const std::string text = SpecHas(v, "text") ? v.Get("text").ToString().Utf8Value() : std::string();

  fm_cf_rule_t rule{};
  rule.id = id.empty() ? nullptr : id.c_str();
  rule.type = static_cast<uint8_t>(SpecPullU32(v, "type", 0U) & 0xFFU);
  rule.priority = SpecPullInt32(v, "priority", 0);
  rule.stop_if_true = SpecPullBool(v, "stopIfTrue", false) ? 1 : 0;
  if (SpecHas(v, "dxfId")) {
    rule.dxf_id_engaged = 1;
    rule.dxf_id = SpecPullU32(v, "dxfId", 0U);
  }
  rule.sqref = ranges_buf.empty() ? nullptr : ranges_buf.data();
  rule.sqref_count = static_cast<uint32_t>(ranges_buf.size());
  rule.formula1 = formula1.empty() ? nullptr : formula1.c_str();
  rule.formula2 = formula2.empty() ? nullptr : formula2.c_str();
  if (SpecHas(v, "op")) {
    rule.op_engaged = 1;
    rule.op = static_cast<uint8_t>(SpecPullU32(v, "op", 0U) & 0xFFU);
  }
  if (SpecHas(v, "rank")) {
    rule.rank_engaged = 1;
    rule.rank = SpecPullInt32(v, "rank", 0);
  }
  rule.percent = SpecPullBool(v, "percent", false) ? 1 : 0;
  rule.bottom = SpecPullBool(v, "bottom", false) ? 1 : 0;
  rule.above_average = SpecPullBool(v, "aboveAverage", true) ? 1 : 0;
  rule.equal_average = SpecPullBool(v, "equalAverage", false) ? 1 : 0;
  if (SpecHas(v, "stdDev")) {
    rule.std_dev_engaged = 1;
    rule.std_dev = SpecPullDouble(v, "stdDev", 0.0);
  }
  rule.text = text.empty() ? nullptr : text.c_str();
  if (SpecHas(v, "timePeriod")) {
    rule.time_period_engaged = 1;
    rule.time_period = static_cast<uint8_t>(SpecPullU32(v, "timePeriod", 0U) & 0xFFU);
  }
  if (SpecHas(v, "colorScale") && v.Get("colorScale").IsObject()) {
    Napi::Object cs = v.Get("colorScale").As<Napi::Object>();
    if (SpecHas(cs, "thresholds") && cs.Get("thresholds").IsArray()) {
      Napi::Array arr = cs.Get("thresholds").As<Napi::Array>();
      color_scale_thresholds.reserve(arr.Length());
      for (uint32_t i = 0; i < arr.Length(); ++i) {
        if (arr.Get(i).IsObject()) {
          color_scale_thresholds.push_back(PullCfvo(arr.Get(i).As<Napi::Object>(), &cfvo_strings));
        }
      }
    }
    if (SpecHas(cs, "colors") && cs.Get("colors").IsArray()) {
      Napi::Array arr = cs.Get("colors").As<Napi::Array>();
      color_scale_colors.reserve(arr.Length());
      for (uint32_t i = 0; i < arr.Length(); ++i) {
        if (arr.Get(i).IsObject()) {
          color_scale_colors.push_back(PullCfColor(arr.Get(i).As<Napi::Object>()));
        }
      }
    }
    rule.color_scale_thresholds = color_scale_thresholds.empty() ? nullptr : color_scale_thresholds.data();
    rule.color_scale_colors = color_scale_colors.empty() ? nullptr : color_scale_colors.data();
    rule.color_scale_count = static_cast<uint32_t>(color_scale_thresholds.size());
  }
  if (SpecHas(v, "dataBar") && v.Get("dataBar").IsObject()) {
    Napi::Object db = v.Get("dataBar").As<Napi::Object>();
    rule.data_bar_engaged = 1;
    if (SpecHas(db, "min") && db.Get("min").IsObject()) {
      rule.data_bar_min = PullCfvo(db.Get("min").As<Napi::Object>(), &cfvo_strings);
    }
    if (SpecHas(db, "max") && db.Get("max").IsObject()) {
      rule.data_bar_max = PullCfvo(db.Get("max").As<Napi::Object>(), &cfvo_strings);
    }
    if (SpecHas(db, "fill") && db.Get("fill").IsObject()) {
      rule.data_bar_fill = PullCfColor(db.Get("fill").As<Napi::Object>());
    }
    rule.data_bar_show_value = SpecPullBool(db, "showValue", true) ? 1 : 0;
    rule.data_bar_min_length_pct = static_cast<uint8_t>(SpecPullU32(db, "minLengthPct", 10U) & 0xFFU);
    rule.data_bar_max_length_pct = static_cast<uint8_t>(SpecPullU32(db, "maxLengthPct", 90U) & 0xFFU);
    // `x14` extension payload. An omitted key leaves the `*_engaged` flag
    // clear, which the C ABI reads as "keep the model default" (gradient
    // on, automatic axis, negative fill equal to the positive fill, no
    // border, black axis).
    if (SpecHas(db, "gradient")) {
      rule.data_bar_gradient_engaged = 1;
      rule.data_bar_gradient = SpecPullBool(db, "gradient", true) ? 1 : 0;
    }
    if (SpecHas(db, "axisPosition")) {
      rule.data_bar_axis_position_engaged = 1;
      rule.data_bar_axis_position = static_cast<uint8_t>(SpecPullU32(db, "axisPosition", 0U) & 0xFFU);
    }
    if (SpecHas(db, "negativeFill") && db.Get("negativeFill").IsObject()) {
      rule.data_bar_negative_fill_engaged = 1;
      rule.data_bar_negative_fill = PullCfColor(db.Get("negativeFill").As<Napi::Object>());
    }
    if (SpecHas(db, "border") && db.Get("border").IsObject()) {
      rule.data_bar_border_engaged = 1;
      rule.data_bar_border = PullCfColor(db.Get("border").As<Napi::Object>());
    }
    if (SpecHas(db, "negativeBorder") && db.Get("negativeBorder").IsObject()) {
      rule.data_bar_negative_border_engaged = 1;
      rule.data_bar_negative_border = PullCfColor(db.Get("negativeBorder").As<Napi::Object>());
    }
    if (SpecHas(db, "axisColor") && db.Get("axisColor").IsObject()) {
      rule.data_bar_axis_color_engaged = 1;
      rule.data_bar_axis_color = PullCfColor(db.Get("axisColor").As<Napi::Object>());
    }
  }
  if (SpecHas(v, "iconSet") && v.Get("iconSet").IsObject()) {
    Napi::Object is = v.Get("iconSet").As<Napi::Object>();
    rule.icon_set_engaged = 1;
    rule.icon_set_name = static_cast<uint8_t>(SpecPullU32(is, "name", 0U) & 0xFFU);
    if (SpecHas(is, "thresholds") && is.Get("thresholds").IsArray()) {
      Napi::Array arr = is.Get("thresholds").As<Napi::Array>();
      icon_set_thresholds.reserve(arr.Length());
      for (uint32_t i = 0; i < arr.Length(); ++i) {
        if (arr.Get(i).IsObject()) {
          icon_set_thresholds.push_back(PullCfvo(arr.Get(i).As<Napi::Object>(), &cfvo_strings));
        }
      }
    }
    rule.icon_set_thresholds = icon_set_thresholds.empty() ? nullptr : icon_set_thresholds.data();
    rule.icon_set_threshold_count = static_cast<uint32_t>(icon_set_thresholds.size());
    rule.icon_set_reverse = SpecPullBool(is, "reverse", false) ? 1 : 0;
    rule.icon_set_show_value = SpecPullBool(is, "showValue", true) ? 1 : 0;
    rule.icon_set_percent = SpecPullBool(is, "percent", true) ? 1 : 0;
  }
  std::size_t new_index = 0;
  fm_status_t rc = fm_sheet_cf_add_rule(handle_, sheet, rule, &new_index);
  if (rc != 0) {
    return MakeNumberFieldResult(env, MakeErrorStatus(env, rc), "index", 0);
  }
  return MakeNumberFieldResult(env, MakeOkStatus(env), "index", static_cast<double>(new_index));
}

Napi::Value Workbook::RemoveConditionalFormatAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  const std::size_t index = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_sheet_cf_remove_at(handle_, sheet, index);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::ClearConditionalFormats(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::size_t sheet = static_cast<std::size_t>(ArgU32(info, 0));
  fm_status_t rc = fm_sheet_cf_clear(handle_, sheet);
  return MakeStatus(env, rc);
}

}  // namespace formulon_node
