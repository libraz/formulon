// Style-table bindings: read / write entries in the xf / font / fill /
// border / numFmt pools plus the per-cell xf-index getters and the named
// cell-style accessors.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {
namespace {

/// Reads the `{kind, rgb, theme, tint, indexed}` colour specification out
/// of `owner[key]`. An absent object leaves `kind` at `kFmColorNone`, so
/// the writer emits the sibling `*Argb` as literal `rgb`. A supplied
/// selector is authoritative; this binding does not resolve theme/indexed /
/// auto colours.
fm_color_spec PullColorSpec(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  fm_color_spec spec{};
  Napi::Object v;
  if (!reader.Object(owner, key, &v)) {
    return spec;
  }
  spec.kind = reader.U8(v, "kind", 0U);
  spec.rgb = reader.U32(v, "rgb", 0U);
  spec.theme = reader.U32(v, "theme", 0U);
  spec.tint = reader.Double(v, "tint", 0.0);
  spec.indexed = reader.U32(v, "indexed", 0U);
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
  out.Set("name", JsString(env, f.name));
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
  out.Set("scheme", Napi::Number::New(env, static_cast<uint32_t>(f.scheme)));
  out.Set("color", ColorSpecToJs(env, f.color));
  return out;
}

/// Reads a font record out of a JS object. `name_storage` owns the font
/// name for the duration of the C ABI call, which borrows the pointer.
void PullFontRecord(CheckedSpecReader& reader, const Napi::Object& record, std::string* name_storage,
                    fm_font_record* out) {
  reader.String(record, "name", name_storage);
  out->name = name_storage->c_str();
  out->size = reader.Double(record, "size", 11.0);
  out->bold = reader.Bool(record, "bold", false) ? 1 : 0;
  out->italic = reader.Bool(record, "italic", false) ? 1 : 0;
  out->strike = reader.Bool(record, "strike", false) ? 1 : 0;
  out->has_bold = reader.Bool(record, "hasBold", false) ? 1 : 0;
  out->has_italic = reader.Bool(record, "hasItalic", false) ? 1 : 0;
  out->has_strike = reader.Bool(record, "hasStrike", false) ? 1 : 0;
  out->underline = reader.U8(record, "underline", 0U);
  out->vert_align = reader.U8(record, "vertAlign", 0U);
  out->has_family = reader.Bool(record, "hasFamily", false) ? 1 : 0;
  out->family = reader.U8(record, "family", 0U);
  out->has_charset = reader.Bool(record, "hasCharset", false) ? 1 : 0;
  out->charset = reader.U8(record, "charset", 0U);
  out->scheme = reader.U8(record, "scheme", 0U);
  out->color_argb = reader.U32(record, "colorArgb", 0xFF000000U);
  out->color = PullColorSpec(reader, record, "color");
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

fm_fill_record PullFillRecord(CheckedSpecReader& reader, const Napi::Object& record) {
  fm_fill_record out{};
  out.pattern = reader.U8(record, "pattern", 0U);
  out.fg_argb = reader.U32(record, "fgArgb", 0U);
  out.bg_argb = reader.U32(record, "bgArgb", 0U);
  out.fg = PullColorSpec(reader, record, "fg");
  out.bg = PullColorSpec(reader, record, "bg");
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

fm_border_side PullBorderSide(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  fm_border_side s{};
  Napi::Object v;
  if (!reader.Object(owner, key, &v)) {
    return s;
  }
  s.style = reader.U8(v, "style", 0U);
  s.color_argb = reader.U32(v, "colorArgb", 0U);
  s.color = PullColorSpec(reader, v, "color");
  return s;
}

fm_border_record PullBorderRecord(CheckedSpecReader& reader, const Napi::Object& record) {
  fm_border_record out{};
  out.left = PullBorderSide(reader, record, "left");
  out.right = PullBorderSide(reader, record, "right");
  out.top = PullBorderSide(reader, record, "top");
  out.bottom = PullBorderSide(reader, record, "bottom");
  out.diagonal = PullBorderSide(reader, record, "diagonal");
  out.diagonal_up = reader.Bool(record, "diagonalUp", false) ? 1 : 0;
  out.diagonal_down = reader.Bool(record, "diagonalDown", false) ? 1 : 0;
  return out;
}

/// Builds the JS mirror of a cell format. Shared by `getCellXf` and
/// `getCellStyleXf`; a zero-initialised `xf` yields the all-default record
/// those two return on their failure paths. `xfId` is set by the callers:
/// a named-style xf is its own parent, so `getCellStyleXf` always reports
/// 0 there.
Napi::Object CellXfToJs(Napi::Env env, const fm_cell_xf& xf) {
  Napi::Object out = Napi::Object::New(env);
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
  out.Set("applyNumberFormat", Napi::Boolean::New(env, xf.apply_number_format != 0));
  out.Set("applyFont", Napi::Boolean::New(env, xf.apply_font != 0));
  out.Set("applyFill", Napi::Boolean::New(env, xf.apply_fill != 0));
  out.Set("applyBorder", Napi::Boolean::New(env, xf.apply_border != 0));
  out.Set("applyAlignment", Napi::Boolean::New(env, xf.apply_alignment != 0));
  out.Set("applyProtection", Napi::Boolean::New(env, xf.apply_protection != 0));
  out.Set("quotePrefix", Napi::Boolean::New(env, xf.quote_prefix != 0));
  out.Set("hasProtection", Napi::Boolean::New(env, xf.has_protection != 0));
  // Without a `<protection>` child the model default applies: locked, not hidden.
  out.Set("locked", Napi::Boolean::New(env, xf.has_protection == 0 || xf.locked != 0));
  out.Set("hidden", Napi::Boolean::New(env, xf.has_protection != 0 && xf.hidden != 0));
  return out;
}

void SetCellXfAlignment(Napi::Env env, const fm_cell_xf& xf, Napi::Object* out) {
  if (xf.has_text_rotation != 0) {
    out->Set("textRotation", Napi::Number::New(env, xf.text_rotation));
  }
  if (xf.has_indent != 0) {
    out->Set("indent", Napi::Number::New(env, xf.indent));
  }
  if (xf.has_relative_indent != 0) {
    out->Set("relativeIndent", Napi::Number::New(env, xf.relative_indent));
  }
  if (xf.has_shrink_to_fit != 0) {
    out->Set("shrinkToFit", Napi::Boolean::New(env, xf.shrink_to_fit != 0));
  }
  if (xf.has_reading_order != 0) {
    out->Set("readingOrder", Napi::Number::New(env, xf.reading_order));
  }
}

/// Builds the `getCellXf` / `getCellStyleXf` result from the C call's
/// status and record; a failed read reports the all-default record.
Napi::Object CellXfResultToJs(Napi::Env env, fm_status_t rc, fm_cell_xf xf) {
  if (rc != 0) {
    xf = fm_cell_xf{};
  }
  Napi::Object out = CellXfToJs(env, xf);
  out.Set("status", MakeStatus(env, rc));
  out.Set("xfId", Napi::Number::New(env, xf.xf_id));
  SetCellXfAlignment(env, xf, &out);
  return out;
}

// Shared body of the style-pool count getters that take no argument.
using StyleCountFn = fm_status_t (*)(fm_workbook_t*, uint32_t*);

Napi::Value StyleCountResult(const Napi::CallbackInfo& info, fm_workbook_t* handle, StyleCountFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeNumberResult(env, kBindingInvalidHandle, 0);
  }
  uint32_t n = 0;
  const fm_status_t rc = fn(handle, &n);
  return MakeNumberResult(env, rc, n);
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
  return MakeNumberFieldResult(env, rc, "xfIndex", xf);
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

Napi::Value Workbook::SetRangeXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_sheet_set_range_xf_index(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                                     ArgU32(info, 3), ArgU32(info, 4), ArgU32(info, 5)));
}

// ---- Style getters --------------------------------------------------

// Every getter below emits its declared payload keys on all three exit
// paths -- dead handle, non-zero rc, success -- so `status.ok` is the only
// thing a caller has to branch on, and the WASM binding does the same for
// the same call. A key present only on success makes the TypeScript
// declaration a lie and turns a missed status check into a `TypeError`.

Napi::Value Workbook::GetCellXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t xf_index = ArgU32(info, 0);
  fm_cell_xf xf{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_cell_xf(handle_, xf_index, &xf) : kBindingInvalidHandle;
  return CellXfResultToJs(env, rc, xf);
}

Napi::Value Workbook::GetFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t font_index = ArgU32(info, 0);
  fm_font_record f{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_font(handle_, font_index, &f) : kBindingInvalidHandle;
  if (rc != 0) {
    f = fm_font_record{};
  }
  Napi::Object out = FontRecordToJs(env, f);
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::GetFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t fill_index = ArgU32(info, 0);
  fm_fill_record f{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_fill(handle_, fill_index, &f) : kBindingInvalidHandle;
  if (rc != 0) {
    f = fm_fill_record{};
  }
  Napi::Object out = FillRecordToJs(env, f);
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::GetBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t border_index = ArgU32(info, 0);
  fm_border_record b{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_border(handle_, border_index, &b) : kBindingInvalidHandle;
  if (rc != 0) {
    b = fm_border_record{};
  }
  Napi::Object out = BorderRecordToJs(env, b);
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::GetNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  const uint32_t num_fmt_id = ArgU32(info, 0);
  const char* s = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_styles_get_num_fmt_string(handle_, static_cast<uint16_t>(num_fmt_id), &s)
                             : kBindingInvalidHandle;
  out.Set("status", MakeStatus(env, rc));
  out.Set("numFmtId", Napi::Number::New(env, rc == 0 ? num_fmt_id : 0U));
  out.Set("formatCode", JsString(env, rc == 0 ? s : nullptr));
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
    num_fmt.Set("formatCode", JsString(env, d.num_fmt_code));
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
  Napi::Object record = ArgObjectOrEmpty(info, 0);
  CheckedSpecReader reader(env);
  std::string name;
  fm_font_record fr{};
  PullFontRecord(reader, record, &name, &fr);
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it. The exception surfaces to JS on
    // return.
    return env.Undefined();
  }
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_font(handle_, fr, &idx);
  return MakeNumberFieldResult(env, rc, "index", idx);
}

Napi::Value Workbook::SetFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t font_index = ArgU32(info, 0);
  Napi::Object record = ArgObjectOrEmpty(info, 1);
  CheckedSpecReader reader(env);
  std::string name;
  fm_font_record fr{};
  PullFontRecord(reader, record, &name, &fr);
  if (!reader.ok()) {
    // See the matching guard in AddFont.
    return env.Undefined();
  }
  return MakeStatus(env, fm_styles_set_font(handle_, font_index, fr));
}

Napi::Value Workbook::SetDefaultFont(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);
  CheckedSpecReader reader(env);
  std::string name;
  fm_font_record fr{};
  PullFontRecord(reader, record, &name, &fr);
  if (!reader.ok()) {
    // See the matching guard in AddFont.
    return env.Undefined();
  }
  return MakeStatus(env, fm_workbook_set_default_font(handle_, fr));
}

Napi::Value Workbook::AddFill(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);
  CheckedSpecReader reader(env);
  const fm_fill_record fr = PullFillRecord(reader, record);
  if (!reader.ok()) {
    // See the matching guard in AddFont. The return value is discarded
    // in favor of the pending exception either way.
    return env.Undefined();
  }
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_fill(handle_, fr, &idx);
  return MakeNumberFieldResult(env, rc, "index", idx);
}

Napi::Value Workbook::AddBorder(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);
  CheckedSpecReader reader(env);
  const fm_border_record br = PullBorderRecord(reader, record);
  if (!reader.ok()) {
    // See the matching guard in AddFont. The return value is discarded
    // in favor of the pending exception either way.
    return env.Undefined();
  }
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_border(handle_, br, &idx);
  return MakeNumberFieldResult(env, rc, "index", idx);
}

Napi::Value Workbook::AddNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "numFmtId", 0);
  }
  const std::string code = ArgString(info, 0);
  uint16_t id = 0;
  fm_status_t rc = fm_styles_add_num_fmt(handle_, code.c_str(), &id);
  return MakeNumberFieldResult(env, rc, "numFmtId", static_cast<uint32_t>(id));
}

Napi::Value Workbook::AddXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);
  CheckedSpecReader reader(env);
  fm_cell_xf xf{};
  xf.font_index = reader.U32(record, "fontIndex", 0U);
  xf.fill_index = reader.U32(record, "fillIndex", 0U);
  xf.border_index = reader.U32(record, "borderIndex", 0U);
  xf.num_fmt_id = reader.U16(record, "numFmtId", 0U);
  bool horizontal_align_present = false;
  bool vertical_align_present = false;
  bool wrap_text_present = false;
  bool justify_last_line_present = false;
  xf.horizontal_align = reader.U8(record, "horizontalAlign", 0U, &horizontal_align_present);
  xf.vertical_align = reader.U8(record, "verticalAlign", 2U, &vertical_align_present);
  xf.wrap_text = reader.Bool(record, "wrapText", false, &wrap_text_present) ? 1 : 0;
  xf.justify_last_line = reader.Bool(record, "justifyLastLine", false, &justify_last_line_present) ? 1 : 0;
  xf.xf_id = reader.U32(record, "xfId", 0U);
  bool has_horizontal_align_present = false;
  bool has_vertical_align_present = false;
  bool has_wrap_text_present = false;
  bool has_justify_last_line_present = false;
  const bool has_horizontal_align_value =
      reader.Bool(record, "hasHorizontalAlign", false, &has_horizontal_align_present);
  const bool has_vertical_align_value = reader.Bool(record, "hasVerticalAlign", false, &has_vertical_align_present);
  const bool has_wrap_text_value = reader.Bool(record, "hasWrapText", false, &has_wrap_text_present);
  const bool has_justify_last_line_value =
      reader.Bool(record, "hasJustifyLastLine", false, &has_justify_last_line_present);
  xf.has_horizontal_align =
      has_horizontal_align_present ? (has_horizontal_align_value ? 1 : 0) : (horizontal_align_present ? 1 : 0);
  xf.has_vertical_align =
      has_vertical_align_present ? (has_vertical_align_value ? 1 : 0) : (vertical_align_present ? 1 : 0);
  xf.has_wrap_text = has_wrap_text_present ? (has_wrap_text_value ? 1 : 0) : (wrap_text_present ? 1 : 0);
  xf.has_justify_last_line =
      has_justify_last_line_present ? (has_justify_last_line_value ? 1 : 0) : (justify_last_line_present ? 1 : 0);
  bool has_alignment_present = false;
  const bool has_alignment_value = reader.Bool(record, "hasAlignment", false, &has_alignment_present);
  bool text_rotation_present = false;
  bool indent_present = false;
  bool relative_indent_present = false;
  bool shrink_to_fit_present = false;
  bool reading_order_present = false;
  xf.text_rotation = reader.U32(record, "textRotation", 0U, &text_rotation_present);
  xf.indent = reader.U32(record, "indent", 0U, &indent_present);
  xf.relative_indent = reader.I32(record, "relativeIndent", 0, &relative_indent_present);
  xf.shrink_to_fit = reader.Bool(record, "shrinkToFit", false, &shrink_to_fit_present) ? 1 : 0;
  xf.reading_order = reader.U32(record, "readingOrder", 0U, &reading_order_present);
  xf.has_text_rotation = text_rotation_present ? 1 : 0;
  xf.has_indent = indent_present ? 1 : 0;
  xf.has_relative_indent = relative_indent_present ? 1 : 0;
  xf.has_shrink_to_fit = shrink_to_fit_present ? 1 : 0;
  xf.has_reading_order = reading_order_present ? 1 : 0;
  const bool has_supplied_alignment = horizontal_align_present || vertical_align_present || wrap_text_present ||
                                      justify_last_line_present || text_rotation_present || indent_present ||
                                      relative_indent_present || shrink_to_fit_present || reading_order_present ||
                                      has_horizontal_align_present || has_vertical_align_present ||
                                      has_wrap_text_present || has_justify_last_line_present;
  xf.has_alignment = has_alignment_present ? (has_alignment_value ? 1 : 0) : (has_supplied_alignment ? 1 : 0);
  xf.apply_number_format = reader.Bool(record, "applyNumberFormat", false) ? 1 : 0;
  xf.apply_font = reader.Bool(record, "applyFont", false) ? 1 : 0;
  xf.apply_fill = reader.Bool(record, "applyFill", false) ? 1 : 0;
  xf.apply_border = reader.Bool(record, "applyBorder", false) ? 1 : 0;
  xf.apply_alignment = reader.Bool(record, "applyAlignment", false) ? 1 : 0;
  xf.apply_protection = reader.Bool(record, "applyProtection", false) ? 1 : 0;
  xf.quote_prefix = reader.Bool(record, "quotePrefix", false) ? 1 : 0;
  bool has_protection_present = false;
  bool locked_present = false;
  bool hidden_present = false;
  const bool has_protection_value = reader.Bool(record, "hasProtection", false, &has_protection_present);
  const bool locked_value = reader.Bool(record, "locked", true, &locked_present);
  const bool hidden_value = reader.Bool(record, "hidden", false, &hidden_present);
  xf.has_protection =
      has_protection_present ? (has_protection_value ? 1 : 0) : ((locked_present || hidden_present) ? 1 : 0);
  xf.locked = locked_value ? 1 : 0;
  xf.hidden = hidden_value ? 1 : 0;
  if (!reader.ok()) {
    // See the matching guard in AddFont. The return value is discarded
    // in favor of the pending exception either way.
    return env.Undefined();
  }
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_cell_xf(handle_, xf, &idx);
  return MakeNumberFieldResult(env, rc, "index", idx);
}

Napi::Value Workbook::AddDxf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeNumberFieldResult(env, NullHandleError(env), "index", 0);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);

  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  fm_dxf_record dxf{};
  CheckedSpecReader reader(env);

  Napi::Object font;
  if (reader.Object(record, "font", &font)) {
    dxf.font_engaged = 1;
    PullFontRecord(reader, font, &font_name, &dxf.font);
  }

  Napi::Object fill;
  if (reader.Object(record, "fill", &fill)) {
    dxf.fill_engaged = 1;
    dxf.fill = PullFillRecord(reader, fill);
  }

  Napi::Object border;
  if (reader.Object(record, "border", &border)) {
    dxf.border_engaged = 1;
    dxf.border = PullBorderRecord(reader, border);
  }

  Napi::Object num_fmt;
  if (reader.Object(record, "numFmt", &num_fmt)) {
    dxf.num_fmt_engaged = 1;
    dxf.num_fmt_id = reader.U16(num_fmt, "numFmtId", 0U);
    reader.String(num_fmt, "formatCode", &num_fmt_code);
    dxf.num_fmt_code = num_fmt_code.c_str();
  }

  reader.String(record, "alignmentXml", &alignment_xml);
  reader.String(record, "protectionXml", &protection_xml);
  dxf.alignment_xml = alignment_xml.c_str();
  dxf.protection_xml = protection_xml.c_str();

  if (!reader.ok()) {
    // See the matching guard in AddFont. The return value is discarded
    // in favor of the pending exception either way.
    return env.Undefined();
  }
  uint32_t idx = 0;
  fm_status_t rc = fm_styles_add_dxf(handle_, dxf, &idx);
  return MakeNumberFieldResult(env, rc, "index", idx);
}

// ---- Style pool counts ----------------------------------------------
//
// `FontCount` / `FillCount` / `BorderCount` / `XfCount` are now emitted
// by the binding codegen (see `src/node_addon/generated/styles_counts.cc`).

Napi::Value Workbook::DxfCount(const Napi::CallbackInfo& info) {
  return StyleCountResult(info, handle_, &fm_styles_get_dxf_count);
}

// ---- Named cell styles ----------------------------------------------

Napi::Value Workbook::CellStyleCount(const Napi::CallbackInfo& info) {
  return StyleCountResult(info, handle_, &fm_styles_get_cell_style_count);
}

Napi::Value Workbook::CellStyleXfCount(const Napi::CallbackInfo& info) {
  return StyleCountResult(info, handle_, &fm_styles_get_cell_style_xf_count);
}

Napi::Value Workbook::GetCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  const uint32_t index = ArgU32(info, 0);
  fm_cell_style_record_t cs{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_cell_style(handle_, index, &cs) : kBindingInvalidHandle;
  if (rc != 0) {
    cs = fm_cell_style_record_t{};
  }
  out.Set("status", MakeStatus(env, rc));
  out.Set("name", JsString(env, cs.name));
  out.Set("xfId", Napi::Number::New(env, cs.xf_id));
  out.Set("builtinId", Napi::Number::New(env, cs.builtin_id));
  out.Set("iLevel", Napi::Number::New(env, cs.i_level));
  out.Set("hidden", Napi::Boolean::New(env, cs.hidden != 0));
  out.Set("customBuiltin", Napi::Boolean::New(env, cs.custom_builtin != 0));
  return out;
}

Napi::Value Workbook::GetCellStyleXf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const uint32_t index = ArgU32(info, 0);
  fm_cell_xf xf{};
  const fm_status_t rc = handle_ != nullptr ? fm_styles_get_cell_style_xf(handle_, index, &xf) : kBindingInvalidHandle;
  return CellXfResultToJs(env, rc, xf);
}

Napi::Value Workbook::SetCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);
  CheckedSpecReader reader(env);
  std::string name;
  reader.String(record, "name", &name);
  fm_cell_style_record_t cs{};
  cs.name = name.c_str();
  cs.xf_id = reader.U32(record, "xfId", 0U);
  cs.builtin_id = reader.U32(record, "builtinId", FM_CELL_STYLE_BUILTIN_ID_NONE);
  cs.i_level = reader.U32(record, "iLevel", 0U);
  cs.hidden = reader.Bool(record, "hidden", false) ? 1 : 0;
  cs.custom_builtin = reader.Bool(record, "customBuiltin", false) ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_styles_set_cell_style(handle_, &cs));
}

Napi::Value Workbook::RemoveCellStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const std::string name = ArgString(info, 0);
  return MakeStatus(env, fm_styles_remove_cell_style(handle_, name.c_str()));
}

// ---- Theme, colour resolution, effective style ----------------------

Napi::Value Workbook::GetTheme(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_theme_colors colors{};
  fm_theme_fonts fonts{};
  int32_t source = 0;
  fm_status_t rc = kBindingInvalidHandle;
  if (handle_ != nullptr) {
    rc = fm_workbook_get_theme_colors(handle_, &colors, &source);
    if (rc == 0) {
      rc = fm_workbook_get_theme_fonts(handle_, &fonts);
    }
  }
  if (rc != 0) {
    colors = fm_theme_colors{};
    fonts = fm_theme_fonts{};
    source = 0;
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
  out.Set("status", MakeStatus(env, rc));
  out.Set("source", Napi::Number::New(env, source));
  out.Set("colors", arr);
  out.Set("fonts", fontsOut);
  return out;
}

Napi::Value Workbook::SetThemeColors(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 1 || !info[0].IsArray() || info[0].As<Napi::Array>().Length() != 12) {
    return MakeBindingArgumentError(env, "setThemeColors expects an array of 12 ARGB numbers");
  }
  Napi::Array arr = info[0].As<Napi::Array>();
  CheckedSpecReader reader(env);
  fm_theme_colors colors{};
  for (uint32_t i = 0; i < 12; ++i) {
    const std::string key = std::to_string(i);
    colors.argb[i] = reader.U32(arr, key.c_str(), 0U);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_workbook_set_theme_colors(handle_, &colors));
}

Napi::Value Workbook::SetThemeFonts(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object record = ArgObjectOrEmpty(info, 0);
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
  fm_theme_fonts fonts{};
  fonts.major_latin = major_latin.c_str();
  fonts.major_east_asian = major_ea.c_str();
  fonts.minor_latin = minor_latin.c_str();
  fonts.minor_east_asian = minor_ea.c_str();
  return MakeStatus(env, fm_workbook_set_theme_fonts(handle_, &fonts));
}

Napi::Value Workbook::ResetTheme(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(env, fm_workbook_reset_theme(handle_));
}

Napi::Value Workbook::ResolveColor(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  CheckedSpecReader reader(env);
  fm_color_spec spec{};
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
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_resolve_color(handle_, spec, ArgI32(info, 1), &argb, &resolution)
                             : kBindingInvalidHandle;
  if (rc != 0) {
    argb = 0;
    resolution = 0;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set("argb", Napi::Number::New(env, argb));
  out.Set("resolution", Napi::Number::New(env, resolution));
  return out;
}

Napi::Value Workbook::GetEffectiveStyle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  fm_effective_style es{};
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_effective_style(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &es)
                         : kBindingInvalidHandle;
  if (rc != 0) {
    es = fm_effective_style{};
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
  out.Set("status", MakeStatus(env, rc));
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
