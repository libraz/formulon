// Style-pool bindings: font / fill / border / numFmt / xf / dxf records and
// the per-cell xf-index accessors.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/styles_internal.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

using detail::CellXfResultToJs;
using detail::StyleCountResult;

namespace {

/// Builds the JS mirror of a font record. Shared by `getFont` and the
/// `<dxf>` font projection so both surface the same field set. An absent
/// element reads as `null`.
Napi::Object FontRecordToJs(Napi::Env env, const fm_font_record_t& f) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsString(env, f.name));
  out.Set("size", JsNumberOrNull(env, f.has_size, f.size));
  out.Set("colorArgb", Napi::Number::New(env, f.color_argb));
  out.Set("bold", JsBoolOrNull(env, f.has_bold, f.bold));
  out.Set("italic", JsBoolOrNull(env, f.has_italic, f.italic));
  out.Set("strike", JsBoolOrNull(env, f.has_strike, f.strike));
  out.Set("underline", Napi::Number::New(env, static_cast<uint32_t>(f.underline)));
  out.Set("vertAlign", Napi::Number::New(env, static_cast<uint32_t>(f.vert_align)));
  out.Set("family", JsNumberOrNull(env, f.has_family, static_cast<uint32_t>(f.family)));
  out.Set("charset", JsNumberOrNull(env, f.has_charset, static_cast<uint32_t>(f.charset)));
  out.Set("scheme", Napi::Number::New(env, static_cast<uint32_t>(f.scheme)));
  out.Set("outline", JsBoolOrNull(env, f.has_outline, f.outline));
  out.Set("shadow", JsBoolOrNull(env, f.has_shadow, f.shadow));
  out.Set("condense", JsBoolOrNull(env, f.has_condense, f.condense));
  out.Set("extend", JsBoolOrNull(env, f.has_extend, f.extend));
  out.Set("color", ColorSpecToJs(env, f.color));
  return out;
}

/// Reads a font record out of a JS object. `name_storage` owns the font
/// name for the duration of the C ABI call, which borrows the pointer. An
/// omitted or `null` element is absent from the record.
void PullFontRecord(CheckedSpecReader& reader, const Napi::Object& record, std::string* name_storage,
                    fm_font_record_t* out) {
  bool present = false;
  const auto flag = [&](const char* key, int32_t* has, int32_t* value) {
    *value = reader.Bool(record, key, false, &present) ? 1 : 0;
    *has = present ? 1 : 0;
  };
  reader.String(record, "name", name_storage);
  out->name = name_storage->c_str();
  out->size = reader.Double(record, "size", 11.0, &present);
  out->has_size = present ? 1 : 0;
  flag("bold", &out->has_bold, &out->bold);
  flag("italic", &out->has_italic, &out->italic);
  flag("strike", &out->has_strike, &out->strike);
  out->underline = reader.U8(record, "underline", 0U);
  out->vert_align = reader.U8(record, "vertAlign", 0U);
  out->family = reader.U8(record, "family", 0U, &present);
  out->has_family = present ? 1 : 0;
  out->charset = reader.U8(record, "charset", 0U, &present);
  out->has_charset = present ? 1 : 0;
  out->scheme = reader.U8(record, "scheme", 0U);
  flag("outline", &out->has_outline, &out->outline);
  flag("shadow", &out->has_shadow, &out->shadow);
  flag("condense", &out->has_condense, &out->condense);
  flag("extend", &out->has_extend, &out->extend);
  out->color_argb = reader.U32(record, "colorArgb", 0xFF000000U);
  out->color = PullColorSpec(reader, record, "color");
}

Napi::Object FillRecordToJs(Napi::Env env, const fm_fill_record_t& f) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("pattern", Napi::Number::New(env, static_cast<uint32_t>(f.pattern)));
  out.Set("fgArgb", Napi::Number::New(env, f.fg_argb));
  out.Set("bgArgb", Napi::Number::New(env, f.bg_argb));
  out.Set("fg", ColorSpecToJs(env, f.fg));
  out.Set("bg", ColorSpecToJs(env, f.bg));
  if (f.has_gradient != 0) {
    Napi::Object gradient = Napi::Object::New(env);
    gradient.Set("type", Napi::Number::New(env, f.gradient_type));
    gradient.Set("degree", Napi::Number::New(env, f.gradient_degree));
    gradient.Set("left", Napi::Number::New(env, f.gradient_left));
    gradient.Set("right", Napi::Number::New(env, f.gradient_right));
    gradient.Set("top", Napi::Number::New(env, f.gradient_top));
    gradient.Set("bottom", Napi::Number::New(env, f.gradient_bottom));
    Napi::Array stops = Napi::Array::New(env, f.gradient_stop_count);
    for (std::size_t i = 0; i < f.gradient_stop_count; ++i) {
      Napi::Object stop = Napi::Object::New(env);
      stop.Set("position", Napi::Number::New(env, f.gradient_stops[i].position));
      stop.Set("colorArgb", Napi::Number::New(env, f.gradient_stops[i].color_argb));
      stop.Set("color", ColorSpecToJs(env, f.gradient_stops[i].color));
      stops.Set(static_cast<uint32_t>(i), stop);
    }
    gradient.Set("stops", stops);
    out.Set("gradient", gradient);
  }
  return out;
}

/// Reads a fill record out of a JS object. `stops` owns the gradient stops
/// for the duration of the C ABI call, which borrows the pointer.
fm_fill_record_t PullFillRecord(CheckedSpecReader& reader, const Napi::Object& record,
                                std::vector<fm_gradient_stop_t>* stops) {
  fm_fill_record_t out{};
  out.pattern = reader.U8(record, "pattern", 0U);
  out.fg_argb = reader.U32(record, "fgArgb", 0U);
  out.bg_argb = reader.U32(record, "bgArgb", 0U);
  out.fg = PullColorSpec(reader, record, "fg");
  out.bg = PullColorSpec(reader, record, "bg");
  Napi::Object gradient;
  if (reader.Object(record, "gradient", &gradient)) {
    out.has_gradient = 1;
    out.gradient_type = static_cast<int32_t>(reader.U32(gradient, "type", 0U));
    out.gradient_degree = reader.Double(gradient, "degree", 0.0);
    out.gradient_left = reader.Double(gradient, "left", 0.0);
    out.gradient_right = reader.Double(gradient, "right", 0.0);
    out.gradient_top = reader.Double(gradient, "top", 0.0);
    out.gradient_bottom = reader.Double(gradient, "bottom", 0.0);
    Napi::Array list;
    if (reader.Array(gradient, "stops", &list)) {
      const uint32_t n = list.Length();
      stops->reserve(n);
      for (uint32_t i = 0; i < n; ++i) {
        Napi::Value element;
        if (!reader.ArrayElement(list, i, &element)) {
          continue;
        }
        Napi::Object item;
        if (!reader.Object(element, "stops[]", &item)) {
          continue;
        }
        fm_gradient_stop_t stop{};
        stop.position = reader.Double(item, "position", 0.0);
        stop.color_argb = reader.U32(item, "colorArgb", 0U);
        stop.color = PullColorSpec(reader, item, "color");
        stops->push_back(stop);
      }
    }
    out.gradient_stops = stops->empty() ? nullptr : stops->data();
    out.gradient_stop_count = stops->size();
  }
  return out;
}

Napi::Object BorderSideToJs(Napi::Env env, const fm_border_side_t& s) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("style", Napi::Number::New(env, static_cast<uint32_t>(s.style)));
  out.Set("colorArgb", Napi::Number::New(env, s.color_argb));
  out.Set("color", ColorSpecToJs(env, s.color));
  return out;
}

Napi::Object BorderRecordToJs(Napi::Env env, const fm_border_record_t& b) {
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

fm_border_side_t PullBorderSide(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  fm_border_side_t s{};
  Napi::Object v;
  if (!reader.Object(owner, key, &v)) {
    return s;
  }
  s.style = reader.U8(v, "style", 0U);
  s.color_argb = reader.U32(v, "colorArgb", 0U);
  s.color = PullColorSpec(reader, v, "color");
  return s;
}

fm_border_record_t PullBorderRecord(CheckedSpecReader& reader, const Napi::Object& record) {
  fm_border_record_t out{};
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
Napi::Object CellXfToJs(Napi::Env env, const fm_cell_xf_t& xf) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("fontIndex", Napi::Number::New(env, xf.font_index));
  out.Set("fillIndex", Napi::Number::New(env, xf.fill_index));
  out.Set("borderIndex", Napi::Number::New(env, xf.border_index));
  out.Set("numFmtId", Napi::Number::New(env, static_cast<uint32_t>(xf.num_fmt_id)));
  out.Set("horizontalAlign",
          JsNumberOrNull(env, xf.has_horizontal_align, static_cast<uint32_t>(xf.horizontal_align)));
  out.Set("verticalAlign", JsNumberOrNull(env, xf.has_vertical_align, static_cast<uint32_t>(xf.vertical_align)));
  out.Set("wrapText", JsBoolOrNull(env, xf.has_wrap_text, xf.wrap_text));
  out.Set("justifyLastLine", JsBoolOrNull(env, xf.has_justify_last_line, xf.justify_last_line));
  out.Set("applyNumberFormat", Napi::Boolean::New(env, xf.apply_number_format != 0));
  out.Set("applyFont", Napi::Boolean::New(env, xf.apply_font != 0));
  out.Set("applyFill", Napi::Boolean::New(env, xf.apply_fill != 0));
  out.Set("applyBorder", Napi::Boolean::New(env, xf.apply_border != 0));
  out.Set("applyAlignment", Napi::Boolean::New(env, xf.apply_alignment != 0));
  out.Set("applyProtection", Napi::Boolean::New(env, xf.apply_protection != 0));
  out.Set("quotePrefix", Napi::Boolean::New(env, xf.quote_prefix != 0));
  // Without a `<protection>` child both read `null`: the model default
  // (locked, not hidden) applies.
  out.Set("locked", JsBoolOrNull(env, xf.has_protection, xf.locked));
  out.Set("hidden", JsBoolOrNull(env, xf.has_protection, xf.hidden));
  return out;
}

void SetCellXfAlignment(Napi::Env env, const fm_cell_xf_t& xf, Napi::Object* out) {
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

// Explicit `has*` flag wins; otherwise the flag is implied by the fields it covers.
inline int Implied(bool has_present, bool has_value, bool implied) {
  return has_present ? (has_value ? 1 : 0) : (implied ? 1 : 0);
}

void PullCellXf(CheckedSpecReader& reader, const Napi::Object& record, fm_cell_xf_t* out) {
  fm_cell_xf_t& xf = *out;
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
  xf.has_horizontal_align = Implied(has_horizontal_align_present, has_horizontal_align_value, horizontal_align_present);
  xf.has_vertical_align = Implied(has_vertical_align_present, has_vertical_align_value, vertical_align_present);
  xf.has_wrap_text = Implied(has_wrap_text_present, has_wrap_text_value, wrap_text_present);
  xf.has_justify_last_line =
      Implied(has_justify_last_line_present, has_justify_last_line_value, justify_last_line_present);
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
  xf.reading_order = reader.I32(record, "readingOrder", 0, &reading_order_present);
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
  xf.has_alignment = Implied(has_alignment_present, has_alignment_value, has_supplied_alignment);
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
  xf.has_protection = Implied(has_protection_present, has_protection_value, locked_present || hidden_present);
  xf.locked = locked_value ? 1 : 0;
  xf.hidden = hidden_value ? 1 : 0;
}

}  // namespace

namespace detail {

/// Builds the `getCellXf` / `getCellStyleXf` record.
Napi::Object CellXfResultToJs(Napi::Env env, const fm_cell_xf_t& xf) {
  Napi::Object out = CellXfToJs(env, xf);
  out.Set("xfId", Napi::Number::New(env, xf.xf_id));
  SetCellXfAlignment(env, xf, &out);
  return out;
}

// Shared body of the style-pool count getters that take no argument.
using StyleCountFn = fm_status_t (*)(const fm_workbook_t*, std::size_t*);

Napi::Value StyleCountResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle, StyleCountFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t n = 0;
  const fm_status_t rc = fn(handle, &n);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(n)));
}

}  // namespace detail

// ---- Cell-XF index --------------------------------------------------

Napi::Value Workbook::GetCellXfIndex(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* xf) {
    const uint32_t sheet = ArgU32(info, 0);
    const uint32_t row = ArgU32(info, 1);
    const uint32_t col = ArgU32(info, 2);
    return fm_sheet_get_cell_xf_index(handle, sheet, row, col, xf);
  });
}

Napi::Value Workbook::SetCellXfIndex(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_cell_xf_index);
}

Napi::Value Workbook::ClearCellXfIndex(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_clear_cell_xf_index);
}

Napi::Value Workbook::SetRangeXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRect(reader, ArgObject(info, 1), &rect)) {
    return env.Undefined();
  }
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_sheet_set_range_xf_index(handle, ArgU32(info, 0), &rect, ArgU32(info, 2));
  });
}

// ---- Style getters --------------------------------------------------

Napi::Value Workbook::GetCellXf(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_cell_xf_t>(
      info,
      [&](fm_workbook_t* handle, fm_cell_xf_t* xf) { return fm_workbook_get_cell_xf(handle, ArgU32(info, 0), xf); },
      CellXfResultToJs);
}

Napi::Value Workbook::GetFont(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_font_record_t>(
      info,
      [&](fm_workbook_t* handle, fm_font_record_t* f) { return fm_workbook_get_font(handle, ArgU32(info, 0), f); },
      FontRecordToJs);
}

Napi::Value Workbook::GetFill(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_fill_record_t>(
      info,
      [&](fm_workbook_t* handle, fm_fill_record_t* f) { return fm_workbook_get_fill(handle, ArgU32(info, 0), f); },
      FillRecordToJs);
}

Napi::Value Workbook::GetBorder(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_border_record_t>(
      info,
      [&](fm_workbook_t* handle, fm_border_record_t* b) { return fm_workbook_get_border(handle, ArgU32(info, 0), b); },
      BorderRecordToJs);
}

Napi::Value Workbook::GetNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  const char* s = nullptr;
  const fm_status_t rc = fm_workbook_get_num_fmt_string(handle_, ArgU32(info, 0), &present, &s);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(JsString(env, s)) : env.Null());
}

Napi::Value Workbook::GetDxf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t dxf_index = ArgU32(info, 0);
  fm_dxf_record_t d{};
  fm_status_t rc = fm_workbook_get_dxf(handle_, dxf_index, &d);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
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
  std::string name;
  fm_font_record_t fr{};
  return AddStyleRecord(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullFontRecord(reader, record, &name, &fr); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_font(handle, &fr, idx); });
}

Napi::Value Workbook::SetFont(const Napi::CallbackInfo& info) {
  std::string name;
  fm_font_record_t fr{};
  return SetStyleRecord(
      info, 1,
      [&](CheckedSpecReader& reader, const Napi::Object& record) { PullFontRecord(reader, record, &name, &fr); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_font(handle, ArgU32(info, 0), &fr); });
}

Napi::Value Workbook::SetDefaultFont(const Napi::CallbackInfo& info) {
  std::string name;
  fm_font_record_t fr{};
  return SetStyleRecord(
      info, 0,
      [&](CheckedSpecReader& reader, const Napi::Object& record) { PullFontRecord(reader, record, &name, &fr); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_default_font(handle, &fr); });
}

Napi::Value Workbook::AddFill(const Napi::CallbackInfo& info) {
  fm_fill_record_t fr{};
  std::vector<fm_gradient_stop_t> stops;
  return AddStyleRecord(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { fr = PullFillRecord(reader, record, &stops); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_fill(handle, &fr, idx); });
}

Napi::Value Workbook::AddBorder(const Napi::CallbackInfo& info) {
  fm_border_record_t br{};
  return AddStyleRecord(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { br = PullBorderRecord(reader, record); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_border(handle, &br, idx); });
}

Napi::Value Workbook::AddNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string code = ArgString(info, 0);
  uint32_t id = 0;
  fm_status_t rc = fm_workbook_add_num_fmt(handle_, code.c_str(), &id);
  return ValueOrThrow(env, rc, Napi::Number::New(env, id));
}

Napi::Value Workbook::AddXf(const Napi::CallbackInfo& info) {
  fm_cell_xf_t xf{};
  return AddStyleRecord(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullCellXf(reader, record, &xf); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_cell_xf(handle, &xf, idx); });
}

Napi::Value Workbook::AddCellStyleXf(const Napi::CallbackInfo& info) {
  fm_cell_xf_t xf{};
  return AddStyleRecord(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullCellXf(reader, record, &xf); },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_cell_style_xf(handle, &xf, idx); });
}

Napi::Value Workbook::AddDxf(const Napi::CallbackInfo& info) {
  std::string font_name;
  std::string num_fmt_code;
  std::string alignment_xml;
  std::string protection_xml;
  std::vector<fm_gradient_stop_t> fill_stops;
  fm_dxf_record_t dxf{};
  return AddStyleRecord(
      info,
      [&](CheckedSpecReader& reader, const Napi::Object& record) {
        Napi::Object font;
        if (reader.Object(record, "font", &font)) {
          dxf.font_engaged = 1;
          PullFontRecord(reader, font, &font_name, &dxf.font);
        }

        Napi::Object fill;
        if (reader.Object(record, "fill", &fill)) {
          dxf.fill_engaged = 1;
          dxf.fill = PullFillRecord(reader, fill, &fill_stops);
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
      },
      [&](fm_workbook_t* handle, std::size_t* idx) { return fm_workbook_add_dxf(handle, &dxf, idx); });
}

// ---- Style pool counts ----------------------------------------------
//
// `FontCount` / `FillCount` / `BorderCount` / `XfCount` are now emitted
// by the binding codegen (see `src/node_addon/generated/styles_counts.cc`).

Napi::Value Workbook::DxfCount(const Napi::CallbackInfo& info) {
  return StyleCountResult(info, handle_, &fm_workbook_get_dxf_count);
}

}  // namespace formulon_node
