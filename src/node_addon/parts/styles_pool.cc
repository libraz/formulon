// Style-pool bindings: font / fill / border / numFmt / xf / dxf records and
// the per-cell xf-index accessors.

#include <cstddef>
#include <cstdint>
#include <cstdlib>
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
Napi::Object FontRecordToJs(Napi::Env env, const fm_font_t& f) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("name", JsString(env, f.name));
  out.Set("sizePt", JsNumberOrNull(env, f.has_size_pt, f.size_pt));
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
void PullFontRecord(CheckedSpecReader& reader, const Napi::Object& record, std::string* name_storage, fm_font_t* out) {
  bool present = false;
  const auto flag = [&](const char* key, int32_t* has, int32_t* value) {
    *value = reader.Bool(record, key, false, &present) ? 1 : 0;
    *has = present ? 1 : 0;
  };
  out->name = reader.String(record, "name", name_storage) ? name_storage->c_str() : nullptr;
  out->size_pt = reader.Double(record, "sizePt", 0.0, &present);
  out->has_size_pt = present ? 1 : 0;
  flag("bold", &out->has_bold, &out->bold);
  flag("italic", &out->has_italic, &out->italic);
  flag("strike", &out->has_strike, &out->strike);
  out->underline = reader.I32(record, "underline", 0);
  out->vert_align = reader.I32(record, "vertAlign", 0);
  out->family = reader.I32(record, "family", 0, &present);
  out->has_family = present ? 1 : 0;
  out->charset = reader.U8(record, "charset", 0U, &present);
  out->has_charset = present ? 1 : 0;
  out->scheme = reader.I32(record, "scheme", 0);
  flag("outline", &out->has_outline, &out->outline);
  flag("shadow", &out->has_shadow, &out->shadow);
  flag("condense", &out->has_condense, &out->condense);
  flag("extend", &out->has_extend, &out->extend);
  out->color = PullColorSpec(reader, record, "color");
}

// ---- <dxf> alignment and protection ----------------------------------
//
// The engine keeps a differential format's `<alignment>` and `<protection>` as serialized elements; the
// JS records are typed, so the attributes are read and written here.

constexpr const char* kHorizontalAlignNames[] = {"general", "left",    "center",           "right",
                                                 "fill",    "justify", "centerContinuous", "distributed"};
constexpr const char* kVerticalAlignNames[] = {"top", "center", "bottom", "justify", "distributed"};

/// The value of attribute `name` in the serialized element `xml`; false when the attribute is absent.
bool XmlAttribute(const std::string& xml, const char* name, std::string* out) {
  const std::string needle = std::string(" ") + name + "=";
  const std::size_t at = xml.find(needle);
  if (at == std::string::npos) {
    return false;
  }
  const std::size_t quote = at + needle.size();
  if (quote >= xml.size() || (xml[quote] != '"' && xml[quote] != '\'')) {
    return false;
  }
  const std::size_t end = xml.find(xml[quote], quote + 1);
  if (end == std::string::npos) {
    return false;
  }
  *out = xml.substr(quote + 1, end - quote - 1);
  return true;
}

bool XmlFlag(const std::string& xml, const char* name, bool dflt) {
  std::string value;
  if (!XmlAttribute(xml, name, &value)) {
    return dflt;
  }
  return value == "1" || value == "true";
}

double XmlNumber(const std::string& xml, const char* name, double dflt) {
  std::string value;
  if (!XmlAttribute(xml, name, &value) || value.empty()) {
    return dflt;
  }
  char* end = nullptr;
  const double number = std::strtod(value.c_str(), &end);
  return end != value.c_str() ? number : dflt;
}

template <std::size_t N>
uint32_t XmlEnumOrdinal(const std::string& xml, const char* name, const char* const (&names)[N], uint32_t dflt) {
  std::string value;
  if (!XmlAttribute(xml, name, &value)) {
    return dflt;
  }
  for (std::size_t i = 0; i < N; ++i) {
    if (value == names[i]) {
      return static_cast<uint32_t>(i);
    }
  }
  return dflt;
}

/// The `Alignment` record of a serialized `<alignment/>`; an absent attribute reads as its schema default.
Napi::Object AlignmentXmlToJs(Napi::Env env, const std::string& xml) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("horizontal", Napi::Number::New(env, XmlEnumOrdinal(xml, "horizontal", kHorizontalAlignNames, 0U)));
  out.Set("vertical", Napi::Number::New(env, XmlEnumOrdinal(xml, "vertical", kVerticalAlignNames, 2U)));
  out.Set("wrapText", Napi::Boolean::New(env, XmlFlag(xml, "wrapText", false)));
  out.Set("shrinkToFit", Napi::Boolean::New(env, XmlFlag(xml, "shrinkToFit", false)));
  out.Set("indent", Napi::Number::New(env, XmlNumber(xml, "indent", 0.0)));
  out.Set("textRotation", Napi::Number::New(env, XmlNumber(xml, "textRotation", 0.0)));
  out.Set("readingOrder", Napi::Number::New(env, XmlNumber(xml, "readingOrder", 0.0)));
  out.Set("justifyLastLine", Napi::Boolean::New(env, XmlFlag(xml, "justifyLastLine", false)));
  out.Set("relativeIndent", Napi::Number::New(env, XmlNumber(xml, "relativeIndent", 0.0)));
  return out;
}

/// The `CellProtection` record of a serialized `<protection/>`.
Napi::Object ProtectionXmlToJs(Napi::Env env, const std::string& xml) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("locked", Napi::Boolean::New(env, XmlFlag(xml, "locked", true)));
  out.Set("hidden", Napi::Boolean::New(env, XmlFlag(xml, "hidden", false)));
  return out;
}

/// Appends ` name="ordinal-name"` for the enum key `key` of `record` when it is present; an ordinal past the
/// table is a RangeError.
template <std::size_t N>
void AppendEnumAttribute(CheckedSpecReader& reader, const Napi::Object& record, const char* key,
                         const char* const (&names)[N], std::string* xml) {
  bool present = false;
  const uint32_t ordinal = reader.U32(record, key, 0U, &present);
  if (!present) {
    return;
  }
  if (ordinal >= N) {
    reader.FailInteger(key, static_cast<double>(ordinal), 0.0);
    return;
  }
  *xml += std::string(" ") + key + "=\"" + names[ordinal] + "\"";
}

void AppendFlagAttribute(CheckedSpecReader& reader, const Napi::Object& record, const char* key, std::string* xml) {
  bool present = false;
  const bool value = reader.Bool(record, key, false, &present);
  if (present) {
    *xml += std::string(" ") + key + "=\"" + (value ? "1" : "0") + "\"";
  }
}

void AppendNumberAttribute(CheckedSpecReader& reader, const Napi::Object& record, const char* key, bool is_signed,
                           std::string* xml) {
  bool present = false;
  const int64_t value = is_signed ? static_cast<int64_t>(reader.I32(record, key, 0, &present))
                                  : static_cast<int64_t>(reader.U32(record, key, 0U, &present));
  if (present) {
    *xml += std::string(" ") + key + "=\"" + std::to_string(value) + "\"";
  }
}

/// The serialized `<alignment/>` of the `Alignment` record `record`: only the attributes it carries.
std::string AlignmentToXml(CheckedSpecReader& reader, const Napi::Object& record) {
  std::string xml = "<alignment";
  AppendEnumAttribute(reader, record, "horizontal", kHorizontalAlignNames, &xml);
  AppendEnumAttribute(reader, record, "vertical", kVerticalAlignNames, &xml);
  AppendFlagAttribute(reader, record, "wrapText", &xml);
  AppendFlagAttribute(reader, record, "shrinkToFit", &xml);
  AppendNumberAttribute(reader, record, "indent", false, &xml);
  AppendNumberAttribute(reader, record, "textRotation", false, &xml);
  AppendNumberAttribute(reader, record, "readingOrder", false, &xml);
  AppendFlagAttribute(reader, record, "justifyLastLine", &xml);
  AppendNumberAttribute(reader, record, "relativeIndent", true, &xml);
  return xml + "/>";
}

/// The serialized `<protection/>` of the `CellProtection` record `record`.
std::string ProtectionToXml(CheckedSpecReader& reader, const Napi::Object& record) {
  std::string xml = "<protection";
  AppendFlagAttribute(reader, record, "locked", &xml);
  AppendFlagAttribute(reader, record, "hidden", &xml);
  return xml + "/>";
}

Napi::Object FillRecordToJs(Napi::Env env, const fm_fill_t& f) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("pattern", Napi::Number::New(env, static_cast<uint32_t>(f.pattern)));
  out.Set("foreground", ColorSpecToJs(env, f.foreground));
  out.Set("background", ColorSpecToJs(env, f.background));
  out.Set("gradient", env.Null());
  if (f.has_gradient != 0) {
    Napi::Object gradient = Napi::Object::New(env);
    gradient.Set("type", Napi::Number::New(env, f.gradient_type));
    gradient.Set("angleDeg", Napi::Number::New(env, f.gradient_angle_deg));
    gradient.Set("left", Napi::Number::New(env, f.gradient_left));
    gradient.Set("right", Napi::Number::New(env, f.gradient_right));
    gradient.Set("top", Napi::Number::New(env, f.gradient_top));
    gradient.Set("bottom", Napi::Number::New(env, f.gradient_bottom));
    Napi::Array stops = Napi::Array::New(env, f.gradient_stop_count);
    for (std::size_t i = 0; i < f.gradient_stop_count; ++i) {
      Napi::Object stop = Napi::Object::New(env);
      stop.Set("position", Napi::Number::New(env, f.gradient_stops[i].position));
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
fm_fill_t PullFillRecord(CheckedSpecReader& reader, const Napi::Object& record,
                         std::vector<fm_gradient_stop_t>* stops) {
  fm_fill_t out{};
  out.pattern = reader.I32(record, "pattern", 0);
  out.foreground = PullColorSpec(reader, record, "foreground");
  out.background = PullColorSpec(reader, record, "background");
  Napi::Object gradient;
  if (reader.Object(record, "gradient", &gradient)) {
    out.has_gradient = 1;
    out.gradient_type = static_cast<int32_t>(reader.U32(gradient, "type", 0U));
    out.gradient_angle_deg = reader.Double(gradient, "angleDeg", 0.0);
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
  out.Set("color", ColorSpecToJs(env, s.color));
  return out;
}

Napi::Object BorderRecordToJs(Napi::Env env, const fm_border_t& b) {
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
  s.style = reader.I32(v, "style", 0);
  s.color = PullColorSpec(reader, v, "color");
  return s;
}

fm_border_t PullBorderRecord(CheckedSpecReader& reader, const Napi::Object& record) {
  fm_border_t out{};
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
  out.Set("numFmtId", Napi::Number::New(env, xf.num_fmt_id));
  out.Set("horizontal", JsNumberOrNull(env, xf.has_horizontal, static_cast<uint32_t>(xf.horizontal)));
  out.Set("vertical", JsNumberOrNull(env, xf.has_vertical, static_cast<uint32_t>(xf.vertical)));
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
  out->Set("textRotation", JsNumberOrNull(env, xf.has_text_rotation, xf.text_rotation));
  out->Set("indent", JsNumberOrNull(env, xf.has_indent, xf.indent));
  out->Set("relativeIndent", JsNumberOrNull(env, xf.has_relative_indent, xf.relative_indent));
  out->Set("shrinkToFit", JsBoolOrNull(env, xf.has_shrink_to_fit, xf.shrink_to_fit));
  out->Set("readingOrder", JsNumberOrNull(env, xf.has_reading_order, xf.reading_order));
}

void PullCellXf(CheckedSpecReader& reader, const Napi::Object& record, fm_cell_xf_t* out) {
  fm_cell_xf_t& xf = *out;
  xf.font_index = reader.U32(record, "fontIndex", 0U);
  xf.fill_index = reader.U32(record, "fillIndex", 0U);
  xf.border_index = reader.U32(record, "borderIndex", 0U);
  xf.num_fmt_id = reader.U32(record, "numFmtId", 0U);
  bool horizontal_align_present = false;
  bool vertical_align_present = false;
  bool wrap_text_present = false;
  bool justify_last_line_present = false;
  xf.horizontal = reader.I32(record, "horizontal", 0, &horizontal_align_present);
  xf.vertical = reader.I32(record, "vertical", 2, &vertical_align_present);
  xf.wrap_text = reader.Bool(record, "wrapText", false, &wrap_text_present) ? 1 : 0;
  xf.justify_last_line = reader.Bool(record, "justifyLastLine", false, &justify_last_line_present) ? 1 : 0;
  xf.style_xf_index = reader.U32(record, "styleXfIndex", 0U);
  xf.has_horizontal = horizontal_align_present ? 1 : 0;
  xf.has_vertical = vertical_align_present ? 1 : 0;
  xf.has_wrap_text = wrap_text_present ? 1 : 0;
  xf.has_justify_last_line = justify_last_line_present ? 1 : 0;
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
                                      relative_indent_present || shrink_to_fit_present || reading_order_present;
  xf.has_alignment = has_supplied_alignment ? 1 : 0;
  xf.apply_number_format = reader.Bool(record, "applyNumberFormat", false) ? 1 : 0;
  xf.apply_font = reader.Bool(record, "applyFont", false) ? 1 : 0;
  xf.apply_fill = reader.Bool(record, "applyFill", false) ? 1 : 0;
  xf.apply_border = reader.Bool(record, "applyBorder", false) ? 1 : 0;
  xf.apply_alignment = reader.Bool(record, "applyAlignment", false) ? 1 : 0;
  xf.apply_protection = reader.Bool(record, "applyProtection", false) ? 1 : 0;
  xf.quote_prefix = reader.Bool(record, "quotePrefix", false) ? 1 : 0;
  bool locked_present = false;
  bool hidden_present = false;
  const bool locked_value = reader.Bool(record, "locked", true, &locked_present);
  const bool hidden_value = reader.Bool(record, "hidden", false, &hidden_present);
  xf.has_protection = (locked_present || hidden_present) ? 1 : 0;
  xf.locked = locked_value ? 1 : 0;
  xf.hidden = hidden_value ? 1 : 0;
}

/// A differential-format record and the strings it borrows.
struct DxfStorage {
  std::string font_name;
  std::string number_format;
  std::string alignment_xml;
  std::string protection_xml;
  std::vector<fm_gradient_stop_t> fill_stops;
  fm_dxf_t dxf{};
};

void PullDxfRecord(CheckedSpecReader& reader, const Napi::Object& record, DxfStorage* storage) {
  fm_dxf_t& dxf = storage->dxf;
  Napi::Object font;
  if (reader.Object(record, "font", &font)) {
    dxf.has_font = 1;
    PullFontRecord(reader, font, &storage->font_name, &dxf.font);
  }

  Napi::Object fill;
  if (reader.Object(record, "fill", &fill)) {
    dxf.has_fill = 1;
    dxf.fill = PullFillRecord(reader, fill, &storage->fill_stops);
  }

  Napi::Object border;
  if (reader.Object(record, "border", &border)) {
    dxf.has_border = 1;
    dxf.border = PullBorderRecord(reader, border);
  }

  // An omitted or `null` number format is not set.
  if (reader.String(record, "numberFormat", &storage->number_format)) {
    dxf.number_format = storage->number_format.c_str();
  }

  Napi::Object alignment;
  if (reader.Object(record, "alignment", &alignment)) {
    storage->alignment_xml = AlignmentToXml(reader, alignment);
  }
  Napi::Object protection;
  if (reader.Object(record, "protection", &protection)) {
    storage->protection_xml = ProtectionToXml(reader, protection);
  }
  dxf.alignment_xml = storage->alignment_xml.c_str();
  dxf.protection_xml = storage->protection_xml.c_str();
}

}  // namespace

namespace detail {

/// Builds the `getCellXf` / `getCellStyleXf` record.
Napi::Object CellXfResultToJs(Napi::Env env, const fm_cell_xf_t& xf) {
  Napi::Object out = CellXfToJs(env, xf);
  out.Set("styleXfIndex", Napi::Number::New(env, xf.style_xf_index));
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
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  uint32_t xf = 0;
  const fm_status_t rc =
      fm_sheet_get_cell_xf_index(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &present, &xf);
  return OptionalOrNull(env, rc, present, Napi::Number::New(env, xf));
}

Napi::Value Workbook::SetCellXfIndex(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_cell_xf_index);
}

Napi::Value Workbook::ClearCellXfIndex(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_reset_cell_xf_index);
}

Napi::Value Workbook::SetRangeXfIndex(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  CheckedSpecReader reader(env);
  fm_cell_rect_t rect{};
  if (!ReadCellRectArg(reader, info, 1, &rect)) {
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
      [&](fm_workbook_t* handle, fm_cell_xf_t* xf) { return fm_workbook_get_cell_xf_at(handle, ArgU32(info, 0), xf); },
      CellXfResultToJs);
}

Napi::Value Workbook::GetFont(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_font_t>(
      info, [&](fm_workbook_t* handle, fm_font_t* f) { return fm_workbook_get_font_at(handle, ArgU32(info, 0), f); },
      FontRecordToJs);
}

// The workbook default font is font 0, the one `setDefaultFont` writes.
Napi::Value Workbook::GetDefaultFont(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_font_t>(
      info, [&](fm_workbook_t* handle, fm_font_t* f) { return fm_workbook_get_font_at(handle, 0U, f); },
      FontRecordToJs);
}

Napi::Value Workbook::GetFill(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_fill_t>(
      info, [&](fm_workbook_t* handle, fm_fill_t* f) { return fm_workbook_get_fill_at(handle, ArgU32(info, 0), f); },
      FillRecordToJs);
}

Napi::Value Workbook::GetBorder(const Napi::CallbackInfo& info) {
  return RecordGetter<fm_border_t>(
      info,
      [&](fm_workbook_t* handle, fm_border_t* b) { return fm_workbook_get_border_at(handle, ArgU32(info, 0), b); },
      BorderRecordToJs);
}

Napi::Value Workbook::GetNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  const char* s = nullptr;
  const fm_status_t rc = fm_workbook_get_number_format(handle_, ArgU32(info, 0), &present, &s);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(JsString(env, s)) : env.Null());
}

Napi::Value Workbook::ListNumberFormats(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array items = Napi::Array::New(env);
  const fm_status_t rc = CollectList<fm_number_format_info_t>(
      items, [&](std::size_t* n) { return fm_workbook_get_number_format_count(handle_, n); },
      [&](std::size_t i, fm_number_format_info_t* out) { return fm_workbook_get_number_format_at(handle_, i, out); },
      [&](const fm_number_format_info_t& entry) {
        Napi::Object out = Napi::Object::New(env);
        out.Set("numFmtId", Napi::Number::New(env, entry.num_fmt_id));
        out.Set("numberFormat", JsString(env, entry.number_format));
        return out;
      });
  return ValueOrThrow(env, rc, items);
}

Napi::Value Workbook::GetDxf(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t dxf_index = ArgU32(info, 0);
  fm_dxf_t d{};
  fm_status_t rc = fm_workbook_get_dxf_at(handle_, dxf_index, &d);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("font", d.has_font != 0 ? Napi::Value(FontRecordToJs(env, d.font)) : env.Null());
  out.Set("fill", d.has_fill != 0 ? Napi::Value(FillRecordToJs(env, d.fill)) : env.Null());
  out.Set("border", d.has_border != 0 ? Napi::Value(BorderRecordToJs(env, d.border)) : env.Null());
  out.Set("numberFormat", JsStringOrNull(env, d.number_format));
  const bool has_alignment = d.alignment_xml != nullptr && d.alignment_xml[0] != '\0';
  out.Set("alignment", has_alignment ? Napi::Value(AlignmentXmlToJs(env, d.alignment_xml)) : env.Null());
  const bool has_protection = d.protection_xml != nullptr && d.protection_xml[0] != '\0';
  out.Set("protection", has_protection ? Napi::Value(ProtectionXmlToJs(env, d.protection_xml)) : env.Null());
  return out;
}

// ---- Style adders ---------------------------------------------------

Napi::Value Workbook::AddFont(const Napi::CallbackInfo& info) {
  std::string name;
  fm_font_t fr{};
  return AddStyleRecord<uint32_t>(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullFontRecord(reader, record, &name, &fr); },
      [&](fm_workbook_t* handle, uint32_t* idx) { return fm_workbook_add_font(handle, &fr, idx); });
}

Napi::Value Workbook::SetFont(const Napi::CallbackInfo& info) {
  std::string name;
  fm_font_t fr{};
  return SetStyleRecord(
      info, 1,
      [&](CheckedSpecReader& reader, const Napi::Object& record) { PullFontRecord(reader, record, &name, &fr); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_font(handle, ArgU32(info, 0), &fr); });
}

Napi::Value Workbook::SetDefaultFont(const Napi::CallbackInfo& info) {
  std::string name;
  fm_font_t fr{};
  return SetStyleRecord(
      info, 0,
      [&](CheckedSpecReader& reader, const Napi::Object& record) { PullFontRecord(reader, record, &name, &fr); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_default_font(handle, &fr); });
}

Napi::Value Workbook::AddFill(const Napi::CallbackInfo& info) {
  fm_fill_t fr{};
  std::vector<fm_gradient_stop_t> stops;
  return AddStyleRecord<uint32_t>(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { fr = PullFillRecord(reader, record, &stops); },
      [&](fm_workbook_t* handle, uint32_t* idx) { return fm_workbook_add_fill(handle, &fr, idx); });
}

Napi::Value Workbook::SetFill(const Napi::CallbackInfo& info) {
  fm_fill_t fr{};
  std::vector<fm_gradient_stop_t> stops;
  return SetStyleRecord(
      info, 1,
      [&](CheckedSpecReader& reader, const Napi::Object& record) { fr = PullFillRecord(reader, record, &stops); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_fill(handle, ArgU32(info, 0), &fr); });
}

Napi::Value Workbook::SetBorder(const Napi::CallbackInfo& info) {
  fm_border_t br{};
  return SetStyleRecord(
      info, 1, [&](CheckedSpecReader& reader, const Napi::Object& record) { br = PullBorderRecord(reader, record); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_border(handle, ArgU32(info, 0), &br); });
}

Napi::Value Workbook::AddBorder(const Napi::CallbackInfo& info) {
  fm_border_t br{};
  return AddStyleRecord<uint32_t>(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { br = PullBorderRecord(reader, record); },
      [&](fm_workbook_t* handle, uint32_t* idx) { return fm_workbook_add_border(handle, &br, idx); });
}

Napi::Value Workbook::AddNumFmt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::string code = ArgString(info, 0);
  uint32_t id = 0;
  fm_status_t rc = fm_workbook_add_number_format(handle_, code.c_str(), &id);
  return ValueOrThrow(env, rc, Napi::Number::New(env, id));
}

Napi::Value Workbook::AddXf(const Napi::CallbackInfo& info) {
  fm_cell_xf_t xf{};
  return AddStyleRecord<uint32_t>(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullCellXf(reader, record, &xf); },
      [&](fm_workbook_t* handle, uint32_t* idx) { return fm_workbook_add_cell_xf(handle, &xf, idx); });
}

Napi::Value Workbook::SetCellXf(const Napi::CallbackInfo& info) {
  fm_cell_xf_t xf{};
  return SetStyleRecord(
      info, 1, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullCellXf(reader, record, &xf); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_cell_xf(handle, ArgU32(info, 0), &xf); });
}

Napi::Value Workbook::AddCellStyleXf(const Napi::CallbackInfo& info) {
  fm_cell_xf_t xf{};
  return AddStyleRecord<uint32_t>(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullCellXf(reader, record, &xf); },
      [&](fm_workbook_t* handle, uint32_t* idx) { return fm_workbook_add_cell_style_xf(handle, &xf, idx); });
}

Napi::Value Workbook::SetCellStyleXf(const Napi::CallbackInfo& info) {
  fm_cell_xf_t xf{};
  return SetStyleRecord(
      info, 1, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullCellXf(reader, record, &xf); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_cell_style_xf(handle, ArgU32(info, 0), &xf); });
}

Napi::Value Workbook::AddDxf(const Napi::CallbackInfo& info) {
  DxfStorage storage;
  return AddStyleRecord<uint32_t>(
      info, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullDxfRecord(reader, record, &storage); },
      [&](fm_workbook_t* handle, uint32_t* idx) { return fm_workbook_add_dxf(handle, &storage.dxf, idx); });
}

Napi::Value Workbook::SetDxf(const Napi::CallbackInfo& info) {
  DxfStorage storage;
  return SetStyleRecord(
      info, 1, [&](CheckedSpecReader& reader, const Napi::Object& record) { PullDxfRecord(reader, record, &storage); },
      [&](fm_workbook_t* handle) { return fm_workbook_set_dxf(handle, ArgU32(info, 0), &storage.dxf); });
}

// ---- Style pool counts ----------------------------------------------
//
// `FontCount` / `FillCount` / `BorderCount` / `XfCount` are now emitted
// by the binding codegen (see `src/node_addon/generated/styles_counts.cc`).

Napi::Value Workbook::DxfCount(const Napi::CallbackInfo& info) {
  return StyleCountResult(info, handle_, &fm_workbook_get_dxf_count);
}

}  // namespace formulon_node
