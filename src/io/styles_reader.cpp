//
// Implementation of the styles reader. See styles_reader.h for the
// public contract.
//
// Built-in number-format ids (0..163) are owned by the writer's TU
// (`styles_writer.cpp`) and exposed via `builtin_num_fmt(id)`; this
// reader resolves builtins through that helper rather than carrying a
// duplicate table.

#include "io/styles_reader.h"

#include <charconv>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <string_view>
#include <type_traits>
#include <utility>

#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace {

/// Extracts an explicit OOXML `rgb` value from a style colour element.
///
/// Recognises `rgb="AARRGGBB"` (8-hex) and `rgb="RRGGBB"` (6-hex; alpha
/// defaults to opaque `0xFF`). Falls back to `fallback` when the element is
/// absent or has no valid `rgb` attribute. Theme, indexed, and auto selectors
/// are deliberately not resolved here; `ColorSpec` carries their
/// authoritative selector and this value is only a literal/compatibility
/// fallback for callers that cannot resolve selectors.
///
/// The hex-decoding loop is shared with `cf_reader.cpp` via
/// `parse_rgb_hex()` in `xml_utils.h`.
std::uint32_t ParseColorArgb(const pugi::xml_node& color, std::uint32_t fallback) {
  if (!color) {
    return fallback;
  }
  return parse_rgb_hex(color.attribute("rgb").value(), fallback);
}

/// Parses the original specification of a `<color>` element (rgb / theme /
/// indexed / auto). Returns `kNone` when the element is absent or carries
/// none of the recognised attributes. The caller keeps the literal RGB or
/// compatibility fallback separately via `ParseColorArgb`; no theme/palette
/// rendering is performed.
ColorSpec ParseColorSpec(const pugi::xml_node& color) {
  ColorSpec spec;
  if (!color) {
    return spec;
  }
  if (pugi::xml_attribute rgb = color.attribute("rgb"); rgb) {
    spec.kind = ColorSpec::Kind::kRgb;
    spec.rgb = parse_rgb_hex(rgb.value(), 0xFF000000U);
    return spec;
  }
  if (pugi::xml_attribute theme = color.attribute("theme"); theme) {
    spec.kind = ColorSpec::Kind::kTheme;
    spec.theme = theme.as_uint(0U);
    // Signed by design (ECMA-376 bounds it to [-1.0, 1.0]), so this takes
    // the plain double lexer rather than the non-negative one; what it
    // must not admit is a NaN or an infinity, which would propagate into
    // every channel of the resolved colour.
    spec.tint = attr_f64(color, "tint", 0.0);
    return spec;
  }
  if (pugi::xml_attribute indexed = color.attribute("indexed"); indexed) {
    spec.kind = ColorSpec::Kind::kIndexed;
    spec.indexed = indexed.as_uint(0U);
    return spec;
  }
  if (color.attribute("auto")) {
    spec.kind = ColorSpec::Kind::kAuto;
  }
  return spec;
}

/// Parses a `<vertAlign val="..."/>` run into the `FontRecord::vert_align`
/// ordinal. Missing / baseline / unknown collapse to 0 (baseline).
std::uint8_t ParseVertAlign(std::string_view s) {
  if (s == "superscript") {
    return 1;
  }
  if (s == "subscript") {
    return 2;
  }
  return 0;
}

/// Maps OOXML border-style strings to the integer ordinal stored in
/// `BorderSide::style`. Unknown strings collapse to `0` (none).
std::uint8_t ParseBorderStyle(std::string_view s) {
  if (s == "none" || s.empty()) {
    return 0;
  }
  if (s == "thin") {
    return 1;
  }
  if (s == "medium") {
    return 2;
  }
  if (s == "dashed") {
    return 3;
  }
  if (s == "dotted") {
    return 4;
  }
  if (s == "thick") {
    return 5;
  }
  if (s == "double") {
    return 6;
  }
  if (s == "hair") {
    return 7;
  }
  if (s == "mediumDashed") {
    return 8;
  }
  if (s == "dashDot") {
    return 9;
  }
  if (s == "mediumDashDot") {
    return 10;
  }
  if (s == "dashDotDot") {
    return 11;
  }
  if (s == "mediumDashDotDot") {
    return 12;
  }
  if (s == "slantDashDot") {
    return 13;
  }
  return 0;
}

std::uint8_t ParseUnderline(std::string_view s) {
  if (s == "single" || s.empty()) {
    return s.empty() ? 0 : 1;
  }
  if (s == "double") {
    return 2;
  }
  if (s == "singleAccounting") {
    return 3;
  }
  if (s == "doubleAccounting") {
    return 4;
  }
  return 0;
}

std::uint8_t ParseFillPattern(std::string_view s) {
  if (s == "none" || s.empty()) {
    return 0;
  }
  if (s == "solid") {
    return 1;
  }
  if (s == "mediumGray") {
    return 2;
  }
  if (s == "darkGray") {
    return 3;
  }
  if (s == "lightGray") {
    return 4;
  }
  if (s == "darkHorizontal") {
    return 5;
  }
  if (s == "darkVertical") {
    return 6;
  }
  if (s == "darkDown") {
    return 7;
  }
  if (s == "darkUp") {
    return 8;
  }
  if (s == "darkGrid") {
    return 9;
  }
  if (s == "darkTrellis") {
    return 10;
  }
  if (s == "lightHorizontal") {
    return 11;
  }
  if (s == "lightVertical") {
    return 12;
  }
  if (s == "lightDown") {
    return 13;
  }
  if (s == "lightUp") {
    return 14;
  }
  if (s == "lightGrid") {
    return 15;
  }
  if (s == "lightTrellis") {
    return 16;
  }
  if (s == "gray125") {
    return 17;
  }
  if (s == "gray0625") {
    return 18;
  }
  return 0;
}

FontRecord ParseFontNode(const pugi::xml_node& f) {
  FontRecord rec;
  pugi::xml_node name = f.child("name");
  if (name) {
    rec.name = name.attribute("val").value();
  }
  pugi::xml_node sz = f.child("sz");
  if (sz) {
    // A font cannot be smaller than nothing, and a non-finite size feeds
    // the row-height estimate, so an unusable value keeps Excel's default.
    double size = 0.0;
    if (parse_xsd_nonneg_double(attr_str(sz, "val"), &size)) {
      rec.size = size;
    }
  }
  if (const pugi::xml_node bold = f.child("b")) {
    rec.has_bold = true;
    rec.bold = read_xsd_bool(bold, "val", true);
  }
  if (const pugi::xml_node italic = f.child("i")) {
    rec.has_italic = true;
    rec.italic = read_xsd_bool(italic, "val", true);
  }
  if (const pugi::xml_node strike = f.child("strike")) {
    rec.has_strike = true;
    rec.strike = read_xsd_bool(strike, "val", true);
  }
  pugi::xml_node u = f.child("u");
  if (u) {
    // `<u/>` defaults to "single" per OOXML.
    const std::string_view val = u.attribute("val").value();
    rec.underline = val.empty() ? 1U : ParseUnderline(val);
  }
  if (pugi::xml_node va = f.child("vertAlign")) {
    rec.vert_align = ParseVertAlign(va.attribute("val").value());
  }
  if (pugi::xml_node family = f.child("family")) {
    rec.has_family = true;
    rec.family = static_cast<std::uint8_t>(family.attribute("val").as_uint(0U));
  }
  if (pugi::xml_node charset = f.child("charset")) {
    rec.has_charset = true;
    rec.charset = static_cast<std::uint8_t>(charset.attribute("val").as_uint(0U));
  }
  rec.color_argb = ParseColorArgb(f.child("color"), 0xFF000000U);
  rec.color = ParseColorSpec(f.child("color"));
  return rec;
}

void ReadFonts(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node fonts = root.child("fonts");
  if (!fonts) {
    table.fonts.emplace_back();
    return;
  }
  for (pugi::xml_node f = fonts.child("font"); f; f = f.next_sibling("font")) {
    table.fonts.push_back(ParseFontNode(f));
  }
  if (table.fonts.empty()) {
    table.fonts.emplace_back();
  }
}

FillRecord ParseFillNode(const pugi::xml_node& fill) {
  FillRecord rec;
  pugi::xml_node pattern = fill.child("patternFill");
  if (pattern) {
    rec.pattern = ParseFillPattern(pattern.attribute("patternType").value());
    rec.fg_argb = ParseColorArgb(pattern.child("fgColor"), 0U);
    rec.bg_argb = ParseColorArgb(pattern.child("bgColor"), 0U);
    rec.fg = ParseColorSpec(pattern.child("fgColor"));
    rec.bg = ParseColorSpec(pattern.child("bgColor"));
  }
  return rec;
}

void ReadFills(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node fills = root.child("fills");
  if (!fills) {
    table.fills.emplace_back();
    return;
  }
  for (pugi::xml_node fill = fills.child("fill"); fill; fill = fill.next_sibling("fill")) {
    table.fills.push_back(ParseFillNode(fill));
  }
  if (table.fills.empty()) {
    table.fills.emplace_back();
  }
}

void ReadBorderSide(const pugi::xml_node& side, BorderSide* out) {
  if (!side) {
    return;
  }
  out->style = ParseBorderStyle(side.attribute("style").value());
  out->color_argb = ParseColorArgb(side.child("color"), 0U);
  out->color = ParseColorSpec(side.child("color"));
}

void ReadBorders(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node borders = root.child("borders");
  if (!borders) {
    table.borders.emplace_back();
    return;
  }
  for (pugi::xml_node b = borders.child("border"); b; b = b.next_sibling("border")) {
    BorderRecord rec;
    rec.diagonal_up = b.attribute("diagonalUp").as_bool(false);
    rec.diagonal_down = b.attribute("diagonalDown").as_bool(false);
    ReadBorderSide(b.child("left"), &rec.left);
    ReadBorderSide(b.child("right"), &rec.right);
    ReadBorderSide(b.child("top"), &rec.top);
    ReadBorderSide(b.child("bottom"), &rec.bottom);
    ReadBorderSide(b.child("diagonal"), &rec.diagonal);
    table.borders.push_back(rec);
  }
  if (table.borders.empty()) {
    table.borders.emplace_back();
  }
}

BorderRecord ParseBorderNode(const pugi::xml_node& b) {
  BorderRecord rec;
  rec.diagonal_up = b.attribute("diagonalUp").as_bool(false);
  rec.diagonal_down = b.attribute("diagonalDown").as_bool(false);
  ReadBorderSide(b.child("left"), &rec.left);
  ReadBorderSide(b.child("right"), &rec.right);
  ReadBorderSide(b.child("top"), &rec.top);
  ReadBorderSide(b.child("bottom"), &rec.bottom);
  ReadBorderSide(b.child("diagonal"), &rec.diagonal);
  return rec;
}

void ReadNumFmts(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node num_fmts = root.child("numFmts");
  if (!num_fmts) {
    return;
  }
  for (pugi::xml_node n = num_fmts.child("numFmt"); n; n = n.next_sibling("numFmt")) {
    NumFmtRecord rec;
    rec.id = static_cast<std::uint16_t>(n.attribute("numFmtId").as_uint(0U));
    rec.format_string_index = static_cast<std::uint32_t>(table.num_fmt_strings.size());
    table.num_fmt_strings.emplace_back(n.attribute("formatCode").value());
    table.num_fmts.push_back(rec);
  }
}

std::string XfContext(std::string_view section, std::size_t index, std::string_view attr, std::string_view value) {
  std::string context = "context=styles_reader part=xl/styles.xml section=";
  context.append(section);
  context.append(" index=");
  context.append(std::to_string(index));
  context.append(" attribute=");
  context.append(attr);
  context.append(" value=");
  context.append(value);
  return context;
}

Error InvalidXfAttribute(std::string_view section, std::size_t index, std::string_view attr, std::string_view value) {
  return make_error(FormulonErrorCode::kIoSheetCorrupt, "styles.xml: invalid style xf attribute",
                    XfContext(section, index, attr, value));
}

std::string CollapseXmlWhitespace(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  bool pending_space = false;
  for (const char ch : text) {
    const bool is_space = ch == ' ' || ch == '\t' || ch == '\r' || ch == '\n';
    if (is_space) {
      if (!out.empty()) {
        pending_space = true;
      }
      continue;
    }
    if (pending_space) {
      out.push_back(' ');
      pending_space = false;
    }
    out.push_back(ch);
  }
  return out;
}

template <typename T>
Expected<T, Error> ParseIntegerAttribute(const pugi::xml_attribute& attr, T default_value, std::string_view section,
                                         std::size_t index) {
  if (!attr) {
    return default_value;
  }
  const std::string collapsed = CollapseXmlWhitespace(attr.value());
  std::string_view text(collapsed);
  if (text.empty()) {
    return InvalidXfAttribute(section, index, attr.name(), text);
  }
  if constexpr (std::is_unsigned_v<T>) {
    if (text.front() == '-') {
      return InvalidXfAttribute(section, index, attr.name(), text);
    }
    if (text.front() == '+') {
      text.remove_prefix(1U);
      if (text.empty()) {
        return InvalidXfAttribute(section, index, attr.name(), text);
      }
    }
  } else if (!text.empty() && text.front() == '+') {
    text.remove_prefix(1U);
    if (text.empty()) {
      return InvalidXfAttribute(section, index, attr.name(), text);
    }
  }
  T value{};
  const auto parsed = std::from_chars(text.data(), text.data() + text.size(), value);
  if (parsed.ec != std::errc() || parsed.ptr != text.data() + text.size()) {
    return InvalidXfAttribute(section, index, attr.name(), text);
  }
  return value;
}

Expected<std::uint32_t, Error> ParseU32Attribute(const pugi::xml_attribute& attr, std::uint32_t default_value,
                                                 std::string_view section, std::size_t index) {
  return ParseIntegerAttribute<std::uint32_t>(attr, default_value, section, index);
}

Expected<std::int32_t, Error> ParseI32Attribute(const pugi::xml_attribute& attr, std::int32_t default_value,
                                                std::string_view section, std::size_t index) {
  return ParseIntegerAttribute<std::int32_t>(attr, default_value, section, index);
}

Expected<std::uint16_t, Error> ParseU16Attribute(const pugi::xml_attribute& attr, std::uint16_t default_value,
                                                 std::string_view section, std::size_t index) {
  auto parsed = ParseU32Attribute(attr, default_value, section, index);
  if (!parsed) {
    return parsed.error();
  }
  if (parsed.value() > std::numeric_limits<std::uint16_t>::max()) {
    return InvalidXfAttribute(section, index, attr.name(), attr.value());
  }
  return static_cast<std::uint16_t>(parsed.value());
}

Expected<bool, Error> ParseBoolAttribute(const pugi::xml_attribute& attr, bool default_value, std::string_view section,
                                         std::size_t index) {
  if (!attr) {
    return default_value;
  }
  const std::string text = CollapseXmlWhitespace(attr.value());
  if (text == "0" || text == "false") {
    return false;
  }
  if (text == "1" || text == "true") {
    return true;
  }
  return InvalidXfAttribute(section, index, attr.name(), text);
}

Expected<std::uint8_t, Error> ParseHorizontalAlignStrict(const pugi::xml_attribute& attr, std::string_view section,
                                                         std::size_t index) {
  if (!attr) {
    return static_cast<std::uint8_t>(0);
  }
  // ST_HorizontalAlignment is an xsd:string-derived lexical type.  Unlike
  // numeric and boolean attributes, its whitespace is significant: XML
  // values such as ` center ` are not the `center` token.
  const std::string_view value(attr.value());
  if (value == "general") {
    return static_cast<std::uint8_t>(0);
  }
  if (value == "left") {
    return static_cast<std::uint8_t>(1);
  }
  if (value == "center") {
    return static_cast<std::uint8_t>(2);
  }
  if (value == "right") {
    return static_cast<std::uint8_t>(3);
  }
  if (value == "fill") {
    return static_cast<std::uint8_t>(4);
  }
  if (value == "justify") {
    return static_cast<std::uint8_t>(5);
  }
  if (value == "centerContinuous") {
    return static_cast<std::uint8_t>(6);
  }
  if (value == "distributed") {
    return static_cast<std::uint8_t>(7);
  }
  return InvalidXfAttribute(section, index, attr.name(), value);
}

Expected<std::uint8_t, Error> ParseVerticalAlignStrict(const pugi::xml_attribute& attr, std::string_view section,
                                                       std::size_t index) {
  if (!attr) {
    return static_cast<std::uint8_t>(2);
  }
  // ST_VerticalAlignment is also xsd:string-derived; preserve its lexical
  // whitespace and reject padded tokens.
  const std::string_view value(attr.value());
  if (value == "bottom") {
    return static_cast<std::uint8_t>(2);
  }
  if (value == "top") {
    return static_cast<std::uint8_t>(0);
  }
  if (value == "center") {
    return static_cast<std::uint8_t>(1);
  }
  if (value == "justify") {
    return static_cast<std::uint8_t>(3);
  }
  if (value == "distributed") {
    return static_cast<std::uint8_t>(4);
  }
  return InvalidXfAttribute(section, index, attr.name(), value);
}

Expected<void, Error> ParseCellXfNode(const pugi::xml_node& xf, CellXf* rec, std::string_view section,
                                      std::size_t index) {
  ASSIGN_OR_RETURN(rec->font_index, ParseU32Attribute(xf.attribute("fontId"), 0U, section, index));
  ASSIGN_OR_RETURN(rec->fill_index, ParseU32Attribute(xf.attribute("fillId"), 0U, section, index));
  ASSIGN_OR_RETURN(rec->border_index, ParseU32Attribute(xf.attribute("borderId"), 0U, section, index));
  ASSIGN_OR_RETURN(rec->num_fmt_id, ParseU16Attribute(xf.attribute("numFmtId"), 0U, section, index));
  ASSIGN_OR_RETURN(rec->xf_id, ParseU32Attribute(xf.attribute("xfId"), 0U, section, index));
  ASSIGN_OR_RETURN(rec->apply_number_format,
                   ParseBoolAttribute(xf.attribute("applyNumberFormat"), false, section, index));
  ASSIGN_OR_RETURN(rec->apply_font, ParseBoolAttribute(xf.attribute("applyFont"), false, section, index));
  ASSIGN_OR_RETURN(rec->apply_fill, ParseBoolAttribute(xf.attribute("applyFill"), false, section, index));
  ASSIGN_OR_RETURN(rec->apply_border, ParseBoolAttribute(xf.attribute("applyBorder"), false, section, index));
  ASSIGN_OR_RETURN(rec->apply_alignment, ParseBoolAttribute(xf.attribute("applyAlignment"), false, section, index));
  ASSIGN_OR_RETURN(rec->apply_protection, ParseBoolAttribute(xf.attribute("applyProtection"), false, section, index));
  ASSIGN_OR_RETURN(rec->quote_prefix, ParseBoolAttribute(xf.attribute("quotePrefix"), false, section, index));

  const pugi::xml_node align = xf.child("alignment");
  rec->has_alignment = static_cast<bool>(align);
  if (align) {
    if (const pugi::xml_attribute attr = align.attribute("horizontal")) {
      rec->has_horizontal_align = true;
      ASSIGN_OR_RETURN(rec->horizontal_align, ParseHorizontalAlignStrict(attr, section, index));
    }
    if (const pugi::xml_attribute attr = align.attribute("vertical")) {
      rec->has_vertical_align = true;
      ASSIGN_OR_RETURN(rec->vertical_align, ParseVerticalAlignStrict(attr, section, index));
    }
    if (const pugi::xml_attribute attr = align.attribute("wrapText")) {
      rec->has_wrap_text = true;
      ASSIGN_OR_RETURN(rec->wrap_text, ParseBoolAttribute(attr, false, section, index));
    }
    if (const pugi::xml_attribute attr = align.attribute("justifyLastLine")) {
      rec->has_justify_last_line = true;
      ASSIGN_OR_RETURN(rec->justify_last_line, ParseBoolAttribute(attr, false, section, index));
    }

    if (const pugi::xml_attribute attr = align.attribute("textRotation")) {
      rec->has_text_rotation = true;
      ASSIGN_OR_RETURN(rec->text_rotation, ParseU32Attribute(attr, 0U, section, index));
      if (rec->text_rotation > 180U && rec->text_rotation != 255U) {
        return InvalidXfAttribute(section, index, "textRotation", attr.value());
      }
    }
    if (const pugi::xml_attribute attr = align.attribute("indent")) {
      rec->has_indent = true;
      ASSIGN_OR_RETURN(rec->indent, ParseU32Attribute(attr, 0U, section, index));
      if (rec->indent > 255U) {
        return InvalidXfAttribute(section, index, "indent", attr.value());
      }
    }
    if (const pugi::xml_attribute attr = align.attribute("relativeIndent")) {
      rec->has_relative_indent = true;
      ASSIGN_OR_RETURN(rec->relative_indent, ParseI32Attribute(attr, 0, section, index));
    }
    if (const pugi::xml_attribute attr = align.attribute("shrinkToFit")) {
      rec->has_shrink_to_fit = true;
      ASSIGN_OR_RETURN(rec->shrink_to_fit, ParseBoolAttribute(attr, false, section, index));
    }
    if (const pugi::xml_attribute attr = align.attribute("readingOrder")) {
      rec->has_reading_order = true;
      ASSIGN_OR_RETURN(rec->reading_order, ParseU32Attribute(attr, 0U, section, index));
      if (rec->reading_order > 2U) {
        return InvalidXfAttribute(section, index, "readingOrder", attr.value());
      }
    }
  }
  if (const pugi::xml_node protection = xf.child("protection")) {
    rec->has_protection = true;
    // Schema defaults: locked=true, hidden=false. A cell is only
    // unlocked when it explicitly carries locked="0".
    ASSIGN_OR_RETURN(rec->locked, ParseBoolAttribute(protection.attribute("locked"), true, section, index));
    ASSIGN_OR_RETURN(rec->hidden, ParseBoolAttribute(protection.attribute("hidden"), false, section, index));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadCellStyleXfs(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node xfs = root.child("cellStyleXfs");
  if (!xfs) {
    return Expected<void, Error>::Ok();
  }
  std::size_t index = 0;
  for (pugi::xml_node xf = xfs.child("xf"); xf; xf = xf.next_sibling("xf")) {
    CellXf rec;
    RETURN_IF_ERROR(ParseCellXfNode(xf, &rec, "cellStyleXfs", index));
    table.cell_style_xfs.push_back(rec);
    ++index;
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadCellXfs(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node xfs = root.child("cellXfs");
  if (!xfs) {
    table.cell_xfs.emplace_back();
    return Expected<void, Error>::Ok();
  }
  std::size_t index = 0;
  for (pugi::xml_node xf = xfs.child("xf"); xf; xf = xf.next_sibling("xf")) {
    CellXf rec;
    RETURN_IF_ERROR(ParseCellXfNode(xf, &rec, "cellXfs", index));
    table.cell_xfs.push_back(rec);
    ++index;
  }
  if (table.cell_xfs.empty()) {
    CellXf def;
    def.vertical_align = 2;
    table.cell_xfs.push_back(def);
  }
  return Expected<void, Error>::Ok();
}

void ReadCellStyles(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node styles = root.child("cellStyles");
  if (!styles) {
    return;
  }
  for (pugi::xml_node cs = styles.child("cellStyle"); cs; cs = cs.next_sibling("cellStyle")) {
    CellStyleRecord rec;
    rec.name = cs.attribute("name").value();
    rec.xf_id = cs.attribute("xfId").as_uint(0U);
    if (pugi::xml_attribute builtin_attr = cs.attribute("builtinId"); builtin_attr) {
      rec.builtin_id = builtin_attr.as_uint(CellStyleRecord::kBuiltinIdNone);
    }
    rec.i_level = cs.attribute("iLevel").as_uint(0U);
    rec.hidden = cs.attribute("hidden").as_bool(false);
    rec.custom_builtin = cs.attribute("customBuiltin").as_bool(false);
    table.cell_styles.push_back(std::move(rec));
  }
}

std::string CaptureRootExtraAttrs(const pugi::xml_node& root) {
  std::string out;
  for (pugi::xml_attribute attr : root.attributes()) {
    const std::string_view name(attr.name());
    if (name == "xmlns" || (name.rfind("xmlns:", 0U) != 0U && name != "mc:Ignorable")) {
      continue;
    }
    // Verbatim-retained attributes go through the same escape rule as
    // modelled ones, so a captured value cannot be written in a weaker form
    // than the tag it is spliced into.
    append_xml_attr(out, name, attr.value());
  }
  return out;
}

void ReadDxfs(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node dxfs = root.child("dxfs");
  if (!dxfs) {
    return;
  }
  for (pugi::xml_node dxf = dxfs.child("dxf"); dxf; dxf = dxf.next_sibling("dxf")) {
    DifferentialFormat rec;
    if (pugi::xml_node font = dxf.child("font")) {
      rec.has_font = true;
      rec.font = ParseFontNode(font);
    }
    if (pugi::xml_node fill = dxf.child("fill")) {
      rec.has_fill = true;
      rec.fill = ParseFillNode(fill);
    }
    if (pugi::xml_node border = dxf.child("border")) {
      rec.has_border = true;
      rec.border = ParseBorderNode(border);
    }
    if (pugi::xml_node num_fmt = dxf.child("numFmt")) {
      rec.has_num_fmt = true;
      rec.num_fmt_id = static_cast<std::uint16_t>(num_fmt.attribute("numFmtId").as_uint(0U));
      rec.num_fmt_code = num_fmt.attribute("formatCode").value();
    }
    // `<alignment>` / `<protection>` are not modelled structurally on a dxf;
    // retain their serialized XML so they round-trip semantically. Parsing
    // may normalize lexical formatting.
    if (pugi::xml_node alignment = dxf.child("alignment")) {
      rec.alignment_xml = raw_xml(alignment);
    }
    if (pugi::xml_node protection = dxf.child("protection")) {
      rec.protection_xml = raw_xml(protection);
    }
    table.dxfs.push_back(std::move(rec));
  }
}

}  // namespace

/// Rewrites every out-of-range index in `table` to the default record 0.
///
/// A third-party writer can emit `<xf fontId="7"/>` against a three-font
/// table; Excel opens such a file and falls back to the default record
/// rather than refusing it, and so do we. The alternative — rejecting
/// the load — would trade a cosmetic loss for an unreadable workbook,
/// which is the same trade `cell_parser.cpp` already declines for a
/// malformed `<c s="...">`. Normalising here (rather than clamping in
/// each getter) is what keeps the header's promise that a stored index
/// resolves, so `fm_styles_get_*` cannot fail on a workbook that loaded,
/// and the writer cannot re-emit a dangling reference.
void NormalizeStyleIndices(StylesTable& table) {
  const auto clamp = [](std::uint32_t& index, std::size_t size) {
    if (index >= size) {
      index = 0U;
    }
  };
  const auto clamp_xf_table = [&](std::vector<CellXf>& xfs, std::size_t style_xf_count) {
    for (CellXf& xf : xfs) {
      clamp(xf.font_index, table.fonts.size());
      clamp(xf.fill_index, table.fills.size());
      clamp(xf.border_index, table.borders.size());
      clamp(xf.xf_id, style_xf_count);
    }
  };
  // `<cellStyleXfs>` entries carry an `xfId` too, but it is meaningless
  // there (the table is its own root), so it is normalised against
  // itself rather than dropped.
  clamp_xf_table(table.cell_style_xfs, table.cell_style_xfs.size());
  clamp_xf_table(table.cell_xfs, table.cell_style_xfs.size());
  for (CellStyleRecord& style : table.cell_styles) {
    clamp(style.xf_id, table.cell_style_xfs.size());
  }
}

Expected<StylesTable, Error> read_styles(const std::vector<std::uint8_t>& styles_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, styles_bytes, "styles_reader", "styles.xml"));
  pugi::xml_node root = doc.child("styleSheet");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "styles.xml: missing <styleSheet> root",
                      "context=styles_reader part=xl/styles.xml");
  }

  StylesTable table;
  ReadNumFmts(root, table);
  ReadFonts(root, table);
  ReadFills(root, table);
  ReadBorders(root, table);
  RETURN_IF_ERROR(ReadCellStyleXfs(root, table));
  RETURN_IF_ERROR(ReadCellXfs(root, table));
  ReadCellStyles(root, table);
  ReadDxfs(root, table);
  NormalizeStyleIndices(table);
  table.root_extra_attrs = CaptureRootExtraAttrs(root);
  if (pugi::xml_node colors = root.child("colors")) {
    table.colors_xml = raw_xml(colors);
  }
  if (pugi::xml_node table_styles = root.child("tableStyles")) {
    table.table_styles_xml = raw_xml(table_styles);
  }
  if (pugi::xml_node ext_lst = root.child("extLst")) {
    table.ext_lst_xml = raw_xml(ext_lst);
  }
  capture_unknown_children(root,
                           {"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs", "cellStyles", "dxfs",
                            "colors", "tableStyles", "extLst"},
                           table.unknown_top_level_xml);
  return table;
}

}  // namespace io
}  // namespace formulon
