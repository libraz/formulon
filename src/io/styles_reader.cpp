// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the styles reader. See styles_reader.h for the
// public contract.
//
// Built-in number-format ids (0..163) are owned by the writer's TU
// (`styles_writer.cpp`) and exposed via `builtin_num_fmt(id)`; this
// reader resolves builtins through that helper rather than carrying a
// duplicate table.

#include "io/styles_reader.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

/// Parses an OOXML `<color>` element attached to a style record.
///
/// Recognises `rgb="AARRGGBB"` (8-hex) and `rgb="RRGGBB"` (6-hex; alpha
/// defaults to opaque `0xFF`). Falls back to `0` when the element is
/// absent or unparseable; the caller (which already has a default
/// fallback in mind for "no colour set") does not need to distinguish
/// the two.
std::uint32_t ParseColorArgb(const pugi::xml_node& color, std::uint32_t fallback) {
  if (!color) {
    return fallback;
  }
  const std::string_view rgb = color.attribute("rgb").value();
  if (rgb.empty()) {
    return fallback;
  }
  // Accept both 8-char and 6-char hex.
  if (rgb.size() != 6 && rgb.size() != 8) {
    return fallback;
  }
  std::uint32_t out = 0;
  for (char c : rgb) {
    std::uint32_t digit = 0;
    if (c >= '0' && c <= '9') {
      digit = static_cast<std::uint32_t>(c - '0');
    } else if (c >= 'a' && c <= 'f') {
      digit = static_cast<std::uint32_t>(c - 'a' + 10);
    } else if (c >= 'A' && c <= 'F') {
      digit = static_cast<std::uint32_t>(c - 'A' + 10);
    } else {
      return fallback;
    }
    out = (out << 4U) | digit;
  }
  if (rgb.size() == 6) {
    out |= 0xFF000000U;
  }
  return out;
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

std::uint8_t ParseHorizontalAlign(std::string_view s) {
  if (s.empty() || s == "general") {
    return 0;
  }
  if (s == "left") {
    return 1;
  }
  if (s == "center") {
    return 2;
  }
  if (s == "right") {
    return 3;
  }
  if (s == "fill") {
    return 4;
  }
  if (s == "justify") {
    return 5;
  }
  if (s == "centerContinuous") {
    return 6;
  }
  if (s == "distributed") {
    return 7;
  }
  return 0;
}

std::uint8_t ParseVerticalAlign(std::string_view s) {
  if (s == "top") {
    return 0;
  }
  if (s == "center") {
    return 1;
  }
  // "bottom" is the OOXML default; treat empty as bottom too.
  if (s.empty() || s == "bottom") {
    return 2;
  }
  if (s == "justify") {
    return 3;
  }
  if (s == "distributed") {
    return 4;
  }
  return 2;
}

void ReadFonts(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node fonts = root.child("fonts");
  if (!fonts) {
    table.fonts.emplace_back();
    return;
  }
  for (pugi::xml_node f = fonts.child("font"); f; f = f.next_sibling("font")) {
    FontRecord rec;
    pugi::xml_node name = f.child("name");
    if (name) {
      rec.name = name.attribute("val").value();
    }
    pugi::xml_node sz = f.child("sz");
    if (sz) {
      rec.size = sz.attribute("val").as_double(11.0);
    }
    rec.bold = static_cast<bool>(f.child("b"));
    rec.italic = static_cast<bool>(f.child("i"));
    rec.strike = static_cast<bool>(f.child("strike"));
    pugi::xml_node u = f.child("u");
    if (u) {
      // `<u/>` defaults to "single" per OOXML.
      const std::string_view val = u.attribute("val").value();
      rec.underline = val.empty() ? 1U : ParseUnderline(val);
    }
    rec.color_argb = ParseColorArgb(f.child("color"), 0xFF000000U);
    table.fonts.push_back(std::move(rec));
  }
  if (table.fonts.empty()) {
    table.fonts.emplace_back();
  }
}

void ReadFills(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node fills = root.child("fills");
  if (!fills) {
    table.fills.emplace_back();
    return;
  }
  for (pugi::xml_node fill = fills.child("fill"); fill; fill = fill.next_sibling("fill")) {
    FillRecord rec;
    pugi::xml_node pattern = fill.child("patternFill");
    if (pattern) {
      rec.pattern = ParseFillPattern(pattern.attribute("patternType").value());
      rec.fg_argb = ParseColorArgb(pattern.child("fgColor"), 0U);
      rec.bg_argb = ParseColorArgb(pattern.child("bgColor"), 0U);
    }
    table.fills.push_back(rec);
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

void ReadCellXfs(const pugi::xml_node& root, StylesTable& table) {
  pugi::xml_node xfs = root.child("cellXfs");
  if (!xfs) {
    table.cell_xfs.emplace_back();
    return;
  }
  for (pugi::xml_node xf = xfs.child("xf"); xf; xf = xf.next_sibling("xf")) {
    CellXf rec;
    rec.font_index = xf.attribute("fontId").as_uint(0U);
    rec.fill_index = xf.attribute("fillId").as_uint(0U);
    rec.border_index = xf.attribute("borderId").as_uint(0U);
    rec.num_fmt_id = static_cast<std::uint16_t>(xf.attribute("numFmtId").as_uint(0U));
    pugi::xml_node align = xf.child("alignment");
    if (align) {
      rec.horizontal_align = ParseHorizontalAlign(align.attribute("horizontal").value());
      rec.vertical_align = ParseVerticalAlign(align.attribute("vertical").value());
      rec.wrap_text = align.attribute("wrapText").as_bool(false);
    } else {
      // Default vertical alignment is "bottom" (ordinal 2) per OOXML.
      rec.vertical_align = 2;
    }
    table.cell_xfs.push_back(rec);
  }
  if (table.cell_xfs.empty()) {
    CellXf def;
    def.vertical_align = 2;
    table.cell_xfs.push_back(def);
  }
}

}  // namespace

Expected<StylesTable, Error> read_styles(const std::vector<std::uint8_t>& styles_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(styles_bytes.data(), styles_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=styles_reader part=xl/styles.xml desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "styles.xml: pugixml parse failed", std::move(ctx));
  }
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
  ReadCellXfs(root, table);
  return table;
}

}  // namespace io
}  // namespace formulon
