// Gradient fills (BrtFill fls 0x28, XFPropGradient / XFPropGradientStop) and
// the dxf font outline / shadow / condense / extend toggles survive the
// xlsx-model -> xl/styles.bin -> model and xlsb -> xlsb paths.

#include <cstdint>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/styles_reader.h"
#include "io/xlsb/dxf_props.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/styles_reader.h"
#include "io/xlsb/styles_writer.h"
#include "io/xlsb/writer.h"
#include "styles.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kStylesXml =
    R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
    R"(<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)"
    R"(<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>)"
    R"(<fills count="4">)"
    R"(<fill><patternFill patternType="none"/></fill>)"
    R"(<fill><patternFill patternType="gray125"/></fill>)"
    R"(<fill><gradientFill degree="45"><stop position="0"><color rgb="FF112233"/></stop>)"
    R"(<stop position="1"><color theme="4" tint="0.5"/></stop></gradientFill></fill>)"
    R"(<fill><gradientFill type="path" left="0.2" right="0.8" top="0.3" bottom="0.7">)"
    R"(<stop position="0"><color indexed="10"/></stop><stop position="0.5"><color rgb="FF00FF00"/></stop>)"
    R"(<stop position="1"><color rgb="FF0000FF"/></stop></gradientFill></fill>)"
    R"(</fills>)"
    R"(<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>)"
    R"(<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>)"
    R"(<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>)"
    R"(<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>)"
    R"(<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/></cellXfs>)"
    R"(<dxfs count="2">)"
    R"(<dxf><font><outline/><shadow/><condense val="0"/><extend/></font></dxf>)"
    R"(<dxf><font><i val="0"/></font><fill><gradientFill degree="90">)"
    R"(<stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color theme="1"/></stop>)"
    R"(</gradientFill></fill></dxf>)"
    R"(</dxfs></styleSheet>)";

StylesTable ParseXlsxStyles() {
  const std::string xml(kStylesXml);
  auto parsed = io::read_styles(std::vector<std::uint8_t>(xml.begin(), xml.end()));
  EXPECT_TRUE(static_cast<bool>(parsed)) << (parsed ? "" : parsed.error().message);
  return parsed ? std::move(parsed.value()) : StylesTable{};
}

void ExpectSameColor(const ColorSpec& a, const ColorSpec& b, const std::string& where) {
  EXPECT_EQ(a.kind, b.kind) << where;
  switch (a.kind) {
    case ColorSpec::Kind::kRgb:
      EXPECT_EQ(a.rgb, b.rgb) << where;
      break;
    case ColorSpec::Kind::kTheme:
      EXPECT_EQ(a.theme, b.theme) << where;
      EXPECT_NEAR(a.tint, b.tint, 1e-4) << where;
      break;
    case ColorSpec::Kind::kIndexed:
      EXPECT_EQ(a.indexed, b.indexed) << where;
      break;
    case ColorSpec::Kind::kNone:
    case ColorSpec::Kind::kAuto:
      break;
  }
}

void ExpectSameGradient(const GradientFill& a, const GradientFill& b, const std::string& where) {
  EXPECT_EQ(a.path, b.path) << where;
  EXPECT_DOUBLE_EQ(a.degree, b.degree) << where;
  EXPECT_DOUBLE_EQ(a.left, b.left) << where;
  EXPECT_DOUBLE_EQ(a.right, b.right) << where;
  EXPECT_DOUBLE_EQ(a.top, b.top) << where;
  EXPECT_DOUBLE_EQ(a.bottom, b.bottom) << where;
  ASSERT_EQ(a.stops.size(), b.stops.size()) << where;
  for (std::size_t i = 0; i < a.stops.size(); ++i) {
    EXPECT_DOUBLE_EQ(a.stops[i].position, b.stops[i].position) << where << " stop " << i;
    ExpectSameColor(a.stops[i].color, b.stops[i].color, where + " stop " + std::to_string(i));
  }
}

void ExpectSameToggles(const FontRecord& a, const FontRecord& b) {
  EXPECT_EQ(a.has_outline, b.has_outline);
  EXPECT_EQ(a.outline, b.outline);
  EXPECT_EQ(a.has_shadow, b.has_shadow);
  EXPECT_EQ(a.shadow, b.shadow);
  EXPECT_EQ(a.has_condense, b.has_condense);
  EXPECT_EQ(a.condense, b.condense);
  EXPECT_EQ(a.has_extend, b.has_extend);
  EXPECT_EQ(a.extend, b.extend);
  EXPECT_EQ(a.has_italic, b.has_italic);
  EXPECT_EQ(a.italic, b.italic);
}

void ExpectFidelity(const StylesTable& source, const StylesTable& back) {
  ASSERT_EQ(back.fills.size(), source.fills.size());
  for (std::size_t i = 0; i < source.fills.size(); ++i) {
    EXPECT_EQ(back.fills[i].has_gradient, source.fills[i].has_gradient) << "fill " << i;
    if (source.fills[i].has_gradient) {
      ExpectSameGradient(source.fills[i].gradient, back.fills[i].gradient, "fill " + std::to_string(i));
    }
  }
  ASSERT_EQ(back.dxfs.size(), source.dxfs.size());
  for (std::size_t i = 0; i < source.dxfs.size(); ++i) {
    EXPECT_EQ(back.dxfs[i].has_font, source.dxfs[i].has_font) << "dxf " << i;
    ExpectSameToggles(source.dxfs[i].font, back.dxfs[i].font);
    EXPECT_EQ(back.dxfs[i].has_fill, source.dxfs[i].has_fill) << "dxf " << i;
    EXPECT_EQ(back.dxfs[i].fill.has_gradient, source.dxfs[i].fill.has_gradient) << "dxf " << i;
    ExpectSameGradient(source.dxfs[i].fill.gradient, back.dxfs[i].fill.gradient, "dxf " + std::to_string(i));
  }
}

TEST(XlsbStylesFidelity, XlsxModelSurvivesStylesBin) {
  const StylesTable source = ParseXlsxStyles();
  ASSERT_EQ(source.fills.size(), 4U);
  ASSERT_TRUE(source.fills[3].gradient.path);
  ASSERT_TRUE(source.dxfs[0].font.has_condense);
  const std::vector<std::uint8_t> bin = write_styles_bin(source);
  auto back = read_styles_bin(SpanOf(bin));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  ExpectFidelity(source, back.value());
  // An explicit `val="0"` stays stated rather than collapsing to absent.
  EXPECT_TRUE(back.value().dxfs[0].font.has_condense);
  EXPECT_FALSE(back.value().dxfs[0].font.condense);
  EXPECT_TRUE(back.value().dxfs[1].font.has_italic);
  EXPECT_FALSE(back.value().dxfs[1].font.italic);
}

TEST(XlsbStylesFidelity, StylesBinRewritesByteForByte) {
  const std::vector<std::uint8_t> first = write_styles_bin(ParseXlsxStyles());
  auto reread = read_styles_bin(SpanOf(first));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  EXPECT_EQ(write_styles_bin(reread.value()), first);
}

/// Payloads of every record of `type` in a styles part.
std::vector<std::vector<std::uint8_t>> Payloads(const std::vector<std::uint8_t>& bin, std::uint16_t type) {
  std::vector<std::vector<std::uint8_t>> out;
  ByteSpan cursor = SpanOf(bin);
  while (cursor.size > 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      break;
    }
    if (rec.value().type == type) {
      out.emplace_back(rec.value().payload.data, rec.value().payload.data + rec.value().payload.size);
    }
  }
  return out;
}

TEST(XlsbStylesFidelity, GradientBrtFillLayout) {
  const auto fills = Payloads(write_styles_bin(ParseXlsxStyles()), static_cast<std::uint16_t>(XlsbRecordType::BrtFill));
  ASSERT_EQ(fills.size(), 4U);
  const std::vector<std::uint8_t>& path = fills[3];
  // fls, 2 x BrtColor, iGradientType, 5 x Xnum, cNumStop, 3 x (BrtColor, Xnum).
  ASSERT_EQ(path.size(), 4U + 16U + 4U + 40U + 4U + 3U * 16U);
  std::uint32_t fls = 0;
  std::memcpy(&fls, path.data(), 4U);
  EXPECT_EQ(fls, kBrtFillGradient);
  std::uint32_t type = 0;
  std::memcpy(&type, path.data() + 20U, 4U);
  EXPECT_EQ(type, 1U);
  double left = 0.0;
  std::memcpy(&left, path.data() + 32U, 8U);
  EXPECT_DOUBLE_EQ(left, 0.2);
  std::uint32_t stops = 0;
  std::memcpy(&stops, path.data() + 64U, 4U);
  EXPECT_EQ(stops, 3U);
  double second_position = 0.0;
  std::memcpy(&second_position, path.data() + 68U + 16U + 8U, 8U);
  EXPECT_DOUBLE_EQ(second_position, 0.5);
}

/// The property types of a BrtDXF payload, in order.
std::vector<std::uint16_t> PropTypes(const std::vector<std::uint8_t>& dxf) {
  std::vector<std::uint16_t> types;
  for (std::size_t at = 6U; at + 4U <= dxf.size();) {
    types.push_back(static_cast<std::uint16_t>(dxf[at] | (dxf[at + 1U] << 8U)));
    const std::size_t size = static_cast<std::size_t>(dxf[at + 2U] | (dxf[at + 3U] << 8U));
    at += size == 0U ? dxf.size() : size;
  }
  return types;
}

TEST(XlsbStylesFidelity, DxfPropertiesUseSpecTypes) {
  const auto dxfs = Payloads(write_styles_bin(ParseXlsxStyles()), kBrtDxf);
  ASSERT_EQ(dxfs.size(), 2U);
  EXPECT_EQ(PropTypes(dxfs[0]),
            (std::vector<std::uint16_t>{kDxfPropOutline, kDxfPropShadow, kDxfPropCondense, kDxfPropExtend}));
  // A gradient replaces the pattern property (XFProps forbids both).
  EXPECT_EQ(PropTypes(dxfs[1]), (std::vector<std::uint16_t>{kDxfPropFillGradient, kDxfPropGradientStop,
                                                            kDxfPropGradientStop, kDxfPropItalic}));
}

TEST(XlsbStylesFidelity, WorkbookSaveKeepsGradientsAndToggles) {
  const StylesTable source = ParseXlsxStyles();
  Workbook wb = Workbook::create_empty();
  wb.sheet(wb.add_sheet("S")).set_cell_value(0U, 0U, Value::number(1.0));
  wb.set_styles(StylesTable(source));
  auto first = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
  auto loaded = read_xlsb(SpanOf(first.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ExpectFidelity(source, loaded.value().workbook.styles());

  // xlsb -> xlsb: the reloaded book saves and reloads with the same styles.
  auto second = write_xlsb(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(second)) << second.error().message;
  auto reloaded = read_xlsb(SpanOf(second.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  ExpectFidelity(source, reloaded.value().workbook.styles());
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
