#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "io/cf_reader.h"
#include "io/cf_writer.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "theme.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon::cf {
namespace {

// Resolved RGB values come from the measured theme/tint table (Office 2013-2022
// scheme): theme index 4 = accent1 etc.
constexpr Color Rgb(std::uint8_t r, std::uint8_t g, std::uint8_t b) {
  return Color{r, g, b, 255, ColorSpec{}};
}

std::vector<ConditionalFormat> ReadBlocks(const std::string& body) {
  pugi::xml_document doc;
  const std::string xml = "<worksheet>" + body + "</worksheet>";
  EXPECT_TRUE(doc.load_string(xml.c_str()));
  auto formats = io::read_conditional_formats(doc.child("worksheet"));
  EXPECT_TRUE(static_cast<bool>(formats));
  return formats.take();
}

const char* kColorScaleXml =
    "<conditionalFormatting sqref=\"A1:A3\"><cfRule type=\"colorScale\" priority=\"1\"><colorScale>"
    "<cfvo type=\"min\"/><cfvo type=\"max\"/>"
    "<color theme=\"4\" tint=\"0.39997558519241921\"/><color theme=\"5\" tint=\"-0.249977111117893\"/>"
    "</colorScale></cfRule></conditionalFormatting>";

const char* kDataBarXml =
    "<conditionalFormatting sqref=\"B1:B3\"><cfRule type=\"dataBar\" priority=\"2\"><dataBar>"
    "<cfvo type=\"min\"/><cfvo type=\"max\"/><color theme=\"8\" tint=\"0.59999389629810485\"/>"
    "</dataBar></cfRule></conditionalFormatting>";

struct Fixture {
  Workbook wb = Workbook::create();
  Arena arena;
  eval::EvalState state;

  explicit Fixture(const std::string& body) {
    Sheet& sheet = wb.sheet(0);
    sheet.set_cell_value(0, 0, Value::number(0.0));
    sheet.set_cell_value(1, 0, Value::number(50.0));
    sheet.set_cell_value(2, 0, Value::number(100.0));
    sheet.set_cell_value(0, 1, Value::number(0.0));
    sheet.set_cell_value(1, 1, Value::number(50.0));
    sheet.set_cell_value(2, 1, Value::number(100.0));
    for (ConditionalFormat& block : ReadBlocks(body)) {
      sheet.mutable_conditional_formats().push_back(std::move(block));
    }
  }

  std::vector<CFRangeCellMatches> Evaluate() {
    eval::EvalContext ctx(wb, wb.sheet(0), state);
    CFHost host;
    host.arena = &arena;
    host.registry = &eval::default_registry();
    host.eval_ctx = &ctx;
    CFCellRange range{};
    range.first = CellAddress{0, 0};
    range.last = CellAddress{2, 1};
    auto out = evaluate_cf_for_range(wb.sheet(0), range, host);
    EXPECT_TRUE(static_cast<bool>(out));
    return out.take();
  }
};

const CFMatch* MatchAt(const std::vector<CFRangeCellMatches>& all, std::uint32_t row, std::uint32_t col) {
  for (const CFRangeCellMatches& entry : all) {
    if (entry.cell.row == row && entry.cell.col == col && !entry.matches.empty()) {
      return &entry.matches[0];
    }
  }
  return nullptr;
}

TEST(CfThemeColor, ColorScaleThemeNotBlack) {
  Fixture f(kColorScaleXml);
  const auto results = f.Evaluate();
  const CFMatch* low = MatchAt(results, 0, 0);
  const CFMatch* high = MatchAt(results, 2, 0);
  ASSERT_NE(low, nullptr);
  ASSERT_NE(high, nullptr);
  ASSERT_TRUE(low->resolved_fill_color.has_value());
  ASSERT_TRUE(high->resolved_fill_color.has_value());
  EXPECT_EQ(*low->resolved_fill_color, Rgb(0x8E, 0xA9, 0xDB));   // accent1, tint 0.4
  EXPECT_EQ(*high->resolved_fill_color, Rgb(0xC6, 0x59, 0x11));  // accent2, tint -0.25
}

TEST(CfThemeColor, DataBarThemeTintFill) {
  Fixture f(kDataBarXml);
  const auto results = f.Evaluate();
  const CFMatch* match = MatchAt(results, 1, 1);
  ASSERT_NE(match, nullptr);
  ASSERT_TRUE(match->data_bar_render.has_value());
  EXPECT_EQ(match->data_bar_render->fill, Rgb(0xBD, 0xD7, 0xEE));  // accent5, tint 0.6
}

TEST(CfThemeColor, IndexedAndAutoColors) {
  Fixture f(
      "<conditionalFormatting sqref=\"A1:A3\"><cfRule type=\"colorScale\" priority=\"1\"><colorScale>"
      "<cfvo type=\"min\"/><cfvo type=\"max\"/><color indexed=\"2\"/><color auto=\"1\"/>"
      "</colorScale></cfRule></conditionalFormatting>"
      "<conditionalFormatting sqref=\"B1:B3\"><cfRule type=\"dataBar\" priority=\"2\"><dataBar>"
      "<cfvo type=\"min\"/><cfvo type=\"max\"/><color indexed=\"4\"/></dataBar></cfRule></conditionalFormatting>");
  const auto results = f.Evaluate();
  const CFMatch* low = MatchAt(results, 0, 0);
  const CFMatch* high = MatchAt(results, 2, 0);
  ASSERT_NE(low, nullptr);
  ASSERT_NE(high, nullptr);
  EXPECT_EQ(*low->resolved_fill_color, Rgb(0xFF, 0x00, 0x00));
  EXPECT_EQ(*high->resolved_fill_color, Rgb(0xFF, 0xFF, 0xFF));  // automatic fill colour is the window colour
  const CFMatch* bar = MatchAt(results, 1, 1);
  ASSERT_NE(bar, nullptr);
  EXPECT_EQ(bar->data_bar_render->fill, Rgb(0x00, 0x00, 0xFF));
}

TEST(CfThemeColor, LiteralRgbIsUntouched) {
  Fixture f(
      "<conditionalFormatting sqref=\"A1:A3\"><cfRule type=\"colorScale\" priority=\"1\"><colorScale>"
      "<cfvo type=\"min\"/><cfvo type=\"max\"/><color rgb=\"FF102030\"/><color rgb=\"FF405060\"/>"
      "</colorScale></cfRule></conditionalFormatting>");
  const auto results = f.Evaluate();
  EXPECT_EQ(*MatchAt(results, 0, 0)->resolved_fill_color, Rgb(0x10, 0x20, 0x30));
  EXPECT_EQ(*MatchAt(results, 2, 0)->resolved_fill_color, Rgb(0x40, 0x50, 0x60));
}

TEST(CfThemeColor, ThemeEditIsReflected) {
  Fixture f(kColorScaleXml);
  ThemeColors colors = f.wb.load_theme().theme.colors;
  colors[4] = 0xFF204060U;  // accent1 (clrScheme order)
  ASSERT_TRUE(static_cast<bool>(f.wb.set_theme_colors(colors)));
  const auto results = f.Evaluate();
  const Color got = *MatchAt(results, 0, 0)->resolved_fill_color;
  EXPECT_NE(got, Rgb(0x8E, 0xA9, 0xDB));
  EXPECT_NE(got, Rgb(0, 0, 0));
  // A lighter tint of a dark accent is lighter than the accent itself.
  EXPECT_GT(got.b, 0x60);
}

TEST(CfThemeColor, WriterKeepsThemeAttributesOnRoundTrip) {
  const auto first = ReadBlocks(std::string(kColorScaleXml) + kDataBarXml);
  ASSERT_EQ(first.size(), 2u);
  const std::string written = io::write_conditional_formattings(first, 0);
  EXPECT_NE(written.find("theme=\"4\""), std::string::npos);
  EXPECT_NE(written.find("theme=\"8\""), std::string::npos);
  EXPECT_NE(written.find("tint=\""), std::string::npos);

  const auto second = ReadBlocks(written);
  ASSERT_EQ(second.size(), 2u);
  const ColorScaleSpec& a = *first[0].rules[0].color_scale;
  const ColorScaleSpec& b = *second[0].rules[0].color_scale;
  ASSERT_EQ(b.colors.size(), 2u);
  EXPECT_EQ(b.colors[0].spec.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(b.colors[0].spec.theme, 4u);
  EXPECT_DOUBLE_EQ(b.colors[0].spec.tint, a.colors[0].spec.tint);
  EXPECT_EQ(b.colors[1].spec.theme, 5u);
  EXPECT_LT(b.colors[1].spec.tint, 0.0);
  const DataBarSpec& bar = *second[1].rules[0].data_bar;
  EXPECT_EQ(bar.fill.spec.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(bar.fill.spec.theme, 8u);
  EXPECT_DOUBLE_EQ(bar.fill.spec.tint, first[1].rules[0].data_bar->fill.spec.tint);
}

TEST(CfThemeColor, IndexedAndAutoSurviveWriterRoundTrip) {
  const auto first = ReadBlocks(
      "<conditionalFormatting sqref=\"A1:A3\"><cfRule type=\"colorScale\" priority=\"1\"><colorScale>"
      "<cfvo type=\"min\"/><cfvo type=\"max\"/><color indexed=\"2\"/><color auto=\"1\"/>"
      "</colorScale></cfRule></conditionalFormatting>");
  const auto second = ReadBlocks(io::write_conditional_formattings(first, 0));
  ASSERT_EQ(second.size(), 1u);
  const auto& colors = second[0].rules[0].color_scale->colors;
  ASSERT_EQ(colors.size(), 2u);
  EXPECT_EQ(colors[0].spec.kind, ColorSpec::Kind::kIndexed);
  EXPECT_EQ(colors[0].spec.indexed, 2u);
  EXPECT_EQ(colors[1].spec.kind, ColorSpec::Kind::kAuto);
}

TEST(CfThemeColor, PackageSaveLoadKeepsThemeAttributes) {
  Fixture f(kColorScaleXml);
  auto bytes = io::write_ooxml(f.wb);
  ASSERT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  ASSERT_TRUE(static_cast<bool>(read));
  const auto& blocks = read.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 1u);
  const auto& colors = blocks[0].rules[0].color_scale->colors;
  ASSERT_EQ(colors.size(), 2u);
  EXPECT_EQ(colors[0].spec.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(colors[0].spec.theme, 4u);
  EXPECT_GT(colors[0].spec.tint, 0.39);
  EXPECT_EQ(colors[1].spec.theme, 5u);
  EXPECT_LT(colors[1].spec.tint, -0.24);
}

}  // namespace
}  // namespace formulon::cf
