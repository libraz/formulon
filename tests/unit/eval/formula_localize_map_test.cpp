#include <cstdint>
#include <string>
#include <string_view>

#include "eval/formula_localize.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/token.h"
#include "utils/utf8_length.h"

namespace formulon {
namespace eval {
namespace {

constexpr ExcelProfile kGerman{ExcelHost::kMac365, ExcelLocale::kDeDE};

parser::TextRange units(std::uint32_t start, std::uint32_t end) {
  parser::TextRange range;
  range.start = start;
  range.end = end;
  return range;
}

TEST(FormulaLocalizeMap, MappedTextMatchesThePlainRewrite) {
  const std::string_view localized = "=WENN(A1>0;SUMME(B1:B3;1,5);WAHR)";
  EXPECT_EQ(canonicalize_formula_text_mapped(localized, kGerman).text, canonicalize_formula_text(localized, kGerman));
  const std::string_view canonical = "=IF(A1>0,SUM(B1:B3,1.5),TRUE)";
  EXPECT_EQ(localize_formula_text_mapped(canonical, kGerman).text, localize_formula_text(canonical, kGerman));
}

TEST(FormulaLocalizeMap, EndOfCanonicalTextMapsToEndOfInput) {
  const std::string_view localized = "=SUMME(1;";
  const MappedFormulaText mapped = canonicalize_formula_text_mapped(localized, kGerman);
  ASSERT_EQ(mapped.text, "=SUM(1,");
  const parser::TextRange range = mapped.source_range(localized, units(7, 7));
  EXPECT_EQ(range.start, 9U);
  EXPECT_EQ(range.end, 9U);
  EXPECT_EQ(range.line, 1U);
  EXPECT_EQ(range.column, 10U);
}

TEST(FormulaLocalizeMap, VerbatimTextKeepsItsInputOffset) {
  const std::string_view localized = "=WENN(A1>0;SUMME(B1:B3;\"x);0)";
  const MappedFormulaText mapped = canonicalize_formula_text_mapped(localized, kGerman);
  const auto quote = static_cast<std::uint32_t>(mapped.text.find('"'));
  const auto length = static_cast<std::uint32_t>(mapped.text.size());
  const parser::TextRange range = mapped.source_range(localized, units(quote, length));
  EXPECT_EQ(range.start, 23U);
  EXPECT_EQ(range.end, static_cast<std::uint32_t>(localized.size()));
  EXPECT_EQ(range.column, 24U);
}

TEST(FormulaLocalizeMap, RewrittenTokenWidensToTheInputToken) {
  const std::string_view localized = "=SUMME(1)";
  const MappedFormulaText mapped = canonicalize_formula_text_mapped(localized, kGerman);
  ASSERT_EQ(mapped.text, "=SUM(1)");
  const parser::TextRange name = mapped.source_range(localized, units(2, 3));
  EXPECT_EQ(name.start, 1U);
  EXPECT_EQ(name.end, 6U);
  const parser::TextRange paren = mapped.source_range(localized, units(4, 5));
  EXPECT_EQ(paren.start, 6U);
  EXPECT_EQ(paren.end, 7U);
}

TEST(FormulaLocalizeMap, OffsetsCountUtf16Units) {
  const std::string localized = "=\"\xC3\xA4\xF0\x9F\x99\x82\"&SUMME(1;";
  const MappedFormulaText mapped = canonicalize_formula_text_mapped(localized, kGerman);
  const std::uint32_t text_units = utf16_units_in(mapped.text);
  const std::uint32_t source_units = utf16_units_in(localized);
  const parser::TextRange end = mapped.source_range(localized, units(text_units, text_units));
  EXPECT_EQ(end.start, source_units);
  EXPECT_EQ(end.end, source_units);
  // `&` follows `=`, two quotes, one BMP and one supplementary character.
  const parser::TextRange amp = mapped.source_range(localized, units(6, 7));
  EXPECT_EQ(amp.start, 6U);
  EXPECT_EQ(amp.end, 7U);
  EXPECT_EQ(amp.column, 7U);
}

TEST(FormulaLocalizeMap, LineAndColumnFollowLineBreaks) {
  const std::string_view localized = "=SUMME(1;\r\n2;";
  const MappedFormulaText mapped = canonicalize_formula_text_mapped(localized, kGerman);
  const auto length = static_cast<std::uint32_t>(mapped.text.size());
  const parser::TextRange range = mapped.source_range(localized, units(length, length));
  EXPECT_EQ(range.start, static_cast<std::uint32_t>(localized.size()));
  EXPECT_EQ(range.line, 2U);
  EXPECT_EQ(range.column, 3U);
}

TEST(FormulaLocalizeMap, LocalizedTextMapsBackToCanonicalInput) {
  const std::string_view canonical = "=SUM(1,2)";
  const MappedFormulaText mapped = localize_formula_text_mapped(canonical, kGerman);
  ASSERT_EQ(mapped.text, "=SUMME(1;2)");
  const parser::TextRange two = mapped.source_range(canonical, units(9, 10));
  EXPECT_EQ(two.start, 7U);
  EXPECT_EQ(two.end, 8U);
  const parser::TextRange separator = mapped.source_range(canonical, units(8, 9));
  EXPECT_EQ(separator.start, 6U);
  EXPECT_EQ(separator.end, 7U);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
