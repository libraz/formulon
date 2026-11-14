//
// Compatibility Version decides how LEN, MID, FIND, SEARCH and REPLACE count:
// version 1 counts UTF-16 units, every other version code points, on every
// profile. LEFT and RIGHT count code points on every version. An unstamped
// book and a new book evaluate as version 2. Measured on Mac Excel 16.115
// (stamped 1 / stamped 2 / unstamped books); the Windows half rests on the
// win-365-ja_JP goldens for stamped-2 books.

#include <cstdint>
#include <optional>
#include <string>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// a 😀 b 😁 c: 7 UTF-16 units, 5 code points.
constexpr const char* kText =
    "a\xF0\x9F\x98\x80"
    "b\xF0\x9F\x98\x81"
    "c";

struct Results {
  double len = 0;
  std::string mid;
  double find = 0;
  double search = 0;
  std::string replace;
  std::string left3;
  std::string right3;
};

std::string Text(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row, 1U);
  return cell != nullptr && cell->cached_value.is_text() ? std::string(cell->cached_value.as_text()) : "<not text>";
}

double Number(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row, 1U);
  return cell != nullptr && cell->cached_value.is_number() ? cell->cached_value.as_number() : -1.0;
}

Results Evaluate(Workbook& wb) {
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 0U, 0U, kText)));
  const char* formulas[] = {"=LEN(A1)",          "=MID(A1,2,3)",           "=FIND(\"b\",A1)",
                            "=SEARCH(\"B\",A1)", "=REPLACE(A1,4,1,\"X\")", "=LEFT(A1,3)",
                            "=RIGHT(A1,3)"};
  for (std::uint32_t i = 0; i < 7U; ++i) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, i + 1U, 1U, formulas[i]))) << formulas[i];
  }
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  return Results{Number(wb, 1U), Text(wb, 2U), Number(wb, 3U), Number(wb, 4U),
                 Text(wb, 5U),   Text(wb, 6U), Text(wb, 7U)};
}

void ExpectCodePoints(const Results& r) {
  EXPECT_EQ(r.len, 5.0);
  EXPECT_EQ(r.mid,
            "\xF0\x9F\x98\x80"
            "b\xF0\x9F\x98\x81");
  EXPECT_EQ(r.find, 3.0);
  EXPECT_EQ(r.search, 3.0);
  EXPECT_EQ(r.replace,
            "a\xF0\x9F\x98\x80"
            "bXc");
}

void ExpectUtf16Units(const Results& r) {
  EXPECT_EQ(r.len, 7.0);
  EXPECT_EQ(r.mid,
            "\xF0\x9F\x98\x80"
            "b");  // UTF-16 units 2..4
  EXPECT_EQ(r.find, 4.0);
  EXPECT_EQ(r.search, 4.0);
  EXPECT_EQ(r.replace,
            "a\xF0\x9F\x98\x80X\xF0\x9F\x98\x81"
            "c");
}

void ExpectLeftRightCodePoints(const Results& r) {
  EXPECT_EQ(r.left3,
            "a\xF0\x9F\x98\x80"
            "b");
  EXPECT_EQ(r.right3,
            "b\xF0\x9F\x98\x81"
            "c");
}

class CompatibilityVersionEval : public ::testing::TestWithParam<ExcelProfile> {};

TEST_P(CompatibilityVersionEval, ANewBookCountsCodePoints) {
  Workbook wb = Workbook::create_for_profile(GetParam());
  EXPECT_EQ(wb.compatibility_version(), 2U);
  const Results r = Evaluate(wb);
  ExpectCodePoints(r);
  ExpectLeftRightCodePoints(r);
}

TEST_P(CompatibilityVersionEval, AnUnstampedBookCountsCodePoints) {
  Workbook wb = Workbook::create_for_profile(GetParam());
  wb.restore_compatibility_stamp(std::nullopt, std::nullopt);
  EXPECT_EQ(wb.compatibility_version(), 2U);
  ExpectCodePoints(Evaluate(wb));
}

TEST_P(CompatibilityVersionEval, AStampedVersion1BookCountsUtf16Units) {
  Workbook wb = Workbook::create_for_profile(GetParam());
  wb.restore_compatibility_stamp(1U, std::nullopt);
  const Results r = Evaluate(wb);
  ExpectUtf16Units(r);
  ExpectLeftRightCodePoints(r);
}

TEST_P(CompatibilityVersionEval, SettingVersion1CountsUtf16UnitsAndBackTo2CodePoints) {
  Workbook wb = Workbook::create_for_profile(GetParam());
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(1U)));
  ExpectUtf16Units(Evaluate(wb));
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(2U)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Number(wb, 1U), 5.0);  // the setter dirtied the formulas
}

// A stored 3 (nested arrays, not implemented) evaluates as 2.
TEST_P(CompatibilityVersionEval, AStampedVersion3BookEvaluatesAsVersion2) {
  Workbook wb = Workbook::create_for_profile(GetParam());
  wb.restore_compatibility_stamp(3U, 2U);
  EXPECT_EQ(wb.compatibility_version(), 3U);
  ExpectCodePoints(Evaluate(wb));
}

INSTANTIATE_TEST_SUITE_P(Profiles, CompatibilityVersionEval,
                         ::testing::Values(mac_365_ja_jp_profile(), mac_365_en_us_profile(), win_365_en_us_profile(),
                                           win_365_ja_jp_profile()),
                         [](const ::testing::TestParamInfo<ExcelProfile>& param_info) {
                           std::string id = excel_profile_id(param_info.param);
                           for (char& c : id) {
                             if (c == '-') {
                               c = '_';
                             }
                           }
                           return id;
                         });

// Version 1, a cut inside a pair: Excel returns a lone surrogate, which UTF-8
// text cannot hold, so the whole character is taken (accepted divergence
// compat_version_1_mid_replace_lone_surrogate).
TEST(CompatibilityVersionDivergence, AVersion1CutInsideAPairTakesTheWholeCharacter) {
  Workbook wb = Workbook::create();
  wb.restore_compatibility_stamp(1U, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 0U, 0U, kText)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=MID(A1,2,4)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(Text(wb, 1U),
            "\xF0\x9F\x98\x80"
            "b\xF0\x9F\x98\x81");
}

TEST(CompatibilityVersionSetter, RefusesOutOfDomainWithoutChange) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(wb.set_compatibility_version(0U).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.set_compatibility_version(4U).error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.compatibility_version(), 2U);
}

// 3 is stored as given and evaluates as 2.
TEST(CompatibilityVersionSetter, Version3IsStoredAndEvaluatesAs2) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(3U)));
  EXPECT_EQ(wb.compatibility_stamp(), std::optional<std::uint32_t>(3U));
  ExpectCodePoints(Evaluate(wb));
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(1U)));
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(3U)));
  EXPECT_EQ(wb.compatibility_stamp(), std::optional<std::uint32_t>(3U));
}

TEST(CompatibilityVersionSetter, ANewBookCarriesExcelsNewBookStamp) {
  const Workbook wb = Workbook::create();
  EXPECT_EQ(wb.compatibility_stamp(), std::optional<std::uint32_t>(2U));
  EXPECT_EQ(wb.compatibility_warn_below(), std::optional<std::uint32_t>(2U));
}

}  // namespace
}  // namespace formulon
