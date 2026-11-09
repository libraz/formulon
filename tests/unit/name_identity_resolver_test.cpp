// The wide name fold and the by-name resolver.
//
// `sheet_names::wide_equal` is the identity a written name is checked
// against: the mac-365-en_US text `=` collation at compare strength, the same
// for every profile. `sheet_names::resolve` looks a typed name up: an exact
// code-point match first, then a case-only match, then a wide-fold match
// only when exactly one name qualifies.

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "eval/collation.h"
#include "eval/declared_rect.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace sheet_names {
namespace {

struct NamePair {
  const char* first;
  const char* second;
  bool same;
};

ExcelProfile MacEnUs() {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id("mac-365-en_US", &profile));
  return profile;
}

bool CollateEqual(const std::string& a, const std::string& b) {
  return eval::collate(a, b, MacEnUs(), eval::CollationStrength::kCompare) == 0;
}

// Expected values are the en-US `=` results; Excel's rename refusals on the
// same pairs agree except where a ja-JP UI keeps the prolonged mark's width.
constexpr NamePair kPairs[] = {
    // Sheet names.
    {"Sheet2", "Sheet2", true},
    {"Sheet2", "SHEET2", true},
    {"Sheet2", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x92", true},  // Ｓｈｅｅｔ２
    {"Sheet2", "\xEF\xBD\x93\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x92", true},  // ｓｈｅｅｔ２
    {"Sheet2", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x99", false},  // Ｓｈｅｅｔ９
    {"ABC", "\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3", true},                                          // ＡＢＣ
    {"\xEF\xBC\xA1", "\xEF\xBD\x81", true},                                                         // Ａ / ａ
    {"\xC3\xA9", "\xC3\x89", true},                                                                 // é / É
    {"\xEF\xBD\xB6\xEF\xBE\x80", "\xE3\x82\xAB\xE3\x82\xBF", true},                                 // ｶﾀ / カタ
    {"\xEF\xBD\xB6\xEF\xBE\x80", "\xE3\x81\x8B\xE3\x81\x9F", true},                                 // ｶﾀ / かた
    {"\xE3\x82\xAB\xE3\x82\xBF", "\xE3\x81\x8B\xE3\x81\x9F", true},                                 // カタ / かた
    {"\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF", "\xE3\x81\xA7\xE3\x83\xBC\xE3\x81\x9F", true},  // データ / でーた
    {"\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", true},                                      // ｶﾞ / ガ
    {"\xE3\x82\xAC", "\xE3\x82\xAB\xE3\x82\x99", true},                                      // ガ / カ + U+3099
    {"\xE3\x82\xAC", "\xE3\x82\xAB\xE3\x82\x9B", false},                                     // ガ / カ゛
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBE\x80", "\xE3\x83\x87\xE3\x82\xBF", true},              // ﾃﾞﾀ / デタ
    {"\xEF\xBD\xB0", "\xE3\x83\xBC", true},                                                  // ｰ / ー
    {"\xEF\xBD\xB0", "\xEF\xBC\x8D", false},                                                 // ｰ / －
    {"-", "\xEF\xBC\x8D", true},                                                             // - / －
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF",
     true},  // ﾃﾞｰﾀ / データ
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xE3\x83\x87\xEF\xBD\xB0\xE3\x82\xBF", true},  // ﾃﾞｰﾀ / デｰタ
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xEF\xBE\x83\xEF\xBE\x9E\xE3\x83\xBC\xEF\xBE\x80",
     true},  // ﾃﾞｰﾀ / ﾃﾞーﾀ
    {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80", "\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80 ",
     false},  // trailing space
    // Collection names.
    {"\xE3\x82\xAB", "\xEF\xBD\xB6", true},               // カ / ｶ
    {"\xE3\x82\xAB", "\xE3\x81\x8B", true},               // カ / か
    {"\xE3\x82\xAB", "\xE3\x82\xAC", false},              // カ / ガ
    {"\xE3\x82\xAB", "\xE3\x82\xAB\xE3\x82\x9B", false},  // カ / カ゛
    {"Ab", "aB", true},
    {"AB", "\xEF\xBC\xA1\xEF\xBC\xA2", true},  // ＡＢ
    // Pivot names.
    {"\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
     "1",
     "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB\xEF\xBC\x91",
     true},  // ピボットテーブル1 / ピボットテーブル１
    {"\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
     "1",
     "\xE3\x81\xB4\xE3\x81\xBC\xE3\x81\xA3\xE3\x81\xA8\xE3\x81\xA6\xE3\x83\xBC\xE3\x81\xB6\xE3\x82\x8B"
     "1",
     true},  // ぴぼっとてーぶる1
    {"\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
     "1",
     "\xEF\xBE\x8B\xEF\xBE\x9F\xEF\xBE\x8E\xEF\xBE\x9E\xEF\xBD\xAF\xEF\xBE\x84\xEF\xBE\x83\xEF\xBD\xB0\xEF\xBE\x8C"
     "\xEF\xBE\x9E\xEF\xBE\x99"
     "1",
     true},  // ﾋﾟﾎﾞｯﾄﾃｰﾌﾞﾙ1
    {"\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
     "1",
     "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
     "9",
     false},  // ピボットテーブル9
    // Punctuation, spaces and apostrophes are significant; ß expands.
    {"a-b", "ab", false},
    {"a b", "ab", false},
    {"a'b", "ab", false},
    {"\xF0\x9F\x93\x8A", "\xF0\x9F\x93\x88", false},  // 📊 / 📈
    {"\xC3\x9F", "ss", true},                         // ß / ss
    {"\xE1\xBA\x9E", "ss", true},                     // ẞ / ss
    {"\xC3\x85", "A\xCC\x8A", true},                  // Å / A + combining ring
    {"\xC3\x84", "\xC3\xA4", true},                   // Ä / ä
    {"\xCE\xA3", "\xCF\x82", true},                   // Σ / ς
    {"\xE2\x84\xAA", "k", true},                      // Kelvin sign / k
    {"\xF0\x90\x90\x80", "\xF0\x90\x90\xA8", false},  // U+10400 / U+10428: supplementary planes compare by code point
    {"\xC4\xB0", "i", false},                         // İ / i
    {"\xC4\xB1", "I", false},                         // ı / I
    {"\xC3\x84", "\xC3\x96", false},                  // Ä / Ö
};

TEST(NameIdentityWideFold, MeasuredPairs) {
  for (const NamePair& p : kPairs) {
    SCOPED_TRACE(std::string(p.first) + " / " + p.second);
    EXPECT_EQ(wide_equal(p.first, p.second), p.same);
    EXPECT_EQ(wide_equal(p.second, p.first), p.same);
    EXPECT_EQ(CollateEqual(p.first, p.second), p.same);
  }
}

TEST(NameIdentityWideFold, AsciiAllPairsAgreeWithCollation) {
  for (int a = 0; a < 128; ++a) {
    for (int b = 0; b < 128; ++b) {
      const std::string x(1, static_cast<char>(a));
      const std::string y(1, static_cast<char>(b));
      EXPECT_EQ(wide_equal(x, y), CollateEqual(x, y)) << a << "," << b;
    }
  }
  // Control characters are ignorable under collation, so they do not separate names.
  std::string with_control = "a";
  with_control.push_back('\x01');
  with_control += "b";
  EXPECT_EQ(wide_equal(with_control, "ab"), CollateEqual(with_control, "ab"));
  std::string with_nul = "A";
  with_nul.push_back('\0');
  with_nul += "Sheet";
  std::string with_nul_folded = "a";
  with_nul_folded.push_back('\0');
  with_nul_folded += "sHEET";
  EXPECT_EQ(wide_equal(with_nul, with_nul_folded), CollateEqual(with_nul, with_nul_folded));
}

TEST(NameIdentityWideFold, IsProfileIndependent) {
  ExcelProfile win;
  ASSERT_TRUE(parse_excel_profile_id("win-365-en_US", &win));
  // Windows `=` does not fold width; identity still does.
  EXPECT_NE(eval::collate("Sheet1", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91", win,
                          eval::CollationStrength::kCompare),
            0);
  EXPECT_TRUE(wide_equal("Sheet1", "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91"));
}

TEST(NameIdentityWideFold, RejectsMalformedUtf8) {
  for (const std::string malformed : {"\xC0\x80", "\xC3", "\xE0\x80\x80", "\xED\xA0\x80", "\xF4\x90\x80\x80", "\x80"}) {
    EXPECT_FALSE(wide_equal(malformed, malformed));
  }
  EXPECT_FALSE(wide_equal("\xC3", "\xC3\xA4"));
}

// --- Resolver ---

constexpr const char* kFullWidthSheet1 = "\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91";
constexpr const char* kUpperFullWidthSheet1 =
    "\xEF\xBC\xB3\xEF\xBC\xA8\xEF\xBC\xA5\xEF\xBC\xA5\xEF\xBC\xB4\xEF\xBC\x91";

TEST(NameIdentityResolver, ExactThenCaseOnlyThenUniqueWideMatch) {
  const std::vector<std::string> names = {"Sheet1", kFullWidthSheet1};
  EXPECT_EQ(resolve(names, "Sheet1"), 0U);
  EXPECT_EQ(resolve(names, kFullWidthSheet1), 1U);
  EXPECT_EQ(resolve(names, "SHEET1"), 0U);
  EXPECT_EQ(resolve(names, kUpperFullWidthSheet1), 1U);
  // Ｓheet1 matches both names only under the wide fold.
  EXPECT_EQ(resolve(names, "\xEF\xBC\xB3heet1"), kNotFound);
  EXPECT_EQ(resolve(names, "Other"), kNotFound);
}

TEST(NameIdentityResolver, SingleWideCandidateResolves) {
  const std::vector<std::string> names = {"Sheet1", "Data"};
  EXPECT_EQ(resolve(names, kUpperFullWidthSheet1), 0U);
  EXPECT_EQ(resolve(names, "\xEF\xBD\x84\xEF\xBD\x81\xEF\xBD\x94\xEF\xBD\x81"), 1U);  // ｄａｔａ
}

TEST(NameIdentityResolver, AmbiguousProlongedMarkVariantsResolveOnlyExactly) {
  const std::vector<std::string> names = {"\xEF\xBE\x83\xEF\xBE\x9E\xEF\xBD\xB0\xEF\xBE\x80",  // ﾃﾞｰﾀ
                                          "\xE3\x83\x87\xE3\x83\xBC\xE3\x82\xBF"};             // データ
  EXPECT_EQ(resolve(names, names[0]), 0U);
  EXPECT_EQ(resolve(names, names[1]), 1U);
  EXPECT_EQ(resolve(names, "\xE3\x81\xA7\xE3\x83\xBC\xE3\x81\x9F"), kNotFound);  // でーた
}

TEST(NameIdentityResolver, EmptyListAndMalformedQuery) {
  EXPECT_EQ(resolve({}, "Sheet1"), kNotFound);
  EXPECT_EQ(resolve({"Sheet1"}, "\xC3"), kNotFound);
}

// --- Range chains ---

// The bounding pair of a `:` chain, its qualifiers resolved among `wb`'s sheets.
bool ChainSheetsAgree(const Workbook& wb, std::string_view formula) {
  Arena arena;
  const parser::AstNode* root = parser::parse_strict(formula, arena);
  EXPECT_NE(root, nullptr) << formula;
  if (root == nullptr) {
    return false;
  }
  parser::Reference lhs{};
  parser::Reference rhs{};
  EXPECT_TRUE(eval::declared_rect_endpoint_pair(*root, &lhs, &rhs, &wb)) << formula;
  return static_cast<bool>(eval::merge_sheet_qualifiers(lhs.sheet, rhs.sheet, &wb));
}

TEST(NameIdentityResolver, RangeChainWidthVariantsAgreeAndAThirdSheetConflicts) {
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet(0).name(), "Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("Sheet2")));
  // 'Ｓｈｅｅｔ１' resolves to Sheet1, so the first chain names one sheet.
  EXPECT_TRUE(
      ChainSheetsAgree(wb,
                       "Sheet1!A1:'\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91'!B2:"
                       "Sheet1!C3"));
  // A third sheet after the width variant is still the conflict the pair carries.
  EXPECT_FALSE(
      ChainSheetsAgree(wb,
                       "Sheet1!A1:'\xEF\xBC\xB3\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91'!B2:"
                       "Sheet2!C3"));
}

}  // namespace
}  // namespace sheet_names
}  // namespace formulon
