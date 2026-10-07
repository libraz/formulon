//
// The four sheet / book quoting rules and the A1 / R1C1 display they drive,
// checked against Mac Excel 365 ja-JP display: local sheets quote by
// notation, while the book and sheet parts of a cross-workbook qualifier
// follow their own, narrower rule.

#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

struct LocalRow {
  const char* name;
  bool quoted_a1;
  bool quoted_r1c1;
};

// Host workbook with sheets of these names: A1 display quotes the cell,
// digit, R/C and bool shapes; R1C1 display drops the quotes of the A1 cell
// shapes only.
constexpr LocalRow kLocalRows[] = {
    {"A1", true, false},    {"S2", true, false},     {"XFD1048576", true, false}, {"XFE1", false, false},
    {"R1C1", true, true},   {"R", true, true},       {"C", true, true},           {"r2c3", true, true},
    {"RC", true, true},     {"R2C", true, true},     {"RC3", true, true},         {"C7", true, true},
    {"TRUE", true, true},   {"false", true, true},   {"2024", true, true},        {"a.b", false, false},
    {"Data", false, false}, {"Sheet 1", true, true}, {"It's", true, true},        {"", true, true},
};

TEST(SheetQuoting, LocalPredicatesFollowMeasuredDisplay) {
  for (const LocalRow& row : kLocalRows) {
    EXPECT_EQ(local_sheet_needs_quoting_a1(row.name), row.quoted_a1) << row.name;
    EXPECT_EQ(local_sheet_needs_quoting_r1c1(row.name), row.quoted_r1c1) << row.name;
  }
}

TEST(SheetQuoting, ExternalBookPredicate) {
  // Bare and displayed bare even when entered quoted.
  for (const char* book : {"Book.xlsx", "2024.xlsx", "1.xlsx", "a.b.c.xlsx", "my_book.xlsx",
                           "\xE5\xA3\xB2\xE4\xB8\x8A.xlsx",  // 売上.xlsx
                           "\xEF\xBD\x81\xEF\xBD\x82.xlsx",  // ａｂ.xlsx
                           "Book", "A1.xlsx", "TRUE.xlsx"}) {
    EXPECT_FALSE(external_book_needs_quoting(book)) << book;
  }
  // Rejected bare, so quoted on display.
  for (const char* book :
       {"a b.xlsx", "a-b.xlsx", "a(b.xlsx", "a&b.xlsx", "a+b.xlsx", "a,b.xlsx", "a$b.xlsx", "a#b.xlsx", "a%b.xlsx",
        "a@b.xlsx", "a~b.xlsx", "a=b.xlsx", "a;b.xlsx", "a^b.xlsx", "a{b.xlsx", "a!b.xlsx"}) {
    EXPECT_TRUE(external_book_needs_quoting(book)) << book;
  }
}

TEST(SheetQuoting, ExternalSheetPredicate) {
  // Not the local rule: cell, R1C1 and bool shapes stay bare after `]`.
  for (const char* sheet :
       {"S2", "A1", "R1C1", "XFD1048576", "XFE1", "TRUE", "R", "C", "a.b", "_x", "\xE5\xA3\xB2\xE4\xB8\x8A"}) {  // 売上
    EXPECT_FALSE(external_sheet_needs_quoting(sheet)) << sheet;
  }
  for (const char* sheet : {"2024", "a b", "a-b"}) {
    EXPECT_TRUE(external_sheet_needs_quoting(sheet)) << sheet;
  }
}

struct DisplayRow {
  const char* entered;
  const char* a1;
  const char* r1c1;  // at host A1
};

constexpr DisplayRow kLocalDisplayRows[] = {
    {"=S2!A1", "'S2'!A1", "S2!RC"},
    {"=A1!A1", "'A1'!A1", "A1!RC"},
    {"=XFE1!A1", "XFE1!A1", "XFE1!RC"},
    {"=SUM(S2!A:A)", "SUM('S2'!A:A)", "SUM(S2!C)"},
    {"=2024!A1", "'2024'!A1", "'2024'!RC"},
    {"=R1C1!A1", "'R1C1'!A1", "'R1C1'!RC"},
    {"=R!A1", "'R'!A1", "'R'!RC"},
    {"=C!A1", "'C'!A1", "'C'!RC"},
    {"='TRUE'!A1", "'TRUE'!A1", "'TRUE'!RC"},
    {"=a.b!A1", "a.b!A1", "a.b!RC"},
    {"=S2!LocalName", "'S2'!LocalName", "S2!LocalName"},
    {"='S2'!A1", "'S2'!A1", "S2!RC"},
};

TEST(SheetQuoting, LocalDisplayInBothNotations) {
  for (const DisplayRow& row : kLocalDisplayRows) {
    Arena arena;
    Parser p(row.entered, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << row.entered;
    ASSERT_TRUE(p.errors().empty()) << row.entered;
    EXPECT_EQ(format_formula(*root), row.a1) << row.entered;
    EXPECT_EQ(format_formula_r1c1(*root, 0, 0), row.r1c1) << row.entered;
  }
}

TEST(SheetQuoting, R1C1IgnoresTheSourceQuotingHint) {
  // A binary-file reader marks every qualifier it decodes as quoted when
  // the A1 rule wants it; R1C1 display applies its own rule regardless.
  Arena arena;
  Reference cell;
  cell.sheet = "S2";
  cell.sheet_quoted = true;
  const AstNode* s2 = make_ref(arena, cell);
  EXPECT_EQ(format_formula(*s2), "'S2'!A1");
  EXPECT_EQ(format_formula_r1c1(*s2, 0, 0), "S2!RC");

  cell.sheet = "Data";
  const AstNode* data = make_ref(arena, cell);
  EXPECT_EQ(format_formula(*data), "'Data'!A1");
  EXPECT_EQ(format_formula_r1c1(*data, 0, 0), "Data!RC");
}

TEST(SheetQuoting, StorageRequoteOnlyForBareQualifiersNeedingQuotes) {
  struct Row {
    const char* entered;
    bool requote;
  };
  constexpr Row kRows[] = {
      {"=S2!A1", true},          {"='S2'!A1", false},        {"=2024!A1", true},           {"=Data!A1", false},
      {"=XFE1!A1", false},       {"=SUM(Data:S2!A1)", true}, {"=SUM(Data:Zed!A1)", false}, {"=S2!LocalName", true},
      {"=1+SUM(R!A1:B2)", true}, {"=A1+B2", false},
  };
  for (const Row& row : kRows) {
    Arena arena;
    Parser p(row.entered, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << row.entered;
    ASSERT_TRUE(p.errors().empty()) << row.entered;
    EXPECT_EQ(formula_needs_storage_requote(*root), row.requote) << row.entered;
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
