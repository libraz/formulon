//
// Bare local sheet qualifiers whose name the tokenizer would otherwise read
// as a cell, a number or a bool (`S2!A1`, `2024!A1`), checked against Mac
// Excel 365 ja-JP entry and display.
//
// Pairwise model (every pair of levels across the three factors appears):
//   name   : S2, 2024, R1C1, R, C, A1, XFE1, TRUE, Data            (9)
//   form   : !A1, !$A$1, !A:A, !1:1, !A1:B2, !LocalName,
//            3-D first sheet (`SUM(<name>:Zed!A1)`),
//            3-D last sheet (`SUM(Data:<name>!A1)`)                (8)
//   space  : none, before `!`, after `!`                          (3)
// Rows are name x form, with space level `(name + form) % 3`; each name and
// each form then meets all three space levels, so the 72 rows cover every
// pair.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

constexpr std::string_view kNames[] = {"S2", "2024", "R1C1", "R", "C", "A1", "XFE1", "TRUE", "Data"};

enum class Form { kCell, kAbsoluteCell, kColumn, kRow, kArea, kLocalName, kSpanFirst, kSpanLast };
constexpr Form kForms[] = {Form::kCell, Form::kAbsoluteCell, Form::kColumn,    Form::kRow,
                           Form::kArea, Form::kLocalName,    Form::kSpanFirst, Form::kSpanLast};

enum class Space { kNone, kBeforeBang, kAfterBang };

// Measured A1 display: quoted when cell-shaped, digit-leading, R/C-shaped or
// a bool word; `XFE1` and `Data` stay bare.
bool QuotedInA1(std::string_view name) {
  return name != "XFE1" && name != "Data";
}

// Measured R1C1 display: the A1 cell shapes print bare.
bool QuotedInR1C1(std::string_view name) {
  return name == "2024" || name == "R1C1" || name == "R" || name == "C" || name == "TRUE";
}

std::string Qualify(std::string_view name, bool quoted) {
  return quoted ? "'" + std::string(name) + "'" : std::string(name);
}

std::string Bang(Space space) {
  switch (space) {
    case Space::kBeforeBang:
      return " !";
    case Space::kAfterBang:
      return "! ";
    case Space::kNone:
      break;
  }
  return "!";
}

struct Expectation {
  bool accepted = false;
  std::string a1;    // format_formula, no leading `=`
  std::string r1c1;  // at host A1; empty when not checked
};

const char* A1Tail(Form form) {
  switch (form) {
    case Form::kCell:
      return "A1";
    case Form::kAbsoluteCell:
      return "$A$1";
    case Form::kColumn:
      return "A:A";
    case Form::kRow:
      return "1:1";
    case Form::kArea:
      return "A1:B2";
    case Form::kLocalName:
      return "LocalName";
    case Form::kSpanFirst:
    case Form::kSpanLast:
      break;
  }
  return "A1";
}

const char* R1C1Tail(Form form) {
  switch (form) {
    case Form::kCell:
      return "RC";
    case Form::kAbsoluteCell:
      return "R1C1";
    case Form::kColumn:
      return "C";
    case Form::kRow:
      return "R";
    case Form::kArea:
      return "RC:R[1]C[1]";
    case Form::kLocalName:
      return "LocalName";
    case Form::kSpanFirst:
    case Form::kSpanLast:
      break;
  }
  return "RC";
}

std::string Formula(std::string_view name, Form form, Space space) {
  const std::string n(name);
  switch (form) {
    case Form::kSpanFirst:
      return "=SUM(" + n + ":Zed" + Bang(space) + "A1)";
    case Form::kSpanLast:
      return "=SUM(Data:" + n + Bang(space) + "A1)";
    default:
      return "=" + n + Bang(space) + A1Tail(form);
  }
}

Expectation Expected(std::string_view name, Form form, Space space) {
  Expectation e;
  // A space between a qualifier and its `!` is rejected for every name.
  if (space == Space::kBeforeBang) {
    return e;
  }
  const bool bool_word = name == "TRUE";
  if (form == Form::kSpanFirst) {
    // A cell-shaped first endpoint stays a cell (`S2:Zed!A1` is a range
    // over cell S2); a bool word is no sheet there.
    if (bool_word) {
      return e;
    }
    e.accepted = true;
    if (name == "S2" || name == "A1") {
      e.a1 = "SUM(" + std::string(name) + ":Zed!A1)";
    } else {
      e.a1 = "SUM(" + Qualify(std::string(name) + ":Zed", QuotedInA1(name)) + "!A1)";
    }
    return e;
  }
  if (form == Form::kSpanLast) {
    // A digit-leading last endpoint is no sheet span: `Data` ranged with
    // `'2024'!A1`. A bool word is accepted there.
    e.accepted = true;
    if (name == "2024") {
      e.a1 = "SUM(Data:'2024'!A1)";
    } else {
      e.a1 = "SUM(" + Qualify("Data:" + std::string(name), QuotedInA1(name)) + "!A1)";
    }
    return e;
  }
  if (bool_word) {
    return e;
  }
  e.accepted = true;
  e.a1 = Qualify(name, QuotedInA1(name)) + "!" + A1Tail(form);
  e.r1c1 = Qualify(name, QuotedInR1C1(name)) + "!" + R1C1Tail(form);
  return e;
}

TEST(LocalSheetQualifier, PairwiseEntryAndDisplay) {
  std::size_t rows = 0;
  for (std::size_t n = 0; n < std::size(kNames); ++n) {
    for (std::size_t f = 0; f < std::size(kForms); ++f) {
      const Space space = static_cast<Space>((n + f) % 3U);
      const std::string src = Formula(kNames[n], kForms[f], space);
      const Expectation want = Expected(kNames[n], kForms[f], space);
      Arena arena;
      Parser p(src, arena);
      const AstNode* root = p.parse();
      const bool accepted = root != nullptr && p.errors().empty();
      ++rows;
      ASSERT_EQ(accepted, want.accepted) << src;
      if (!accepted) {
        continue;
      }
      EXPECT_EQ(format_formula(*root), want.a1) << src;
      if (!want.r1c1.empty()) {
        EXPECT_EQ(format_formula_r1c1(*root, 0, 0), want.r1c1) << src;
      }
    }
  }
  EXPECT_EQ(rows, 72U);
}

TEST(LocalSheetQualifier, BoolWordIsRejectedInAnyCase) {
  for (const char* src : {"=TRUE!A1", "=true!A1", "=False!A1", "=SUM(TRUE!A1:B2)"}) {
    Arena arena;
    Parser p(src, arena);
    (void)p.parse();
    EXPECT_FALSE(p.errors().empty()) << src;
  }
}

TEST(LocalSheetQualifier, QuotedQualifierBeforeSpacedBangIsRejected) {
  Arena arena;
  Parser p("='S2' !A1", arena);
  (void)p.parse();
  EXPECT_FALSE(p.errors().empty());
}

TEST(LocalSheetQualifier, CellShapedQualifierKeepsItsSheet) {
  // The run is a sheet name, not the cell it looks like.
  Arena arena;
  Parser p("=A1!B2", arena);
  const AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  ASSERT_TRUE(p.errors().empty());
  ASSERT_EQ(root->kind(), NodeKind::Ref);
  EXPECT_EQ(root->as_ref().sheet, "A1");
  EXPECT_EQ(root->as_ref().col, 1U);
  EXPECT_EQ(root->as_ref().row, 1U);
}

TEST(LocalSheetQualifier, OrdinaryCellRangesAreUnchanged) {
  // The lookahead only fires on a run glued to `!`.
  for (const char* src : {"=A1:B2", "=SUM(S2:S5)", "=2024+1", "=Sheet1!A1:Sheet1!B2"}) {
    Arena arena;
    Parser p(src, arena);
    const AstNode* root = p.parse();
    ASSERT_NE(root, nullptr) << src;
    EXPECT_TRUE(p.errors().empty()) << src;
    EXPECT_EQ(format_formula(*root), std::string_view(src).substr(1)) << src;
  }
}

}  // namespace
}  // namespace parser
}  // namespace formulon
