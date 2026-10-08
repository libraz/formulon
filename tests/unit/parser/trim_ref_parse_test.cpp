//
// Parser and formatter tests for the trim-reference operators `.:`, `:.`
// and `.:.`. Each parses to the one-argument `_TRO_LEADING` /
// `_TRO_TRAILING` / `_TRO_ALL` call Excel stores, over the plain range, and
// prints back in operator form. The `.` they add must not disturb a
// decimal, a defined name or a sheet name that contains one.

#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/ast_dump.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {
namespace {

// Parses `src` (with `=`), failing the test on any diagnostic.
const AstNode* ParseOk(Arena& arena, std::string_view src) {
  Parser p(src, arena);
  const AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << src;
  EXPECT_TRUE(p.errors().empty()) << src;
  return root;
}

// The `_xlfn.` speller the writers apply, reduced to what these formulas call.
std::string SpellForTest(std::string_view name) {
  return name.substr(0, 5) == "_TRO_" ? "_xlfn." + std::string(name) : std::string(name);
}

struct TrimCase {
  const char* text;
  TrimRefMode mode;
  const char* range;  // the same range written with a plain `:`
};

constexpr TrimCase kTrimCases[] = {
    {"=A1:.A10", TrimRefMode::Trailing, "=A1:A10"},
    {"=A1.:A10", TrimRefMode::Leading, "=A1:A10"},
    {"=A1.:.A10", TrimRefMode::Both, "=A1:A10"},
    {"=A:.A", TrimRefMode::Trailing, "=A:A"},
    {"=A.:A", TrimRefMode::Leading, "=A:A"},
    {"=A.:.A", TrimRefMode::Both, "=A:A"},
    {"=A:.B", TrimRefMode::Trailing, "=A:B"},
    {"=$A:.$B", TrimRefMode::Trailing, "=$A:$B"},
    {"=1:.1", TrimRefMode::Trailing, "=1:1"},
    {"=1.:3", TrimRefMode::Leading, "=1:3"},
    {"=$1.:.$3", TrimRefMode::Both, "=$1:$3"},
    {"=A1:.C10", TrimRefMode::Trailing, "=A1:C10"},
    {"=$A$1:.$A$10", TrimRefMode::Trailing, "=$A$1:$A$10"},
    {"=Sheet2!A1:.A10", TrimRefMode::Trailing, "=Sheet2!A1:A10"},
    {"='My Sheet'!B2.:D9", TrimRefMode::Leading, "='My Sheet'!B2:D9"},
    {"=Sheet2!A:.A", TrimRefMode::Trailing, "=Sheet2!A:A"},
};

TEST(TrimRefParse, OperatorsParseToTheStoredCallOverThePlainRange) {
  for (const TrimCase& c : kTrimCases) {
    Arena arena;
    const AstNode* root = ParseOk(arena, c.text);
    ASSERT_NE(root, nullptr);
    ASSERT_EQ(trim_ref_call_mode(*root), c.mode) << c.text;
    EXPECT_EQ(root->as_call_name(), trim_ref_function_name(c.mode)) << c.text;
    const AstNode* plain = ParseOk(arena, c.range);
    ASSERT_NE(plain, nullptr);
    EXPECT_EQ(dump_sexpr(root->as_call_arg(0)), dump_sexpr(*plain)) << c.text;
  }
}

TEST(TrimRefParse, FormatterPrintsTheOperatorBack) {
  for (const TrimCase& c : kTrimCases) {
    Arena arena;
    const AstNode* root = ParseOk(arena, c.text);
    ASSERT_NE(root, nullptr);
    EXPECT_EQ("=" + format_formula(*root), c.text);
  }
}

TEST(TrimRefParse, StorageFormIsTheFutureFunctionCall) {
  const struct {
    const char* text;
    const char* stored;
  } cases[] = {
      {"=ROWS(A1:.A10)", "ROWS(_xlfn._TRO_TRAILING(A1:A10))"},
      {"=ROWS(A1.:A10)", "ROWS(_xlfn._TRO_LEADING(A1:A10))"},
      {"=ROWS(A1.:.A10)", "ROWS(_xlfn._TRO_ALL(A1:A10))"},
      {"=ROWS(A:.A)", "ROWS(_xlfn._TRO_TRAILING(A:A))"},
      {"=COLUMNS(1:.1)", "COLUMNS(_xlfn._TRO_TRAILING(1:1))"},
      {"=SUM(Sheet1!A1:.A10)", "SUM(_xlfn._TRO_TRAILING(Sheet1!A1:A10))"},
      {"=ROWS($A$1:.$A$10)", "ROWS(_xlfn._TRO_TRAILING($A$1:$A$10))"},
  };
  for (const auto& c : cases) {
    Arena arena;
    const AstNode* root = ParseOk(arena, c.text);
    ASSERT_NE(root, nullptr);
    EXPECT_EQ(format_formula_storage(*root, &SpellForTest), c.stored) << c.text;
  }
}

TEST(TrimRefParse, StoredCallsReadBackAsOperators) {
  EXPECT_EQ(spell_storage_operators(strip_storage_prefixes("=ROWS(_xlfn._TRO_TRAILING(A1:A10))")), "=ROWS(A1:.A10)");
  EXPECT_EQ(spell_storage_operators("=ROWS(_TRO_TRAILING(A1:A10))"), "=ROWS(A1:.A10)");
  EXPECT_EQ(spell_storage_operators("=ROWS(_TRO_LEADING(A1:A10))"), "=ROWS(A1.:A10)");
  EXPECT_EQ(spell_storage_operators("=ROWS(_TRO_ALL(A1:A10))"), "=ROWS(A1.:.A10)");
  EXPECT_EQ(spell_storage_operators("=ROWS(_TRO_TRAILING(A:A))"), "=ROWS(A:.A)");
  EXPECT_EQ(spell_storage_operators("=COLUMNS(_TRO_TRAILING(1:1))"), "=COLUMNS(1:.1)");
  EXPECT_EQ(spell_storage_operators("=SUM(_TRO_TRAILING(Sheet1!A1:A10))"), "=SUM(Sheet1!A1:.A10)");
  EXPECT_EQ(spell_storage_operators("=ROWS(_TRO_TRAILING($A$1:$A$10))"), "=ROWS($A$1:.$A$10)");
  EXPECT_EQ(spell_storage_operators("=SUM(_TRO_TRAILING(A1:C10))"), "=SUM(A1:.C10)");
}

TEST(TrimRefParse, DotsOutsideTheOperatorKeepTheirMeaning) {
  const struct {
    const char* text;
    const char* sexpr;
  } cases[] = {
      {"=1.5+.5", nullptr},
      {"=1.", nullptr},
      {"={1.5,.5}", nullptr},
      {"=Tax.Rate", "(name Tax.Rate)"},
      {"=Tax.Rate*2", nullptr},
      {"=Sheet.1!A1", nullptr},
      {"=SUM(Sheet.1!A1:A3)", nullptr},
      {"=A1:A10", nullptr},
  };
  for (const auto& c : cases) {
    Arena arena;
    const AstNode* root = ParseOk(arena, c.text);
    ASSERT_NE(root, nullptr);
    EXPECT_EQ(trim_ref_call_mode(*root), TrimRefMode::None) << c.text;
    if (c.sexpr != nullptr) {
      EXPECT_EQ(dump_sexpr(*root), c.sexpr) << c.text;
    }
  }
  Arena arena;
  const AstNode* sheet_dot = ParseOk(arena, "=Sheet.1!A1");
  ASSERT_NE(sheet_dot, nullptr);
  ASSERT_EQ(sheet_dot->kind(), NodeKind::Ref);
  EXPECT_EQ(sheet_dot->as_ref().sheet, "Sheet.1");
  const AstNode* decimal = ParseOk(arena, "=1.5+.5");
  ASSERT_NE(decimal, nullptr);
  ASSERT_EQ(decimal->kind(), NodeKind::BinaryOp);
  EXPECT_DOUBLE_EQ(decimal->as_binary_lhs().as_literal().as_number(), 1.5);
  EXPECT_DOUBLE_EQ(decimal->as_binary_rhs().as_literal().as_number(), 0.5);
}

TEST(TrimRefParse, TrimRefInsideLargerExpressions) {
  for (const char* text : {"=SUM(A1.:.A10)", "=ROWS(A1:.C10)&\"x\"&COLUMNS(A1:.C10)", "=-A1:.A3", "=A1:.A3+1",
                           "=INDEX(A:.A,2)", "=LET(r,A1:.A10,ROWS(r))"}) {
    Arena arena;
    const AstNode* root = ParseOk(arena, text);
    ASSERT_NE(root, nullptr);
    EXPECT_EQ("=" + format_formula(*root), text);
  }
}

TEST(TrimRefParse, RelativeShiftMovesTheRangeAndKeepsTheOperator) {
  Arena arena;
  const AstNode* root = ParseOk(arena, "=SUM(A1:.A10)+ROWS($B$2.:C3)+COLUMNS(1:.1)");
  ASSERT_NE(root, nullptr);
  const AstNode* shifted = shift_relative_refs(*root, arena, 1, 1);
  ASSERT_NE(shifted, nullptr);
  EXPECT_EQ(format_formula(*shifted), "SUM(B2:.B11)+ROWS($B$2.:D4)+COLUMNS(2:.2)");
}

TEST(FunctionValueParse, StoragePrefixIsDroppedOnIngestion) {
  EXPECT_EQ(strip_storage_prefixes("=TYPE(_xleta.SUM)"), "=TYPE(SUM)");
  EXPECT_EQ(strip_storage_prefixes("=_xlfn.LET(_xlpm.f,_xleta.ABS,_xlpm.f(-2))"), "=LET(f,ABS,f(-2))");
  EXPECT_EQ(strip_storage_prefixes("=CHOOSE(1,_xleta.SUM,_xleta.ABS)(5)"), "=CHOOSE(1,SUM,ABS)(5)");
  EXPECT_EQ(strip_storage_prefixes("=\"_xleta.SUM\""), "=\"_xleta.SUM\"");
}

TEST(FunctionValueParse, CallOfAComputedFunctionValueParsesAndPrintsBack) {
  for (const char* text : {"=CHOOSE(1,SUM,ABS)(5)", "=IF(TRUE,ABS,SUM)(-3)"}) {
    Arena arena;
    const AstNode* root = ParseOk(arena, text);
    ASSERT_NE(root, nullptr);
    ASSERT_EQ(root->kind(), NodeKind::LambdaCall) << text;
    EXPECT_EQ(root->as_lambda_call_callee().kind(), NodeKind::Call) << text;
    EXPECT_EQ("=" + format_formula(*root), text);
  }
}

TEST(FunctionValueParse, StorageFormSpellsListedNamesWithTheValuePrefix) {
  Arena arena;
  const AstNode* root = ParseOk(arena, "=LET(f,abs,f(-2))+TYPE(Rate)");
  ASSERT_NE(root, nullptr);
  const AstNode& let = root->as_binary_lhs();
  const std::vector<const AstNode*> values = {&let.as_let_binding_expr(0)};
  EXPECT_EQ(format_formula_storage(*root, &SpellForTest, nullptr, nullptr, &values),
            "LET(_xlpm.f,_xleta.ABS,_xlpm.f(-2))+TYPE(Rate)");
}

}  // namespace
}  // namespace parser
}  // namespace formulon
