//
// A criterion text naming an error (`"#N/A"`, `"=#N/A"`, `"<>#N/A"`) matches
// cells by error code, the same way an error Value criterion does.

#include <cstddef>
#include <iterator>
#include <string>
#include <string_view>

#include "eval/criteria.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/locale_text.h"
#include "eval/tree_walker.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx = test::workbook_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// A1 #N/A, A2 #DIV/0!, A3 5.
Workbook ErrorColumnWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::error(ErrorCode::NA));
  s.set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  s.set_cell_value(2, 0, Value::number(5.0));
  return wb;
}

// ErrorColumnWorkbook plus A4 holding the text "#N/A".
Workbook ErrorColumnWithTextWorkbook() {
  Workbook wb = ErrorColumnWorkbook();
  wb.sheet(0).set_cell_text(3, 0, "#N/A");
  return wb;
}

void ExpectNumber(const Value& v, double expected) {
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), expected);
}

TEST(CriteriaErrorText, CountIfBareErrorTextCountsThatError) {
  const Workbook wb = ErrorColumnWorkbook();
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A3,\"#N/A\")", wb, wb.sheet(0)), 1.0);
}

TEST(CriteriaErrorText, CountIfEqErrorTextCountsThatError) {
  const Workbook wb = ErrorColumnWorkbook();
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A3,\"=#DIV/0!\")", wb, wb.sheet(0)), 1.0);
}

TEST(CriteriaErrorText, CountIfNotEqErrorTextCountsEveryOtherCell) {
  const Workbook wb = ErrorColumnWorkbook();
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A3,\"<>#N/A\")", wb, wb.sheet(0)), 2.0);
}

TEST(CriteriaErrorText, SumIfNotEqSumsOtherErrorCell) {
  // A1 is excluded; A2 matches and its error propagates out of the sum.
  const Workbook wb = ErrorColumnWorkbook();
  const Value v = EvalSourceIn("=SUMIF(A1:A3,\"<>#N/A\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(CriteriaErrorText, ErrorNameCriteriaMatchTheErrorCellNotTheText) {
  const Workbook wb = ErrorColumnWithTextWorkbook();
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"#N/A\")", wb, wb.sheet(0)), 1.0);
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"=#N/A\")", wb, wb.sheet(0)), 1.0);
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"<>#N/A\")", wb, wb.sheet(0)), 3.0);
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\">#N/A\")", wb, wb.sheet(0)), 0.0);
}

TEST(CriteriaErrorText, ErrorNameCriteriaIgnoreLetterCase) {
  const Workbook wb = ErrorColumnWithTextWorkbook();
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"#n/a\")", wb, wb.sheet(0)), 1.0);
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"=#n/a\")", wb, wb.sheet(0)), 1.0);
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"<>#n/a\")", wb, wb.sheet(0)), 3.0);
  ExpectNumber(EvalSourceIn("=COUNTIF(A1:A4,\"#div/0!\")", wb, wb.sheet(0)), 1.0);
}

TEST(CriteriaErrorText, EveryErrorNameMatchesItsOwnCode) {
  const ExcelProfile profile = default_excel_profile();
  const ScopedEvalProfile scope(profile);
  for (std::size_t i = 0; i < std::size(kErrorTable); ++i) {
    const ErrorCode code = static_cast<ErrorCode>(i);
    const std::string name(locale_error_text(code));
    const std::string eq_name = "=" + name;
    const std::string ne_name = "<>" + name;
    const ErrorCode other = code == ErrorCode::NA ? ErrorCode::Value : ErrorCode::NA;
    for (const std::string* text : {&name, &eq_name}) {
      const ParsedCriterion eq = parse_criterion(Value::text(*text));
      EXPECT_TRUE(matches_criterion(Value::error(code), eq, profile)) << *text;
      EXPECT_FALSE(matches_criterion(Value::error(other), eq, profile)) << *text;
      EXPECT_FALSE(matches_criterion(Value::number(5.0), eq, profile)) << *text;
    }
    const ParsedCriterion ne = parse_criterion(Value::text(ne_name));
    EXPECT_FALSE(matches_criterion(Value::error(code), ne, profile)) << name;
    EXPECT_TRUE(matches_criterion(Value::error(other), ne, profile)) << name;
    EXPECT_TRUE(matches_criterion(Value::number(5.0), ne, profile)) << name;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
