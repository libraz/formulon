//
// The date1904-sensitive financial family (DISC, PRICE, COUP*, AMOR*, ...) lifts element by element over array
// arguments, as Excel 365 does and as the calendar family does: the result spans the widest argument, a size-1 axis
// stretches, a position an argument cannot supply is #N/A, and each element keeps its own error.

#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

Value Eval(std::string_view src) {
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
  return evaluate(*root, eval_arena, default_registry());
}

double Number(std::string_view src) {
  const Value v = Eval(src);
  EXPECT_TRUE(v.is_number()) << src;
  return v.is_number() ? v.as_number() : 0.0;
}

ErrorCode Error(std::string_view src) {
  const Value v = Eval(src);
  EXPECT_TRUE(v.is_error()) << src;
  return v.is_error() ? v.as_error() : ErrorCode::Null;
}

TEST(FinancialDateArrayLift, AnArrayFirstArgumentLiftsToItsShape) {
  EXPECT_EQ(Number("=COLUMNS(DISC({45306,45306},DATE(2024,7,15),99,100))"), 2.0);
  EXPECT_EQ(Number("=ROWS(COUPNUM({45306;45306;45306},DATE(2026,11,15),2,0))"), 3.0);
  EXPECT_EQ(Number("=INDEX(DISC({45306,45306},DATE(2024,7,15),99,100),2)"),
            Number("=DISC(45306,DATE(2024,7,15),99,100)"));
}

TEST(FinancialDateArrayLift, ALaterArgumentLiftsToo) {
  EXPECT_EQ(Number("=INDEX(TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),{0.09,0.05}),2)"),
            Number("=TBILLPRICE(DATE(2008,3,31),DATE(2008,6,1),0.05)"));
}

TEST(FinancialDateArrayLift, TwoArraysBroadcastAndAShortAxisReadsNa) {
  EXPECT_EQ(Number("=COLUMNS(DISC({45306,45306},{45488;45488},99,100))"), 2.0);
  EXPECT_EQ(Number("=ROWS(DISC({45306,45306},{45488;45488},99,100))"), 2.0);
  EXPECT_EQ(Error("=INDEX(DISC({45306,45306,45306},{45488,45488},99,100),3)"), ErrorCode::NA);
}

TEST(FinancialDateArrayLift, EachElementKeepsItsOwnResult) {
  // The first settlement is after maturity (#NUM!), the second is valid, the third is an error element.
  EXPECT_EQ(Error("=INDEX(DISC({45600,45306,#DIV/0!},DATE(2024,7,15),99,100),1)"), ErrorCode::Num);
  EXPECT_EQ(Number("=INDEX(DISC({45600,45306,#DIV/0!},DATE(2024,7,15),99,100),2)"),
            Number("=DISC(45306,DATE(2024,7,15),99,100)"));
  EXPECT_EQ(Error("=INDEX(DISC({45600,45306,#DIV/0!},DATE(2024,7,15),99,100),3)"), ErrorCode::Div0);
}

TEST(FinancialDateArrayLift, ScalarCallsStayScalar) {
  EXPECT_TRUE(Eval("=DISC(45306,DATE(2024,7,15),99,100)").is_number());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
