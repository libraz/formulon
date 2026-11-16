//
// SUMIF / AVERAGEIF read their value range anchored at its top-left cell and
// shaped like the criteria range, whatever the value range's own size. The
// *IFS variants keep rejecting a shape mismatch with #VALUE!.

#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
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

// A1 "x", B2 "x", C1 1, D2 20: a 2x2 criteria range whose matches sit on
// the diagonal, so only a 2x2 value range anchored at C1 reaches D2.
Workbook DiagonalWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::text("x"));
  s.set_cell_value(1, 1, Value::text("x"));
  s.set_cell_value(0, 2, Value::number(1.0));
  s.set_cell_value(1, 3, Value::number(20.0));
  return wb;
}

void ExpectNumber(const Value& v, double expected) {
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), expected);
}

TEST(SumIfRangeResize, ColumnSumRangeTakesCriteriaShape) {
  const Workbook wb = DiagonalWorkbook();
  ExpectNumber(EvalSourceIn("=SUMIF(A1:B2,\"x\",C1:C2)", wb, wb.sheet(0)), 21.0);
}

TEST(SumIfRangeResize, LargerSumRangeShrinksToCriteriaShape) {
  const Workbook wb = DiagonalWorkbook();
  ExpectNumber(EvalSourceIn("=SUMIF(A1:B2,\"x\",C1:C5)", wb, wb.sheet(0)), 21.0);
}

TEST(SumIfRangeResize, SingleCellSumRangeExtends) {
  const Workbook wb = DiagonalWorkbook();
  ExpectNumber(EvalSourceIn("=SUMIF(A1:B2,\"x\",C1)", wb, wb.sheet(0)), 21.0);
}

TEST(SumIfRangeResize, AverageRangeTakesCriteriaShape) {
  const Workbook wb = DiagonalWorkbook();
  ExpectNumber(EvalSourceIn("=AVERAGEIF(A1:B2,\"x\",C1:C2)", wb, wb.sheet(0)), 10.5);
}

TEST(SumIfRangeResize, IfsVariantsKeepShapeMismatchError) {
  const Workbook wb = DiagonalWorkbook();
  for (const char* src : {"=SUMIFS(C1:C2,A1:B2,\"x\")", "=AVERAGEIFS(C1:C2,A1:B2,\"x\")", "=MAXIFS(C1:C2,A1:B2,\"x\")",
                          "=MINIFS(C1:C2,A1:B2,\"x\")"}) {
    const Value v = EvalSourceIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << src;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << src;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
