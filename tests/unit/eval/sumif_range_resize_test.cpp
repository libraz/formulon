//
// SUMIF / AVERAGEIF read their value range anchored at its top-left cell and
// shaped like the criteria range, whatever the value range's own size. The
// *IFS variants keep rejecting a shape mismatch with #VALUE!.

#include <cstdint>
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

// A1:A5 = "x","y","x",(blank),"x"; B1:B5 = 10..50.
Workbook KeyedColumns() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const char* keys[] = {"x", "y", "x", nullptr, "x"};
  for (std::uint32_t r = 0; r < 5U; ++r) {
    if (keys[r] != nullptr) {
      s.set_cell_value(r, 0, Value::text(keys[r]));
    }
    s.set_cell_value(r, 1, Value::number(10.0 * (r + 1U)));
  }
  return wb;
}

// A reference a call returns is read as the same typed range: its value range
// is anchored and resized, its whole-axis tail counts, and its declared shape
// is compared.
TEST(SumIfRangeResize, ReturnedReferencesReadAsTypedRanges) {
  const Workbook wb = KeyedColumns();
  const struct {
    const char* formula;
    double expected;
  } numbers[] = {
      {"=COUNTIF(INDIRECT(\"A:A\"),\"\")", 1048572.0},
      {"=SUMIF(A1:A3,\"x\",OFFSET(B1,0,0))", 40.0},
      {"=COUNTIFS(INDIRECT(\"A:A\"),\"x\",INDIRECT(\"B:B\"),\">0\")", 3.0},
      {"=SUMIF(A1:A5,\"x\",INDEX(B:B,1))", 90.0},
      {"=AVERAGEIF(A1:A3,\"x\",INDIRECT(\"B1\"))", 20.0},
      {"=SUMIF(INDIRECT(\"A:A\"),\"\",B:B)", 40.0},
      {"=SUMIFS(OFFSET(B1,0,0,5),OFFSET(A1,0,0,5),\"x\")", 90.0},
      {"=SUMIF(OFFSET(A1,0,0,5),\"x\",OFFSET(B1:B2,1,0))", 60.0},
      {"=MAXIFS(INDIRECT(\"B1:B5\"),A1:A5,\"x\")", 50.0},
      {"=SUMIFS(B1:B5,INDIRECT(\"A1:A5\"),\"x\",OFFSET(A1,0,1,5),\">15\")", 80.0},
  };
  for (const auto& c : numbers) {
    SCOPED_TRACE(c.formula);
    ExpectNumber(EvalSourceIn(c.formula, wb, wb.sheet(0)), c.expected);
  }
  for (const char* src : {"=COUNTIFS(INDIRECT(\"A:A\"),\"x\",B1:B5,\">0\")",
                          "=COUNTIFS(OFFSET(A1,0,0,3),\"x\",OFFSET(B1,0,0,1,3),\">0\")"}) {
    const Value v = EvalSourceIn(src, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << src;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << src;
  }
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
