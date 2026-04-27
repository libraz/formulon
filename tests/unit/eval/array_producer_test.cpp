// Copyright 2026 libraz. Licensed under the MIT License.
//
// Direct unit tests for the array-context evaluator helpers
// `eval_node_as_array` and `eval_binop_array_ctx`. Both functions are
// declared in `eval/shape_ops_lazy.h` and are not yet wired into any
// production caller; the tests here are their only callers.
//
// The cases pin the dispatch surface end-to-end:
//   * Range / Ref / ArrayLiteral expansion via `resolve_range_arg`.
//   * Cellwise broadcasting (1x1 against shape, equal shapes, mismatch).
//   * Per-cell error-cell propagation vs. scalar-error short-circuit.
//   * Unary cellwise application.
//   * Scalar -> 1x1 wrapping fallback for non-range expressions.

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/shape_ops_lazy.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Bundles the per-test parser + evaluator arenas so tests can keep the
// returned `Value::Array` alive for assertions (the array's cells live in
// the evaluator arena, which must outlive the value).
struct EvalHarness {
  Arena parse_arena;
  Arena eval_arena;
  parser::AstNode* root = nullptr;

  bool parse(std::string_view src) {
    parser::Parser p(src, parse_arena);
    root = p.parse();
    return root != nullptr;
  }
};

// Calls `eval_node_as_array` against a parsed source with the supplied
// workbook + sheet bound. References inside `src` resolve through the
// usual `EvalContext::expand_range` path.
Value RunArrayCtx(EvalHarness* h, std::string_view src, const Workbook& wb, const Sheet& current) {
  EXPECT_TRUE(h->parse(src)) << "parse failed for: " << src;
  if (h->root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return eval_node_as_array(*h->root, h->eval_arena, default_registry(), ctx);
}

// Same, but for sources that need no live workbook (literals only).
Value RunArrayCtxNoSheet(EvalHarness* h, std::string_view src) {
  Workbook wb = Workbook::create();
  return RunArrayCtx(h, src, wb, wb.sheet(0));
}

// ---------------------------------------------------------------------------
// Range / Ref / ArrayLiteral expansion
// ---------------------------------------------------------------------------

TEST(ArrayProducer, RefProducesOneByOneArray) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(7.0));
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(1U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  EXPECT_EQ(Value::number(7.0), v.as_array_cells()[0]);
}

TEST(ArrayProducer, ColumnRangeOpProducesCorrectShape) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5U; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1U)));
  }
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A5", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(5U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  for (std::size_t i = 0; i < 5; ++i) {
    EXPECT_EQ(Value::number(static_cast<double>(i + 1)), v.as_array_cells()[i]) << "i=" << i;
  }
}

TEST(ArrayProducer, TwoDimensionalRangeOpRowMajorOrder) {
  // A1:B3 -> 3 rows, 2 cols. Fill with row-major value = row*2 + col.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3U; ++r) {
    for (std::uint32_t c = 0; c < 2U; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 2U + c)));
    }
  }
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:B3", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(3U, v.as_array_rows());
  EXPECT_EQ(2U, v.as_array_cols());
  for (std::size_t i = 0; i < 6; ++i) {
    EXPECT_EQ(Value::number(static_cast<double>(i)), v.as_array_cells()[i]) << "i=" << i;
  }
}

TEST(ArrayProducer, ArrayLiteralExpands) {
  EvalHarness h;
  // {1,2;3,4} parses as 2 rows of 2 cols (semicolon = row separator).
  const Value v = RunArrayCtxNoSheet(&h, "={1,2;3,4}");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(2U, v.as_array_rows());
  EXPECT_EQ(2U, v.as_array_cols());
  EXPECT_EQ(Value::number(1.0), v.as_array_cells()[0]);
  EXPECT_EQ(Value::number(2.0), v.as_array_cells()[1]);
  EXPECT_EQ(Value::number(3.0), v.as_array_cells()[2]);
  EXPECT_EQ(Value::number(4.0), v.as_array_cells()[3]);
}

// ---------------------------------------------------------------------------
// BinaryOp broadcasting
// ---------------------------------------------------------------------------

TEST(ArrayProducer, ComparisonRangeGreaterThanScalar) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5U; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1U)));
  }
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A5>2", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(5U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  const bool expected[] = {false, false, true, true, true};
  for (std::size_t i = 0; i < 5; ++i) {
    ASSERT_TRUE(v.as_array_cells()[i].is_boolean()) << "i=" << i;
    EXPECT_EQ(expected[i], v.as_array_cells()[i].as_boolean()) << "i=" << i;
  }
}

TEST(ArrayProducer, ArithmeticArrayTimesScalar) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5U; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1U)));
  }
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A5*2", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(5U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  for (std::size_t i = 0; i < 5; ++i) {
    ASSERT_TRUE(v.as_array_cells()[i].is_number()) << "i=" << i;
    EXPECT_DOUBLE_EQ(static_cast<double>((i + 1) * 2), v.as_array_cells()[i].as_number()) << "i=" << i;
  }
}

TEST(ArrayProducer, RangePlusRangeMatchingShapes) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A3+B1:B3", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(3U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  EXPECT_EQ(Value::number(11.0), v.as_array_cells()[0]);
  EXPECT_EQ(Value::number(22.0), v.as_array_cells()[1]);
  EXPECT_EQ(Value::number(33.0), v.as_array_cells()[2]);
}

TEST(ArrayProducer, RangePlusRangeMismatchedShapesIsScalarValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  for (std::uint32_t r = 0; r < 5U; ++r) {
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>(r + 10U)));
  }
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A3+B1:B5", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(ErrorCode::Value, v.as_error());
}

TEST(ArrayProducer, UnaryNegationCellwise) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(-2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=-A1:A3", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(3U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  EXPECT_EQ(Value::number(-1.0), v.as_array_cells()[0]);
  EXPECT_EQ(Value::number(2.0), v.as_array_cells()[1]);
  EXPECT_EQ(Value::number(-3.0), v.as_array_cells()[2]);
}

// ---------------------------------------------------------------------------
// Error propagation
// ---------------------------------------------------------------------------

TEST(ArrayProducer, ErrorCellPreservedPerCellInComparison) {
  // A2 is a #DIV/0! cell. The comparison should produce an Array whose
  // middle cell is the error verbatim while the surrounding cells stay
  // boolean.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A3>2", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(3U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  ASSERT_TRUE(v.as_array_cells()[0].is_boolean());
  EXPECT_FALSE(v.as_array_cells()[0].as_boolean());
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(ErrorCode::Div0, v.as_array_cells()[1].as_error());
  ASSERT_TRUE(v.as_array_cells()[2].is_boolean());
  EXPECT_TRUE(v.as_array_cells()[2].as_boolean());
}

TEST(ArrayProducer, ScalarErrorOperandShortCircuitsToScalar) {
  // A literal `#DIV/0!` parses as an `ErrorLiteral` AST node which goes
  // through `eval_node_as_array`'s scalar fallback. That fallback
  // forwards scalar errors verbatim (it does NOT wrap them in a 1x1
  // array), so the outer broadcaster sees `rhs.is_error()` on the raw
  // `Value` and short-circuits the whole expression to scalar `#DIV/0!`.
  // (Contrast: `(1/0)` would evaluate through `eval_binop_array_ctx`
  // first, which produces a 1x1 array containing the error - that path
  // would broadcast cellwise, not short-circuit.)
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=A1:A3+#DIV/0!", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(ErrorCode::Div0, v.as_error());
}

// ---------------------------------------------------------------------------
// Scalar fallback
// ---------------------------------------------------------------------------

TEST(ArrayProducer, ScalarLiteralWrapsToOneByOne) {
  EvalHarness h;
  const Value v = RunArrayCtxNoSheet(&h, "=42");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(1U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  EXPECT_EQ(Value::number(42.0), v.as_array_cells()[0]);
}

TEST(ArrayProducer, ScalarComparisonBothOneByOne) {
  EvalHarness h;
  const Value v = RunArrayCtxNoSheet(&h, "=5>3");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(1U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  ASSERT_TRUE(v.as_array_cells()[0].is_boolean());
  EXPECT_TRUE(v.as_array_cells()[0].as_boolean());
}

TEST(ArrayProducer, NestedBroadcastBooleanTimesScalar) {
  // (A1:A3>2)*1 -> [F,F,T] coerced to numbers via apply_arithmetic == [0,0,1].
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  EvalHarness h;
  const Value v = RunArrayCtx(&h, "=(A1:A3>2)*1", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(3U, v.as_array_rows());
  EXPECT_EQ(1U, v.as_array_cols());
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(0.0, v.as_array_cells()[0].as_number());
  ASSERT_TRUE(v.as_array_cells()[1].is_number());
  EXPECT_DOUBLE_EQ(0.0, v.as_array_cells()[1].as_number());
  ASSERT_TRUE(v.as_array_cells()[2].is_number());
  EXPECT_DOUBLE_EQ(1.0, v.as_array_cells()[2].as_number());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
