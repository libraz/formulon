// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the paired sum-of-products family:
// SUMX2PY2, SUMX2MY2, and SUMXMY2. These share the pairwise shape /
// error / non-numeric-drop rules with the regression family
// (CORREL et al.) but use the opposite argument order: the SUMX
// series is `(array_x, array_y)`, whereas CORREL is `(known_y,
// known_x)`. The order matters for SUMX2MY2 because `x^2 - y^2` is
// antisymmetric.
//
// Reference dataset used for the main numeric cases:
//   x = [1, 2, 3]
//   y = [4, 5, 6]
// Derived:
//   Σ (x² + y²)  = (1+16) + (4+25) + (9+36) = 91
//   Σ (x² - y²)  = (1-16) + (4-25) + (9-36) = -63
//   Σ (x - y)²   = 9 + 9 + 9 = 27

#include <cmath>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
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

Value EvalSource(std::string_view src) {
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
  return evaluate(*root, eval_arena);
}

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
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// SUMX2PY2
// ---------------------------------------------------------------------------

TEST(SumXY, SUMX2PY2ReferenceLiteral) {
  const Value v = EvalSource("=SUMX2PY2({1;2;3}, {4;5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 91.0);
}

TEST(SumXY, SUMX2PY2ShapeMismatchIsNA) {
  const Value v = EvalSource("=SUMX2PY2({1;2;3}, {4;5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(SumXY, SUMX2PY2EmptyPairSetIsNA) {
  // Neither column has a numeric cell, so every pair is dropped. Excel's
  // SUMX family returns #N/A for the empty case.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("y"));
  const Value v = EvalSourceIn("=SUMX2PY2(A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(SumXY, SUMX2PY2ErrorInXPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=SUMX2PY2(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(SumXY, SUMX2PY2SkipsNonNumericPair) {
  // A = [1, "x", 3], B = [4, 5, 6]. Only pairs (1,4) and (3,6) survive.
  // Σ (x² + y²) = (1+16) + (9+36) = 62.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=SUMX2PY2(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 62.0);
}

TEST(SumXY, SUMX2PY2ArityUnder) {
  const Value v = EvalSource("=SUMX2PY2({1;2;3})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(SumXY, SUMX2PY2ArityOver) {
  const Value v = EvalSource("=SUMX2PY2({1;2;3}, {4;5;6}, {7;8;9})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SUMX2MY2
// ---------------------------------------------------------------------------

TEST(SumXY, SUMX2MY2ReferenceLiteral) {
  const Value v = EvalSource("=SUMX2MY2({1;2;3}, {4;5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -63.0);
}

TEST(SumXY, SUMX2MY2OrderMatters) {
  // Swap x and y: Σ (y² - x²) = 63. Confirms we do not accidentally
  // treat the first argument as y.
  const Value v = EvalSource("=SUMX2MY2({4;5;6}, {1;2;3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 63.0);
}

TEST(SumXY, SUMX2MY2ShapeMismatchIsNA) {
  const Value v = EvalSource("=SUMX2MY2({1;2;3}, {4;5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// SUMXMY2
// ---------------------------------------------------------------------------

TEST(SumXY, SUMXMY2ReferenceLiteral) {
  const Value v = EvalSource("=SUMXMY2({1;2;3}, {4;5;6})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 27.0);
}

TEST(SumXY, SUMXMY2SymmetricInArguments) {
  // (x - y)^2 == (y - x)^2, so swapping must not change the result.
  const Value v = EvalSource("=SUMXMY2({4;5;6}, {1;2;3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 27.0);
}

TEST(SumXY, SUMXMY2ShapeMismatchIsNA) {
  const Value v = EvalSource("=SUMXMY2({1;2;3}, {4;5})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(SumXY, SUMXMY2SkipsNonNumericPair) {
  // A = [1, "x", 3], B = [4, 5, 6]. Surviving pairs: (1,4) and (3,6).
  // (1-4)^2 + (3-6)^2 = 9 + 9 = 18.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(6.0));
  const Value v = EvalSourceIn("=SUMXMY2(A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 18.0);
}

// ---------------------------------------------------------------------------
// Registry pin
// ---------------------------------------------------------------------------
//
// These three ride the lazy-dispatch table in `tree_walker.cpp`, not the
// eager registry, so they do NOT appear in `FunctionRegistry::lookup`.
// The invariant is instead enforced by `RegistryCatalog` - each name
// appears in `tools/catalog/functions.txt`. We still pin a smoke test
// to make sure the parser can resolve the call without #NAME?.

TEST(SumXY, RegistryPin) {
  const Value p = EvalSource("=SUMX2PY2({1},{1})");
  EXPECT_FALSE(p.is_error() && p.as_error() == ErrorCode::Name);
  const Value m = EvalSource("=SUMX2MY2({1},{1})");
  EXPECT_FALSE(m.is_error() && m.as_error() == ErrorCode::Name);
  const Value d = EvalSource("=SUMXMY2({1},{1})");
  EXPECT_FALSE(d.is_error() && d.as_error() == ErrorCode::Name);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
