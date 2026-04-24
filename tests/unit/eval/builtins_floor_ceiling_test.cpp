// Copyright 2026 libraz. Licensed under the MIT License.
//
// Regression tests for the Mac Excel 365 asymmetric sign-mismatch rule on
// the legacy FLOOR / CEILING forms:
//
//   * number > 0 AND significance < 0 -> #NUM!
//   * number < 0 AND significance > 0 -> numeric (math floor / ceil)
//   * signs match                     -> numeric (magnitude operation)
//
// See oracle reference `tests/oracle/cases/floor_ceiling_edges.yaml` and
// the sibling matching-sign cases in `tests/oracle/cases/math.yaml`.

#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Mirrors the EvalSource helper in builtins_math_test.cpp: parses `src`
// and evaluates the AST through the default registry. Arenas are reset
// per call to prevent cross-test payload contamination.
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

// ---------------------------------------------------------------------------
// FLOOR sign-mismatch (pos-num, neg-sig -> #NUM!)
// ---------------------------------------------------------------------------

TEST(FloorPosNumNegSig, ReturnsNumError) {
  const Value v = EvalSource("=FLOOR(10, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FloorPosNumNegSigFraction, ReturnsNumError) {
  const Value v = EvalSource("=FLOOR(7.1, -0.1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FloorBoolTrueNegSig, ReturnsNumError) {
  // TRUE coerces to 1, so this hits the (n > 0, s < 0) branch.
  const Value v = EvalSource("=FLOOR(TRUE, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FloorBoolFalseNegSig, ReturnsZero) {
  // FALSE coerces to 0 -> zero short-circuit must win over sign-mismatch.
  const Value v = EvalSource("=FLOOR(FALSE, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(FloorZeroNegSig, ReturnsZero) {
  // Explicit zero-num must short-circuit before the sign-mismatch check.
  const Value v = EvalSource("=FLOOR(0, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(FloorNegNumNegSig, ReturnsNumeric) {
  // Matching signs -> magnitude floor toward zero.
  const Value v = EvalSource("=FLOOR(-7.1, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -7.0);
}

TEST(FloorNegNumPosSig, ReturnsNumeric) {
  // Regression guard for the (neg-num, pos-sig) oracle case in
  // tests/oracle/cases/math.yaml (`floor_sign_mismatch_rounds_toward_minus_inf`):
  // math floor on the signed value -> toward -infinity.
  const Value v = EvalSource("=FLOOR(-4.3, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -5.0);
}

// ---------------------------------------------------------------------------
// CEILING sign-mismatch (pos-num, neg-sig -> #NUM!)
// ---------------------------------------------------------------------------

TEST(CeilingPosNumNegSig, ReturnsNumError) {
  const Value v = EvalSource("=CEILING(10, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(CeilingNegNumPosSig, ReturnsNumeric) {
  // Regression guard for the (neg-num, pos-sig) oracle case in
  // tests/oracle/cases/math.yaml (`ceiling_sign_mismatch_rounds_toward_plus_inf`):
  // math ceil on the signed value -> toward +infinity.
  const Value v = EvalSource("=CEILING(-4.3, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -4.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
