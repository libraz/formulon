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

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

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

// Bound-workbook variant for cases that need A1-style cell references
// (e.g. blank cell refs to MROUND, where the dispatcher's blank-scalar
// policy distinguishes literal-empty from Ref-to-blank).
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

// ---------------------------------------------------------------------------
// MROUND blank-scalar policy
// ---------------------------------------------------------------------------
//
// Mac Excel 365 distinguishes between a parser-injected literal-empty arg
// slot (`=MROUND(,5)` / `=MROUND(5,)`) and a Ref to a blank cell
// (`=MROUND(A1,B1)` with A1/B1 blank). The probe golden
// `tests/oracle/golden/lowrisk_probes.golden.json` records the full table.

TEST(MRoundBlankScalar, LiteralEmptyFirstArgYieldsNA) {
  const Value v = EvalSource("=MROUND(,5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(MRoundBlankScalar, LiteralEmptySecondArgYieldsNA) {
  const Value v = EvalSource("=MROUND(5,)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(MRoundBlankScalar, BothBlankRefsCoerceToZero) {
  // Neither A1 nor B1 has a value; both Refs resolve to Blank, which the
  // RejectLiteralEmpty policy lets through to MRound's normal coercion.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=MROUND(A1,B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MRoundBlankScalar, ZeroMultipleStillZero) {
  const Value v = EvalSource("=MROUND(0,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MRoundBlankScalar, NumberWithZeroMultipleStillZero) {
  const Value v = EvalSource("=MROUND(5,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MRoundBlankScalar, BasicSanityFifteenAndFive) {
  const Value v = EvalSource("=MROUND(15,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 15.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
