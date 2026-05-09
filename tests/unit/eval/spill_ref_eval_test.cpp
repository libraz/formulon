// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Evaluator tests for the spilled-range `=A1#` operator.
//
// The parser produces `NodeKind::SpillRef`; `eval_node` resolves it via
// `Sheet::spill_region_at_anchor` and returns a `Value::Array`. Unlike
// SEQUENCE (which produces a Value::Array via `dispatch_array_result`),
// SpillRef reads an *already-committed* spill — these tests pre-seed the
// sheet's spill table directly and then evaluate the formula.

#include <cstdint>
#include <string_view>
#include <vector>

#include "cell.h"
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
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src`, evaluates it under `ctx`, returns the resulting `Value`.
// Both arenas live on the caller's stack so any text / array payloads
// remain readable for assertions.
Value EvalUnder(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, *eval_arena, default_registry(), ctx);
}

TEST(SpillRefEval, ResolvesCommittedSpill) {
  // Pre-commit a 3x1 spill at A1 with cells [10, 20, 30]; `=A1#` should
  // return a Value::Array with the same shape and cells.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* read_cells = v.as_array_cells();
  EXPECT_EQ(read_cells[0], Value::number(10.0));
  EXPECT_EQ(read_cells[1], Value::number(20.0));
  EXPECT_EQ(read_cells[2], Value::number(30.0));
}

TEST(SpillRefEval, NoSpillReturnsRef) {
  // No spill anchored at A1 -> `#REF!`. Mac Excel behaves this way: the `#`
  // operator only resolves to a real spill region; querying a cell that
  // never spilled is an explicit reference error.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(SpillRefEval, SumOverSpillRef) {
  // SUM consumes a SpillRef like any other range. Pre-commit [10, 20, 30]
  // at A1 and call `=SUM(A1#)`; result is 60.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SUM(A1#)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(SpillRefEval, SpillRefArithmeticBroadcasts1x1) {
  // `=A1#+B1#` over two 1x1 spills: the BinaryOp now broadcasts cellwise
  // when either operand is a Value::Array, so the result is a 1x1 Array
  // containing 50. Without a mutable_sheet bound the EvalContext leaves
  // the Array un-spilled (see `EvalContext::dispatch_array_result`'s
  // mutable_sheet gate), so callers see Value::Array directly here.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 1U, 1U, std::vector<Value>{Value::number(42.0)}));
  ASSERT_TRUE(sheet.commit_spill(0U, 1U, 1U, 1U, std::vector<Value>{Value::number(8.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#+B1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 50.0);
}

// ---------------------------------------------------------------------------
// Top-level array broadcasting (BinaryOp / UnaryOp)
// ---------------------------------------------------------------------------
//
// These pin the cellwise broadcast path through `eval_node`'s BinaryOp /
// UnaryOp dispatch. SpillRef is the most convenient way to introduce a
// Value::Array operand without needing to wire a recalc driver, but the
// behaviour exercised here is independent of the Array's source: any
// future array-producing builtin (FILTER, SORT, RANDARRAY) will go through
// the same code path.

TEST(SpillRefBroadcast, EqualShapeArraysAddCellwise) {
  // Two 3x1 spills added cellwise -> 3x1 Array of [11, 22, 33].
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));
  ASSERT_TRUE(sheet.commit_spill(0U, 1U, 3U, 1U,
                                 std::vector<Value>{Value::number(1.0), Value::number(2.0), Value::number(3.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#+B1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 22.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 33.0);
}

TEST(SpillRefBroadcast, ScalarOnRightBroadcastsAcrossArray) {
  // `=A1#+1` -> 3x1 Array of [11, 21, 31]. Scalar RHS broadcasts.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#+1", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 31.0);
}

TEST(SpillRefBroadcast, ScalarOnLeftBroadcastsAcrossArray) {
  // `=2*A1#` -> 3x1 Array of [20, 40, 60]. Scalar LHS broadcasts.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=2*A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 40.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 60.0);
}

TEST(SpillRefBroadcast, ShapeMismatchYieldsValue) {
  // 3x1 + 2x1: incompatible shapes (no 1x1 broadcast slot) collapse to
  // scalar `#VALUE!`. Mac Excel does NOT spill an array of `#VALUE!` cells
  // in this case — the whole expression short-circuits.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));
  ASSERT_TRUE(sheet.commit_spill(0U, 1U, 2U, 1U, std::vector<Value>{Value::number(1.0), Value::number(2.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#+B1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(SpillRefBroadcast, UnaryMinusBroadcastsAcrossArray) {
  // `=-A1#` -> 3x1 Array of [-10, -20, -30]. Top-level UnaryOp dispatch
  // routes through broadcast_unary when the operand is a Value::Array.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=-A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), -10.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), -20.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), -30.0);
}

TEST(SpillRefBroadcast, ConcatBroadcastsAcrossArray) {
  // `=A1#&"!"` -> 2x1 Array of ["10!", "20!"]. The `&` concat operator
  // broadcasts through the same per-cell helper.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, std::vector<Value>{Value::number(10.0), Value::number(20.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#&\"!\"", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "10!");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "20!");
}

TEST(SpillRefBroadcast, ComparisonBroadcastsAcrossArray) {
  // `=A1#>15` -> 3x1 Array of [FALSE, TRUE, TRUE]. Relational operators
  // broadcast cellwise through apply_comparison.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#>15", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0].as_boolean(), false);
  EXPECT_EQ(v.as_array_cells()[1].as_boolean(), true);
  EXPECT_EQ(v.as_array_cells()[2].as_boolean(), true);
}

TEST(SpillRefBroadcast, PerCellErrorPropagatesPerCell) {
  // A spill region containing a #DIV/0! cell propagates that error in just
  // the corresponding output cell, not the entire array. The remaining
  // cells compute normally. Mac Excel's behaviour: per-cell error transit.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(
      0U, 0U, 3U, 1U, std::vector<Value>{Value::number(10.0), Value::error(ErrorCode::Div0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#+1", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Div0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 31.0);
}

TEST(SpillRefBroadcast, ArrayMinusArrayBroadcastsViaSubtraction) {
  // Combined sanity check: subtraction operator over two equal-shape arrays.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, std::vector<Value>{Value::number(100.0), Value::number(50.0)}));
  ASSERT_TRUE(sheet.commit_spill(0U, 1U, 2U, 1U, std::vector<Value>{Value::number(40.0), Value::number(20.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#-B1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 60.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 30.0);
}

TEST(SpillRefEval, SpillRefAnchorOnPhantomReturnsRef) {
  // Pre-commit a 3x1 spill at A1; the cells at A2 / A3 are phantoms of the
  // anchor. `=A2#` queries a phantom address (not the anchor) and must
  // return `#REF!` because only anchors carry the spill region.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A2#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(SpillRefEval, UnboundContextReturnsName) {
  // Without a current sheet bound to the EvalContext there is nothing to
  // query; SpillRef surfaces `#NAME?` to mirror the corresponding Ref
  // resolution path.
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=A1#", &parse_arena, &eval_arena, test::mac_context());
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(SpillRefEval, QualifiedSpillRef) {
  // `=Sheet2!A1#` resolves the anchor against Sheet2's spill table even
  // though the EvalContext's `current_sheet_` is Sheet1.
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  Sheet& s1 = wb.sheet(0);
  Sheet& s2 = wb.sheet(1);
  ASSERT_TRUE(s2.commit_spill(0U, 0U, 2U, 1U, std::vector<Value>{Value::number(7.0), Value::number(11.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, s1, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=Sheet2!A1#", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0], Value::number(7.0));
  EXPECT_EQ(v.as_array_cells()[1], Value::number(11.0));
}

// ---------------------------------------------------------------------------
// LET passthrough
// ---------------------------------------------------------------------------
//
// `is_range_shaped_ast` now includes SpillRef, so a LET-bound SpillRef
// preserves its range provenance when consumed by range-aware callers
// (SUM, AVERAGE, COUNTIF, lookups). The companion fix in
// `resolve_range_arg` lets non-aggregator consumers expand the spill into
// their own row-major buffer.

TEST(SpillRefLet, SumOverLetBoundSpill) {
  // `=LET(s, A1#, SUM(s))` -> 60. The aggregator dispatcher's existing
  // SpillRef branch fires after LET passthrough substitutes the SpillRef
  // AST for the NameRef.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LET(s, A1#, SUM(s))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(SpillRefLet, BroadcastOverLetBoundSpill) {
  // `=LET(s, A1#, s+1)` -> 3x1 Array of [11, 21, 31]. The BinaryOp eager
  // path sees the NameRef return a Value::Array (the bound spill region's
  // cells) and routes through broadcast_binop.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LET(s, A1#, s+1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 21.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 31.0);
}

TEST(SpillRefLet, CountIfOverLetBoundSpill) {
  // COUNTIF goes through `resolve_range_arg`, which now has its own
  // SpillRef branch. Pre-commit a 4-cell spill with two values >= 20 and
  // verify COUNTIF resolves the binding through the new branch.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(
      0U, 0U, 4U, 1U,
      std::vector<Value>{Value::number(5.0), Value::number(20.0), Value::number(30.0), Value::number(15.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LET(s, A1#, COUNTIF(s, \">=20\"))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(SpillRefLet, TransitiveLetBindingPreservesShape) {
  // `=LET(s, A1#, t, s, SUM(t))` -> 60. The NameRef-on-NameRef transitivity
  // in LetBinding's `expr_for_binding` lookup propagates the SpillRef AST
  // through `t`'s binding so SUM still sees a range.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LET(s, A1#, t, s, SUM(t))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

// ---------------------------------------------------------------------------
// `_xlfn.ANCHORARRAY(ref)` — the OOXML internal encoding of `ref#`.
// `strip_future_prefix` reduces the call name to bare `ANCHORARRAY` before
// dispatch, so these tests parse both spellings and expect the same result.
// ---------------------------------------------------------------------------

TEST(AnchorArrayLazy, RefOverCommittedSpill) {
  // Pre-commit a 3x1 spill at A1; `_xlfn.ANCHORARRAY(A1)` must return the
  // same Value::Array shape and cells as the postfix-`#` SpillRef branch.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells{Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=_xlfn.ANCHORARRAY(A1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* read_cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(read_cells[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(read_cells[1].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(read_cells[2].as_number(), 30.0);
}

TEST(AnchorArrayLazy, BareNameWithoutXlfnPrefix) {
  // Plain `ANCHORARRAY(A1)` should dispatch identically — the registry
  // entry is name-canonicalised, and bare callers (e.g. user-typed
  // formulas) must work the same as the OOXML-encoded `_xlfn.` form.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 1U, 1U, std::vector<Value>{Value::number(7.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=ANCHORARRAY(A1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 7.0);
}

TEST(AnchorArrayLazy, NoSpillReturnsRef) {
  // No spill anchored at A1 -> #REF!. Same surface as `=A1#` against an
  // un-spilled cell; the IronCalc spill_operator suite's E2 / E4 / E6
  // cases hit this exact path.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=_xlfn.ANCHORARRAY(A1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(AnchorArrayLazy, SumOverAnchorArray) {
  // SUM consumes the resulting Value::Array like any other range-shaped
  // arg, so `=SUM(_xlfn.ANCHORARRAY(A1))` over [10, 20, 30] is 60. This
  // is the same shape as `=SUM(A1#)`.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U,
                                 std::vector<Value>{Value::number(10.0), Value::number(20.0), Value::number(30.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SUM(_xlfn.ANCHORARRAY(A1))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(AnchorArrayLazy, SumOverSequenceCallFlattens) {
  // The eager dispatcher's generic Array-result flatten kicks in here:
  // SEQUENCE returns a Value::Array, which the fallback unpacks row-major
  // before calling SUM's impl. Mac Excel returns 6 for `=SUM(SEQUENCE(3))`
  // (1+2+3); the same fix that lets `SUM(_xlfn.ANCHORARRAY(A1))` work also
  // covers any other lazy builtin that returns an Array.
  Arena parse_arena;
  Arena eval_arena;
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  const Value v = EvalUnder("=SUM(SEQUENCE(3))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(AnchorArrayLazy, NestedOffsetCallResolvesAnchor) {
  // `_xlfn.ANCHORARRAY(OFFSET(A1, 1, 1))` should resolve OFFSET to its
  // top-left rectangle (B2) and look up the spill anchored there. This
  // covers the IronCalc `I2` case (`=SUM(_xlfn.ANCHORARRAY(OFFSET(A1,1,1)))`)
  // which the plain Ref shape does not exercise.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ASSERT_TRUE(sheet.commit_spill(
      1U, 1U, 4U, 1U,
      std::vector<Value>{Value::number(12.0), Value::number(24.0), Value::number(34.0), Value::number(1245.0)}));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SUM(_xlfn.ANCHORARRAY(OFFSET(A1, 1, 1)))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1315.0);
}

TEST(AnchorArrayLazy, WrongArityIsValueError) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value zero = EvalUnder("=_xlfn.ANCHORARRAY()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(zero.is_error());
  EXPECT_EQ(zero.as_error(), ErrorCode::Value);

  Arena parse_arena2;
  Arena eval_arena2;
  const Value two = EvalUnder("=_xlfn.ANCHORARRAY(A1, B2)", &parse_arena2, &eval_arena2, ctx);
  ASSERT_TRUE(two.is_error());
  EXPECT_EQ(two.as_error(), ErrorCode::Value);
}

TEST(AnchorArrayLazy, NonReferenceArgIsValueError) {
  // A literal arg has no anchor cell to look up; mirrors Excel's
  // `=_xlfn.ANCHORARRAY(1)` -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=_xlfn.ANCHORARRAY(1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
