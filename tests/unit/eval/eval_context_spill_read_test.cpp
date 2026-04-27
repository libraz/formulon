// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the phantom-aware read path in `EvalContext::resolve_ref`.
//
// The previous-phase change made a formula cell that returns `Value::Array`
// commit a spill region on the bound mutable sheet. This phase routes
// cross-cell reads of phantom coordinates through `Sheet::resolve_cell_value`,
// so a separate cell holding `=A2` (where A2 is a phantom of A1's spill)
// observes the spilled value rather than `Value::blank()`.
//
// These tests pin:
//
//   * Phantom coordinates surface the spilled value via the 3-arg
//     `resolve_ref` overload.
//   * Clearing a spill reverts phantom coordinates to `blank()`.
//   * On a `#SPILL!` collision (no spill registered), pre-existing literals
//     still resolve verbatim and the anchor surfaces the error.
//   * The integration path: a formula cell at C1 evaluating `=A2` observes
//     20 after the spill at A1 has been committed by the recursive resolver.
//   * The 1-arg `resolve_ref` overload also benefits from the change,
//     proving the fix sits in the shared `resolve_prefix` helper rather
//     than being localised to one of the two call sites.

#include <cstdint>

#include "cell.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Builds a reference to `(row, col)` on the implicit current sheet (no
// sheet qualifier).
parser::Reference MakeLocalRef(std::uint32_t row, std::uint32_t col) {
  parser::Reference ref{};
  ref.row = row;
  ref.col = col;
  return ref;
}

// ---------------------------------------------------------------------------
// Direct phantom reads through resolve_ref
// ---------------------------------------------------------------------------

TEST(EvalContextSpillRead, PhantomCellReadReturnsSpilledValue) {
  // Sheet has a 3x1 spill at A1 committed directly through the Sheet API
  // — no evaluator involvement, so the test isolates the read-path change.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena arena;
  const FunctionRegistry registry;

  // Anchor: stored Cell with cached_value = cells[0]; the formula branch is
  // not taken because the anchor's formula_text is empty (commit_spill alone
  // does not set a formula). The new read path therefore goes through
  // resolve_cell_value, which returns the same number.
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(0U, 0U), arena, registry), Value::number(10.0));
  // Phantom rows: previously returned blank, now surface the spilled values.
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(1U, 0U), arena, registry), Value::number(20.0));
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(2U, 0U), arena, registry), Value::number(30.0));
}

// ---------------------------------------------------------------------------
// Spill clearing reverts phantom reads to blank
// ---------------------------------------------------------------------------

TEST(EvalContextSpillRead, PhantomReadRevertsAfterClear) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena arena;
  const FunctionRegistry registry;

  // Sanity: phantom is visible before clear.
  ASSERT_EQ(ctx.resolve_ref(MakeLocalRef(1U, 0U), arena, registry), Value::number(20.0));

  sheet.clear_spill(0U, 0U);

  // After clear: no spill region, no stored cell at A2 → blank.
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(1U, 0U), arena, registry), Value::blank());
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(2U, 0U), arena, registry), Value::blank());
}

// ---------------------------------------------------------------------------
// Collision case: anchor surfaces #SPILL!, phantom coords return literal
// ---------------------------------------------------------------------------

TEST(EvalContextSpillRead, PhantomReadOnSpillCollision) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Pre-populate A2 with a literal that conflicts with the proposed
  // 3x1 spill footprint anchored at A1.
  sheet.set_cell_value(1U, 0U, Value::number(99.0));

  // commit_spill must fail and set the anchor's cached_value to #SPILL!.
  std::vector<Value> cells = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_FALSE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));
  ASSERT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena arena;
  const FunctionRegistry registry;

  // Anchor: cached_value is #SPILL!, formula_text is empty → goes through
  // the new resolve_cell_value path, which falls back to the stored
  // cached_value because no spill region exists at the anchor.
  const Value anchor_v = ctx.resolve_ref(MakeLocalRef(0U, 0U), arena, registry);
  ASSERT_TRUE(anchor_v.is_error());
  EXPECT_EQ(anchor_v.as_error(), ErrorCode::Spill);

  // A2: pre-existing literal must round-trip unchanged. No spill was
  // registered, so resolve_cell_value returns the stored cached_value.
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(1U, 0U), arena, registry), Value::number(99.0));
  // A3: never written, no spill covers it → blank.
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(2U, 0U), arena, registry), Value::blank());
}

// ---------------------------------------------------------------------------
// Integration: cross-cell formula `=A2` sees the phantom value
// ---------------------------------------------------------------------------

TEST(EvalContextSpillRead, CrossCellFormulaSeesPhantom) {
  // Commit a 3x1 spill at A1 directly (no upstream formula needed — this
  // test isolates the cross-cell read path). Then evaluate C1's formula
  // `=A2`. The reference resolves the phantom A2 of A1's spill and must
  // observe 20 through the new `resolve_cell_value` read.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  // No builtins are needed — the formula `=A2` only walks the reference
  // resolver — but `evaluate()` requires a registry, so an empty one suffices.
  const FunctionRegistry registry;
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  // C1 is the formula-cell anchor; the mutable sheet binding is unnecessary
  // for the phantom read but kept for parity with the recursive resolver's
  // standard context shape.
  const EvalContext c1_ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 2U);

  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=A2", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value c1_value = evaluate(*root, eval_arena, registry, c1_ctx);
  ASSERT_TRUE(c1_value.is_number());
  EXPECT_DOUBLE_EQ(c1_value.as_number(), 20.0);
}

// ---------------------------------------------------------------------------
// 1-arg overload also picks up the change
// ---------------------------------------------------------------------------

TEST(EvalContextSpillRead, OneArgResolveRefSeesPhantom) {
  // Proves the fix lives in the shared `resolve_prefix` helper: the 1-arg
  // overload (no arena, no registry) must observe phantom values too,
  // since both overloads route through the same preamble.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(0U, 0U)), Value::number(10.0));
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(1U, 0U)), Value::number(20.0));
  EXPECT_EQ(ctx.resolve_ref(MakeLocalRef(2U, 0U)), Value::number(30.0));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
