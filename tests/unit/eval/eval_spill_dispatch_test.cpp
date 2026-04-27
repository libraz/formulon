// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `EvalContext::dispatch_array_result` and the recursive-
// resolver wiring that calls it. These tests pin the cell-level dynamic-
// array spill dispatch contract:
//
//   * Scalar passthrough: non-array values flow through unchanged.
//   * Opt-in safety: without a bound `mutable_sheet()` (read-only context)
//     or without a formula-cell anchor (ad-hoc evaluation), an Array result
//     is returned verbatim and no spill is committed.
//   * Anchor commit: a top-level Array dispatched against a bound
//     mutable sheet + formula cell registers a spill region and returns
//     the anchor scalar (cells[0]).
//   * Collision: a footprint that overlaps a pre-existing literal yields
//     `#SPILL!` and leaves no spill registered.
//   * Recursive resolver integration: `=A1` (where A1's formula returns
//     an array) triggers the dispatch on the recursed-into A1 anchor,
//     spills there, and propagates the anchor scalar to the calling cell.
//   * Memoisation: the recursive path stores the post-dispatch scalar,
//     not the raw array, on `EvalState`.
//   * Cross-sheet recursive: dispatch is gated on `target_sheet ==
//     mutable_sheet`, so an array returned by a formula on a different
//     sheet does NOT mutate that sheet's spill table.

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "cell.h"
#include "eval/builtins.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/shape_ops_lazy.h"
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

// Test-only function impl that returns a 3x1 column array of [10, 20, 30].
// The ArrayValue and its cells are allocated in the supplied arena, so
// callers must keep `arena` alive for the lifetime of the returned Value.
// Bypassing the bare-ArrayLiteral `#VALUE!` path in `eval_node` is the
// only way to produce a top-level `Value::Array` from the tree walker; a
// future SEQUENCE / TOROW builtin will offer the same hook in production.
Value TestArrayImpl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& arena) {
  Value* cells = arena.create_array<Value>(3U);
  cells[0] = Value::number(10.0);
  cells[1] = Value::number(20.0);
  cells[2] = Value::number(30.0);
  ArrayValue* arr = arena.create<ArrayValue>();
  arr->rows = 3U;
  arr->cols = 1U;
  arr->cells = cells;
  return Value::array(arr);
}

// Builds a registry that contains every Formulon builtin plus a single
// test-only `TEST_ARRAY()` function (zero-arity) that yields a Value::Array.
// Returned by value because `FunctionRegistry` is move-constructible.
FunctionRegistry BuildRegistryWithTestArray() {
  FunctionRegistry r;
  register_builtins(r);
  r.register_function(FunctionDef{"TEST_ARRAY", 0U, 0U, &TestArrayImpl});
  return r;
}

// Bundles per-test parser + evaluator arenas. The arenas must outlive any
// `Value::Array` returned from the evaluator (cells live in `eval_arena`).
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

// Produces a `Value::Array` from an array-literal source by evaluating it
// in array context. Returning Array directly via top-level `evaluate()`
// would surface `#VALUE!` because `eval_node` rejects bare ArrayLiterals;
// `eval_node_as_array` is the documented producer that wraps them.
Value MakeArrayFromLiteral(EvalHarness* h, std::string_view src, const EvalContext& ctx) {
  EXPECT_TRUE(h->parse(src)) << "parse failed for: " << src;
  if (h->root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return eval_node_as_array(*h->root, h->eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// dispatch_array_result: anchor-only (top-level) path
// ---------------------------------------------------------------------------

TEST(EvalSpillDispatch, ColumnArrayCommitsSpillAndReturnsAnchorScalar) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);

  EvalHarness h;
  const Value array_v = MakeArrayFromLiteral(&h, "={1;2;3}", ctx);
  ASSERT_TRUE(array_v.is_array());

  const Value out = ctx.dispatch_array_result(array_v);
  ASSERT_TRUE(out.is_number());
  EXPECT_DOUBLE_EQ(out.as_number(), 1.0);

  // Spill region exists at the anchor with the expected shape and cells.
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 1U);
  ASSERT_EQ(region->cells.size(), 3U);
  EXPECT_EQ(region->cells[0], Value::number(1.0));
  EXPECT_EQ(region->cells[1], Value::number(2.0));
  EXPECT_EQ(region->cells[2], Value::number(3.0));

  // Phantoms are spill-aware via resolve_cell_value.
  EXPECT_EQ(sheet.resolve_cell_value(1U, 0U), Value::number(2.0));
  EXPECT_EQ(sheet.resolve_cell_value(2U, 0U), Value::number(3.0));
}

// ---------------------------------------------------------------------------
// dispatch_array_result: opt-in safety (no mutable_sheet / no anchor)
// ---------------------------------------------------------------------------

TEST(EvalSpillDispatch, NoMutableSheetIsPassthrough) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  // Anchor only - no with_mutable_sheet.
  const EvalContext ctx = base_ctx.with_formula_cell(0U, 0U);

  EvalHarness h;
  const Value array_v = MakeArrayFromLiteral(&h, "={1;2;3}", ctx);
  ASSERT_TRUE(array_v.is_array());

  const Value out = ctx.dispatch_array_result(array_v);
  ASSERT_TRUE(out.is_array());
  EXPECT_EQ(out.as_array(), array_v.as_array());
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);
}

TEST(EvalSpillDispatch, NoFormulaCellAnchorIsPassthrough) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  // Mutable sheet only - no with_formula_cell.
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet);

  EvalHarness h;
  const Value array_v = MakeArrayFromLiteral(&h, "={1;2;3}", ctx);
  ASSERT_TRUE(array_v.is_array());

  const Value out = ctx.dispatch_array_result(array_v);
  ASSERT_TRUE(out.is_array());
  EXPECT_EQ(out.as_array(), array_v.as_array());
  // No spill anywhere.
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);
}

// ---------------------------------------------------------------------------
// dispatch_array_result: scalar passthrough
// ---------------------------------------------------------------------------

TEST(EvalSpillDispatch, ScalarValuesArePassthroughEvenWhenFullyBound) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);

  // A range of scalar variants exercises the kind-check; each must round-trip.
  EXPECT_EQ(ctx.dispatch_array_result(Value::number(42.0)), Value::number(42.0));
  EXPECT_EQ(ctx.dispatch_array_result(Value::boolean(true)), Value::boolean(true));
  EXPECT_EQ(ctx.dispatch_array_result(Value::blank()), Value::blank());
  EXPECT_EQ(ctx.dispatch_array_result(Value::error(ErrorCode::Div0)), Value::error(ErrorCode::Div0));
  // No spill registered as a side effect.
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);
}

// ---------------------------------------------------------------------------
// dispatch_array_result: collision
// ---------------------------------------------------------------------------

TEST(EvalSpillDispatch, CollisionWithExistingLiteralReturnsSpillError) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Pre-populate A2 (row 1, col 0) with a literal - this is in the
  // proposed footprint of the 3x1 spill anchored at A1.
  sheet.set_cell_value(1U, 0U, Value::number(99.0));

  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);

  EvalHarness h;
  const Value array_v = MakeArrayFromLiteral(&h, "={1;2;3}", ctx);
  ASSERT_TRUE(array_v.is_array());

  const Value out = ctx.dispatch_array_result(array_v);
  ASSERT_TRUE(out.is_error());
  EXPECT_EQ(out.as_error(), ErrorCode::Spill);

  // No spill region was registered; the pre-existing literal is intact.
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);
  const Cell* literal = sheet.cell_at(1U, 0U);
  ASSERT_NE(literal, nullptr);
  ASSERT_TRUE(literal->cached_value.is_number());
  EXPECT_DOUBLE_EQ(literal->cached_value.as_number(), 99.0);
}

// ---------------------------------------------------------------------------
// Recursive resolver integration: `=A1` where A1's formula returns an array
// ---------------------------------------------------------------------------

TEST(EvalSpillDispatch, RecursiveResolverDispatchesOnReferencedFormulaCell) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // A1 is a formula whose top-level evaluation produces a 3x1 array via
  // the test-only TEST_ARRAY() builtin. The recursive resolver evaluates
  // A1 with the formula-cell anchor at A1 and the mutable sheet still
  // bound; the in-flight dispatch fires and commits the spill at A1,
  // memoising the anchor scalar (10).
  sheet.set_cell_formula(0U, 0U, "=TEST_ARRAY()");

  // Evaluate C1's formula `=A1` against a context where C1 is the anchor
  // and `sheet` is the mutable target. C1 is far from A1's spill area
  // (cols A:A vs. col C), so no collision is possible from C1's side.
  const FunctionRegistry registry = BuildRegistryWithTestArray();
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext c1_ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 2U);

  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=A1", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value c1_value = evaluate(*root, eval_arena, registry, c1_ctx);

  // C1 should observe A1's anchor scalar.
  ASSERT_TRUE(c1_value.is_number());
  EXPECT_DOUBLE_EQ(c1_value.as_number(), 10.0);

  // The spill region was committed at A1.
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 1U);

  // Phantoms read back the spilled values.
  EXPECT_EQ(sheet.resolve_cell_value(1U, 0U), Value::number(20.0));
  EXPECT_EQ(sheet.resolve_cell_value(2U, 0U), Value::number(30.0));

  // EvalState memoisation stores the post-dispatch scalar, not the array.
  // Without this property, a second `=A1` reference inside the same
  // evaluate() call would observe an Array rather than the spilled anchor.
  const Value* memo = state.lookup_memo(&sheet, 0U, 0U);
  ASSERT_NE(memo, nullptr);
  ASSERT_TRUE(memo->is_number());
  EXPECT_DOUBLE_EQ(memo->as_number(), 10.0);
}

// ---------------------------------------------------------------------------
// Cross-sheet recursive: dispatch is sheet-gated
// ---------------------------------------------------------------------------

TEST(EvalSpillDispatch, CrossSheetRecursiveDoesNotCommitOnOtherSheet) {
  // Sheet1 is the mutable target; Sheet2 hosts an array-producing formula
  // at A1. A formula on Sheet1 referencing `Sheet2!A1` recurses into
  // Sheet2, but `dispatch_array_result` must NOT fire there because
  // Sheet2 != mutable_sheet. The raw Array therefore propagates back,
  // and Sheet2's spill table remains empty.
  //
  // `add_sheet` may reallocate the `sheets_` vector and invalidate any
  // earlier `wb.sheet(...)` references, so all sheet bindings are taken
  // AFTER the workbook layout is final.
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  Sheet& sheet1 = wb.sheet(0);
  Sheet& sheet2 = wb.sheet(1);
  sheet2.set_cell_formula(0U, 0U, "=TEST_ARRAY()");

  const FunctionRegistry registry = BuildRegistryWithTestArray();
  EvalState state;
  const EvalContext base_ctx(wb, sheet1, state);
  // B1 on Sheet1 owns the formula `=Sheet2!A1`; mutable sheet is Sheet1.
  const EvalContext b1_ctx = base_ctx.with_mutable_sheet(sheet1).with_formula_cell(0U, 1U);

  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=Sheet2!A1", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value b1_value = evaluate(*root, eval_arena, registry, b1_ctx);

  // No spill committed on Sheet2 - the cross-sheet gate held.
  EXPECT_EQ(sheet2.spill_region_at_anchor(0U, 0U), nullptr);
  // No spill committed on Sheet1 either: B1's outer formula evaluator does
  // not invoke dispatch_array_result on its own (that's the eventual
  // top-level recalc layer's job, not Phase 2's). The raw Array surfaces
  // verbatim here, which is the documented Phase 2 behaviour for
  // cross-sheet array references.
  EXPECT_EQ(sheet1.spill_region_at_anchor(0U, 1U), nullptr);
  EXPECT_TRUE(b1_value.is_array() || b1_value.is_error());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
