// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for the dynamic-array (spilling) built-ins. Two layers:
//
//   1. Direct function-pointer tests: invoke the registered SEQUENCE impl
//      against constructed `Value` args. This isolates SEQUENCE's coercion
//      and shape rules from the parser / evaluator.
//   2. End-to-end spill-pipeline tests: parse `=SEQUENCE(...)`, evaluate
//      with `with_mutable_sheet` + `with_formula_cell`, then call
//      `dispatch_array_result` to commit the spill. These exercise the
//      pipeline established by Phases 1-4 (sheet spill table, EvalContext
//      dispatch, spill-aware reads, OOXML anchor / phantom suppression);
//      SEQUENCE is the first acceptance test for that machinery.

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "cell.h"
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

// Looks up SEQUENCE in the process-wide default registry. Asserts on miss
// so a registration regression surfaces as a named test failure instead of
// a null-deref later in each case.
const FunctionDef* SequenceDef() {
  const FunctionDef* def = default_registry().lookup("SEQUENCE");
  EXPECT_NE(def, nullptr) << "SEQUENCE must be registered in the default registry";
  return def;
}

// Invokes SEQUENCE's impl with a fixed-size args array and the given
// arity, returning the produced `Value`. The arena lives on the caller's
// stack so the returned ArrayValue stays readable for assertions.
Value CallSequence(Arena& arena, std::initializer_list<Value> args) {
  const FunctionDef* def = SequenceDef();
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  // Copy into a local array so the caller's std::initializer_list backing
  // storage doesn't outlive the call. `Value`'s default ctor is private; we
  // seed the buffer with `Value::blank()` then overwrite the live slots.
  Value buf[4] = {Value::blank(), Value::blank(), Value::blank(), Value::blank()};
  std::size_t i = 0;
  for (const Value& v : args) {
    if (i >= 4U) {
      return Value::error(ErrorCode::Value);
    }
    buf[i++] = v;
  }
  return def->impl(buf, static_cast<std::uint32_t>(args.size()), arena);
}

// ---------------------------------------------------------------------------
// Direct SEQUENCE function-pointer tests
// ---------------------------------------------------------------------------

TEST(BuiltinsDynamicArray, Sequence_OneArg_ReturnsColumnVector) {
  // SEQUENCE(3) is a 3x1 column vector [1, 2, 3] (rows-first convention,
  // confirmed against Mac Excel 365).
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(3.0)});
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0], Value::number(1.0));
  EXPECT_EQ(cells[1], Value::number(2.0));
  EXPECT_EQ(cells[2], Value::number(3.0));
}

TEST(BuiltinsDynamicArray, Sequence_TwoArgs_Returns2DGrid) {
  // SEQUENCE(2, 3) is 2x3 row-major: [1,2,3 / 4,5,6].
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(2.0), Value::number(3.0)});
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  for (std::size_t i = 0; i < 6U; ++i) {
    EXPECT_EQ(cells[i], Value::number(static_cast<double>(i + 1U))) << "i=" << i;
  }
}

TEST(BuiltinsDynamicArray, Sequence_FourArgs_RespectsStartAndStep) {
  // SEQUENCE(2, 2, 10, 5) -> [10, 15, 20, 25] in row-major order.
  Arena arena;
  const Value v =
      CallSequence(arena, {Value::number(2.0), Value::number(2.0), Value::number(10.0), Value::number(5.0)});
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0], Value::number(10.0));
  EXPECT_EQ(cells[1], Value::number(15.0));
  EXPECT_EQ(cells[2], Value::number(20.0));
  EXPECT_EQ(cells[3], Value::number(25.0));
}

TEST(BuiltinsDynamicArray, Sequence_NegativeStep) {
  // SEQUENCE(3, 1, 10, -2) -> [10, 8, 6]. Step is signed; the formula
  // `start + i*step` works for any sign without special casing.
  Arena arena;
  const Value v =
      CallSequence(arena, {Value::number(3.0), Value::number(1.0), Value::number(10.0), Value::number(-2.0)});
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0], Value::number(10.0));
  EXPECT_EQ(cells[1], Value::number(8.0));
  EXPECT_EQ(cells[2], Value::number(6.0));
}

TEST(BuiltinsDynamicArray, Sequence_FractionalStartStep) {
  // SEQUENCE(3, 1, 0.5, 0.25) -> [0.5, 0.75, 1.0]. Exact in IEEE-754
  // because every intermediate is a power-of-two-scaled small integer.
  Arena arena;
  const Value v =
      CallSequence(arena, {Value::number(3.0), Value::number(1.0), Value::number(0.5), Value::number(0.25)});
  ASSERT_TRUE(v.is_array());
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 0.5);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 0.75);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 1.0);
}

TEST(BuiltinsDynamicArray, Sequence_NonIntegerRowsTruncates) {
  // SEQUENCE(3.7) truncates rows toward zero -> 3 rows. Mac Excel does
  // the same; tested via direct impl call so the parser's number-literal
  // path is irrelevant.
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(3.7)});
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
}

TEST(BuiltinsDynamicArray, Sequence_ZeroRowsReturnsValue) {
  // SEQUENCE(0) -> #VALUE! per Mac Excel (rows must be a positive
  // integer; 0 fails the `> 0` guard).
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(0.0)});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDynamicArray, Sequence_NegativeRowsReturnsValue) {
  // SEQUENCE(-1) -> #VALUE! (truncates to -1, fails `> 0`).
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(-1.0)});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDynamicArray, Sequence_TooManyCellsReturnsNum) {
  // 2,000,000 rows * 1 col exceeds the per-call cap (~1M cells). Surface
  // #NUM! to prevent runaway arena allocation -- matches Excel's overflow
  // code for SEQUENCE.
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(2'000'000.0)});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDynamicArray, Sequence_ColsExceedingMaxReturnsNum) {
  // cols > Sheet::kMaxCols -> #NUM! before the rows*cols multiplication
  // even runs.
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(1.0), Value::number(50'000.0)});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDynamicArray, Sequence_PropagatesErrorArg) {
  // A pre-evaluated error in any arg propagates through the impl's own
  // `coerce_to_number` (the dispatcher would short-circuit it earlier in
  // production; here we hit the impl directly to exercise the in-impl
  // guard).
  Arena arena;
  const Value v = CallSequence(arena, {Value::error(ErrorCode::Div0)});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsDynamicArray, Sequence_TextArgReturnsValue) {
  // A non-numeric text arg fails coercion -> #VALUE!.
  Arena arena;
  const Value v = CallSequence(arena, {Value::text("abc")});
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsDynamicArray, Sequence_DefaultColsIsOne) {
  // Three-arg form (rows, cols, start) with cols=1 still defaults step to 1.
  Arena arena;
  const Value v = CallSequence(arena, {Value::number(2.0), Value::number(1.0), Value::number(7.0)});
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0], Value::number(7.0));
  EXPECT_EQ(cells[1], Value::number(8.0));
}

// ---------------------------------------------------------------------------
// End-to-end spill pipeline tests (parse -> evaluate -> dispatch)
// ---------------------------------------------------------------------------
//
// These tests verify that SEQUENCE flows through the full producer -> spill
// pipeline established by Phases 1-4: the function returns a Value::Array,
// `EvalContext::dispatch_array_result` commits a spill region anchored at
// the formula cell, and downstream reads via `Sheet::resolve_cell_value`
// observe the spilled phantoms. This is the canonical acceptance test for
// the cell-level dynamic-array machinery.

// Convenience: parse `src`, run `evaluate`, then dispatch any array result
// against the supplied context. Returns the post-dispatch scalar value that
// a downstream consumer (the recalc layer, eventually) would propagate.
struct EvalResult {
  Value value;  // Post-dispatch value at the formula cell.
};

EvalResult EvalAndDispatch(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return {Value::error(ErrorCode::Name)};
  }
  const Value raw = evaluate(*root, *eval_arena, default_registry(), ctx);
  return {ctx.dispatch_array_result(raw)};
}

TEST(SpillPipeline, SequenceCommitsThreeByOne) {
  // `=SEQUENCE(3)` typed at A1 spills to A1:A3. The dispatcher returns
  // the anchor scalar (1); the spill region carries the full [1, 2, 3]
  // vector and phantom reads through `resolve_cell_value` find the spilled
  // values at A2 / A3.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);

  Arena parse_arena;
  Arena eval_arena;
  const EvalResult r = EvalAndDispatch("=SEQUENCE(3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(r.value.is_number());
  EXPECT_DOUBLE_EQ(r.value.as_number(), 1.0);

  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 1U);
  ASSERT_EQ(region->cells.size(), 3U);
  EXPECT_EQ(region->cells[0], Value::number(1.0));
  EXPECT_EQ(region->cells[1], Value::number(2.0));
  EXPECT_EQ(region->cells[2], Value::number(3.0));

  // Phantom reads via the spill-aware accessor.
  EXPECT_EQ(sheet.resolve_cell_value(1U, 0U), Value::number(2.0));
  EXPECT_EQ(sheet.resolve_cell_value(2U, 0U), Value::number(3.0));
}

TEST(SpillPipeline, SequenceCommitsTwoByThreeGrid) {
  // `=SEQUENCE(2, 3)` at A1 spills to a 2x3 rectangle A1:C2 with the
  // row-major fill 1..6. Verifies the dispatcher handles non-degenerate
  // 2D shapes, not just column vectors.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);

  Arena parse_arena;
  Arena eval_arena;
  const EvalResult r = EvalAndDispatch("=SEQUENCE(2,3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(r.value.is_number());
  EXPECT_DOUBLE_EQ(r.value.as_number(), 1.0);

  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 2U);
  EXPECT_EQ(region->cols, 3U);
  ASSERT_EQ(region->cells.size(), 6U);
  // Phantoms at (0,1), (0,2), (1,0), (1,1), (1,2).
  EXPECT_EQ(sheet.resolve_cell_value(0U, 1U), Value::number(2.0));
  EXPECT_EQ(sheet.resolve_cell_value(0U, 2U), Value::number(3.0));
  EXPECT_EQ(sheet.resolve_cell_value(1U, 0U), Value::number(4.0));
  EXPECT_EQ(sheet.resolve_cell_value(1U, 1U), Value::number(5.0));
  EXPECT_EQ(sheet.resolve_cell_value(1U, 2U), Value::number(6.0));
}

TEST(SpillPipeline, SequenceCollisionEmitsSpillError) {
  // Pre-populate A2 (in the proposed spill footprint) with a literal.
  // Dispatch must return #SPILL!, leave the literal intact, and register
  // no spill region.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(1U, 0U, Value::number(99.0));
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);

  Arena parse_arena;
  Arena eval_arena;
  const EvalResult r = EvalAndDispatch("=SEQUENCE(3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(r.value.is_error());
  EXPECT_EQ(r.value.as_error(), ErrorCode::Spill);

  // No spill committed; the pre-existing literal at A2 is intact.
  EXPECT_EQ(sheet.spill_region_at_anchor(0U, 0U), nullptr);
  const Cell* literal = sheet.cell_at(1U, 0U);
  ASSERT_NE(literal, nullptr);
  ASSERT_TRUE(literal->cached_value.is_number());
  EXPECT_DOUBLE_EQ(literal->cached_value.as_number(), 99.0);
}

TEST(SpillPipeline, PostSpillCellReadObservesPhantom) {
  // After SEQUENCE(3) spills to A1:A3, evaluating `=A2` from a separate
  // formula cell observes the phantom value at A2 (=2). This exercises
  // Phase 3's spill-aware read switch wired into `EvalContext::resolve_ref`.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);

  // Step 1: commit the spill at A1 by evaluating the producer formula.
  {
    EvalState state;
    const EvalContext base_ctx(wb, sheet, state);
    const EvalContext ctx = base_ctx.with_mutable_sheet(sheet).with_formula_cell(0U, 0U);
    Arena parse_arena;
    Arena eval_arena;
    const EvalResult r = EvalAndDispatch("=SEQUENCE(3)", &parse_arena, &eval_arena, ctx);
    ASSERT_TRUE(r.value.is_number());
    ASSERT_NE(sheet.spill_region_at_anchor(0U, 0U), nullptr);
  }

  // Step 2: evaluate `=A2` from C1 (anchor at row 0, col 2). The reference
  // resolver must observe the phantom at A2 = 2, NOT the underlying blank.
  EvalState state;
  const EvalContext base_ctx(wb, sheet, state);
  const EvalContext c1_ctx = base_ctx.with_formula_cell(0U, 2U);

  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=A2", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value out = evaluate(*root, eval_arena, default_registry(), c1_ctx);
  ASSERT_TRUE(out.is_number());
  EXPECT_DOUBLE_EQ(out.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// TRANSPOSE tests (lazy-dispatched, defined in shape_ops_lazy.cpp)
// ---------------------------------------------------------------------------
//
// TRANSPOSE conceptually belongs to the dynamic-array family but rides the
// lazy-dispatch seam in `tree_walker.cpp` because it must inspect the
// per-argument AST shape (the eager dispatcher would flatten a Range arg
// to a row-major scalar vector and TRANSPOSE would lose the 2D shape).
// The end-to-end pipeline still treats the produced ArrayValue as a
// dynamic-array spill candidate when committed via `dispatch_array_result`.

// Convenience: parse `src`, evaluate, return the resulting Value with no
// dispatch. The returned ArrayValue lives in `eval_arena`.
Value EvalNoDispatch(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  EXPECT_TRUE(p.errors().empty()) << "unexpected errors for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, *eval_arena, default_registry(), ctx);
}

TEST(BuiltinsTranspose, Transpose_2x3ArrayLiteral) {
  // `=TRANSPOSE({1,2,3;4,5,6})` -> 3x2 with cells [1,4,2,5,3,6] in
  // row-major order (i.e. column-major over the original literal).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE({1,2,3;4,5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0], Value::number(1.0));
  EXPECT_EQ(cells[1], Value::number(4.0));
  EXPECT_EQ(cells[2], Value::number(2.0));
  EXPECT_EQ(cells[3], Value::number(5.0));
  EXPECT_EQ(cells[4], Value::number(3.0));
  EXPECT_EQ(cells[5], Value::number(6.0));
}

TEST(BuiltinsTranspose, Transpose_RangeArg) {
  // `=TRANSPOSE(A1:B2)` over A1:B2 = {{1, 2}; {3, 4}} -> {{1, 3}; {2, 4}}.
  // This exercises the lazy-dispatch seam: a Range argument retains its
  // 2D shape via `eval_node_as_array` rather than collapsing to a vector.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  sheet.set_cell_value(0U, 1U, Value::number(2.0));
  sheet.set_cell_value(1U, 0U, Value::number(3.0));
  sheet.set_cell_value(1U, 1U, Value::number(4.0));
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE(A1:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0], Value::number(1.0));
  EXPECT_EQ(cells[1], Value::number(3.0));
  EXPECT_EQ(cells[2], Value::number(2.0));
  EXPECT_EQ(cells[3], Value::number(4.0));
}

TEST(BuiltinsTranspose, Transpose_Scalar) {
  // `=TRANSPOSE(42)` -> 1x1 array with the scalar wrapped. Mac Excel
  // returns the same scalar; the engine wraps it in a degenerate 1x1
  // ArrayValue so the spill pipeline observes a uniform Array shape.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE(42)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_EQ(v.as_array_cells()[0], Value::number(42.0));
}

TEST(BuiltinsTranspose, Transpose_PreservesErrors) {
  // `=TRANSPOSE({1,#N/A;2,3})` -> 2x2 with the error cell pass-through.
  // TRANSPOSE does not propagate errors at the function level — it just
  // moves cells; `#N/A` lands at its transposed position intact.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE({1,#N/A;2,3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  // Expected layout (row-major): [1, 2, #N/A, 3].
  EXPECT_EQ(cells[0], Value::number(1.0));
  EXPECT_EQ(cells[1], Value::number(2.0));
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::NA);
  EXPECT_EQ(cells[3], Value::number(3.0));
}

TEST(BuiltinsTranspose, Transpose_Twice) {
  // TRANSPOSE is its own inverse: `=TRANSPOSE(TRANSPOSE({1,2,3;4,5,6}))`
  // returns the original 2x3 array.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE(TRANSPOSE({1,2,3;4,5,6}))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  for (std::size_t i = 0; i < 6U; ++i) {
    EXPECT_EQ(cells[i], Value::number(static_cast<double>(i + 1U))) << "i=" << i;
  }
}

TEST(BuiltinsTranspose, Transpose_PropagatesArgError) {
  // `=TRANSPOSE(1/0)`: the argument evaluates to `#DIV/0!`. TRANSPOSE
  // surfaces the scalar error verbatim (no array wrapper) so callers see
  // the canonical short-circuit shape.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE(1/0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsTranspose, Transpose_TextValuesPassThrough) {
  // `=TRANSPOSE({"a","b";"c","d"})` -> 2x2 with cells {"a","c","b","d"}.
  // Text payloads survive the transposition unchanged (the cells are
  // shallow-copied; the string_view storage is the literal's arena which
  // outlives the eval arena for the duration of the assertion).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalNoDispatch("=TRANSPOSE({\"a\",\"b\";\"c\",\"d\"})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_text());
  EXPECT_EQ(cells[0].as_text(), std::string_view("a"));
  ASSERT_TRUE(cells[1].is_text());
  EXPECT_EQ(cells[1].as_text(), std::string_view("c"));
  ASSERT_TRUE(cells[2].is_text());
  EXPECT_EQ(cells[2].as_text(), std::string_view("b"));
  ASSERT_TRUE(cells[3].is_text());
  EXPECT_EQ(cells[3].as_text(), std::string_view("d"));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
