//
// Unit tests for Excel 365 dynamic-array spill semantics on bare ranges
// (`=A1:A5`) and implicit intersection on `@`-prefixed ranges
// (`=@A1:A5`). Verified Mac semantics:
// `tests/oracle/cases/implicit_intersection.yaml` and corresponding
// `tests/oracle/golden/implicit_intersection.golden.json`.
//
// Bare range in a value context: spills the whole rectangle as a
// dynamic array, independent of the formula cell's position. Blank source
// cells surface as 0 per Excel's grid contract. A degenerate `A1:A1`
// range collapses to the scalar cell.
// `@`-prefixed range: projects the range onto the formula cell's row or
// column. Single-column range requires the formula row to fall inside
// the range; single-row range requires the formula column to fall
// inside the range; 2D ranges return `#VALUE!`.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
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

// Evaluate `src` against a bound workbook + current sheet without an
// anchored formula cell. Used for the bare-range spill-anchor cases.
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

// Evaluate `src` against a bound workbook anchored at the formula cell
// (`row`, `col`) on the current sheet. Used for the `@`-prefixed
// implicit-intersection cases.
Value EvalSourceAt(std::string_view src, const Workbook& wb, const Sheet& current, std::uint32_t row,
                   std::uint32_t col) {
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
  const EvalContext ctx = test::workbook_context(wb, current, state).with_formula_cell(row, col);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Assert that `v` is a spilled array of the given shape whose cells equal
// `expected` (row-major). Blank source cells surface as 0 per Excel's grid
// contract, so `expected` carries the already-promoted numbers.
void ExpectNumberArray(const Value& v, std::uint32_t rows, std::uint32_t cols, const std::vector<double>& expected) {
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), rows);
  EXPECT_EQ(v.as_array_cols(), cols);
  ASSERT_EQ(expected.size(), static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols));
  const Value* cells = v.as_array_cells();
  for (std::size_t i = 0; i < expected.size(); ++i) {
    ASSERT_TRUE(cells[i].is_number()) << "cell " << i << " is not a number";
    EXPECT_EQ(cells[i].as_number(), expected[i]) << "cell " << i;
  }
}

// ---------------------------------------------------------------------------
// Bare range -> dynamic-array spill of the whole rectangle.
// ---------------------------------------------------------------------------

TEST(RangeOp, SpillsSingleColumn) {
  // =A1:A5 spills the whole column; the two unset cells surface as 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(99.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(100.0));
  const Value v = EvalSourceIn("=A1:A5", wb, wb.sheet(0));
  ExpectNumberArray(v, 5U, 1U, {42.0, 99.0, 100.0, 0.0, 0.0});
}

TEST(RangeOp, SpillsSingleRow) {
  // =A1:E1 spills the whole row; unset C1/D1 surface as 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(5.0));
  const Value v = EvalSourceIn("=A1:E1", wb, wb.sheet(0));
  ExpectNumberArray(v, 1U, 5U, {1.0, 2.0, 0.0, 0.0, 5.0});
}

TEST(RangeOp, Spills2D) {
  // =A1:B5 spills the 5x2 rectangle row-major; the unset tail is 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));
  const Value v = EvalSourceIn("=A1:B5", wb, wb.sheet(0));
  ExpectNumberArray(v, 5U, 2U, {11.0, 12.0, 21.0, 22.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0});
}

TEST(RangeOp, SpillsReverseOrderNormalised) {
  // =A5:A1 -- endpoints swapped. Normalised to A1:A5 then spilled.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(99.0));
  const Value v = EvalSourceIn("=A5:A1", wb, wb.sheet(0));
  ExpectNumberArray(v, 5U, 1U, {42.0, 0.0, 0.0, 0.0, 99.0});
}

TEST(RangeOp, SingleCellDegenerate) {
  // =A1:A1 -- degenerate single-cell range.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=A1:A1", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

// ---------------------------------------------------------------------------
// `@`-prefixed range -> implicit intersection projected onto the formula
// cell's row or column.
// ---------------------------------------------------------------------------

TEST(ImplicitIntersection, AlignedRowProjectsToFormulaRow) {
  // Formula at Z3 with =@A1:A5 -- single-column range; row 3 is in
  // [1..5], so the projection is A3.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));  // A3 (row=2, col=0)
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(5.0));
  // Formula cell at Z3 -> row=2, col=25 (Z is the 26th column, 0-based 25).
  const Value v = EvalSourceAt("=@A1:A5", wb, wb.sheet(0), 2U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(ImplicitIntersection, UnalignedReturnsValue) {
  // Formula at Z10 with =@A1:A5 -- formula row 10 is NOT in [1..5].
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(5.0));
  // Z10 -> row=9, col=25.
  const Value v = EvalSourceAt("=@A1:A5", wb, wb.sheet(0), 9U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ImplicitIntersection, SingleRowAlignedColumn) {
  // Formula at C5 with =@A1:E1 -- single-row range; column C (col=2) is
  // in [A..E] (cols 0..4), so the projection is C1 (row=0, col=2).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(30.0));  // C1
  wb.sheet(0).set_cell_value(0, 3, Value::number(40.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(50.0));
  // C5 -> row=4, col=2.
  const Value v = EvalSourceAt("=@A1:E1", wb, wb.sheet(0), 4U, 2U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 30.0);
}

TEST(ImplicitIntersection, TwoDColumnOutsideReturnsValue) {
  // 2D range: the formula column falls outside the range's column span, so
  // the row/column intersection is empty -> #VALUE!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));
  // Z3 -> row=2, col=25; range spans rows 0..4 cols 0..1: column 25 is
  // outside the column span.
  const Value v = EvalSourceAt("=@A1:B5", wb, wb.sheet(0), 2U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ImplicitIntersection, TwoDAlignedReturnsIntersectionCell) {
  // 2D range where the formula cell falls inside both the row and column
  // spans: the result is the single cell at (formula_row, formula_col).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));  // B2 (row=1, col=1)
  // Formula at B2 -> row=1, col=1; inside rows 0..4 and cols 0..1 -> B2.
  const Value v = EvalSourceAt("=@A1:B5", wb, wb.sheet(0), 1U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 22.0);
}

TEST(ImplicitIntersection, AtBindsTighterThanMultiply) {
  // `=@D3:D5*2` binds `@` to the range first, then multiplies the
  // intersected scalar. Formula at D4 (row 3) -> D4=20 -> 40.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 3, Value::number(10.0));  // D3
  wb.sheet(0).set_cell_value(3, 3, Value::number(20.0));  // D4
  wb.sheet(0).set_cell_value(4, 3, Value::number(30.0));  // D5
  const Value in_row = EvalSourceAt("=@D3:D5*2", wb, wb.sheet(0), 3U, 3U);
  ASSERT_TRUE(in_row.is_number());
  EXPECT_EQ(in_row.as_number(), 40.0);
  // Formula outside the range's rows -> #VALUE! (propagates through `*`).
  const Value out_row = EvalSourceAt("=@D3:D5*2", wb, wb.sheet(0), 0U, 3U);
  ASSERT_TRUE(out_row.is_error());
  EXPECT_EQ(out_row.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// `@` on a whole-axis reference. `A:A` / `1:1` parse as a single `Ref` and
// `A:C` / `1:3` as a `RangeOp` over two of them; both declare a rectangle
// spanning a grid axis and project exactly like a bounded one.
// ---------------------------------------------------------------------------

TEST(ImplicitIntersection, WholeColumnSingleRefProjectsFormulaRow) {
  // `=@C:C` at Z4 reads C4. Every formula row is inside a whole column, so
  // this form never runs out of span.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(3, 2, Value::number(34.0));  // C4
  const Value v = EvalSourceAt("=@C:C", wb, wb.sheet(0), 3U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 34.0);
}

TEST(ImplicitIntersection, WholeRowSingleRefProjectsFormulaColumn) {
  // `=@2:2` at D9 reads D2.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 3, Value::number(42.0));  // D2
  const Value v = EvalSourceAt("=@2:2", wb, wb.sheet(0), 8U, 3U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(ImplicitIntersection, WholeColumnSpanIntersectsOnBothAxes) {
  // `=@A:C` declares columns A..C over the whole grid height, so it is a 2-D
  // rectangle: inside the column span it reads (formula row, formula col).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(4, 1, Value::number(52.0));  // B5
  const Value inside = EvalSourceAt("=@A:C", wb, wb.sheet(0), 4U, 1U);
  ASSERT_TRUE(inside.is_number());
  EXPECT_EQ(inside.as_number(), 52.0);
}

TEST(ImplicitIntersection, WholeColumnSpanOutsideColumnsIsValue) {
  // Formula at Z5: column 25 is outside A..C, so the intersection is empty.
  // The pre-fix code read the structure-default row of a full-column
  // endpoint and answered #VALUE! for row 1 alone.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(4, 1, Value::number(52.0));
  const Value v = EvalSourceAt("=@A:C", wb, wb.sheet(0), 4U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ImplicitIntersection, WholeRowSpanIntersectsOnBothAxes) {
  // `=@1:3` declares rows 1..3 over the whole grid width.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 4, Value::number(35.0));  // E3
  const Value inside = EvalSourceAt("=@1:3", wb, wb.sheet(0), 2U, 4U);
  ASSERT_TRUE(inside.is_number());
  EXPECT_EQ(inside.as_number(), 35.0);

  const Value outside = EvalSourceAt("=@1:3", wb, wb.sheet(0), 9U, 4U);
  ASSERT_TRUE(outside.is_error());
  EXPECT_EQ(outside.as_error(), ErrorCode::Value);
}

TEST(ImplicitIntersection, SyntacticFormsOfOneRectangleAgree) {
  // The three spellings that declare a rectangle — bounded `Ref:Ref`,
  // full-axis `Ref:Ref`, and the single `Ref` the parser folds `A:A` into —
  // must select the same cell from the same formula position whenever they
  // cover it. B4 is inside all three of A1:B6, A:B and B:B.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(3, 1, Value::number(64.0));  // B4
  const Value bounded = EvalSourceAt("=@A1:B6", wb, wb.sheet(0), 3U, 1U);
  const Value whole_span = EvalSourceAt("=@A:B", wb, wb.sheet(0), 3U, 1U);
  const Value whole_single = EvalSourceAt("=@B:B", wb, wb.sheet(0), 3U, 1U);
  ASSERT_TRUE(bounded.is_number());
  ASSERT_TRUE(whole_span.is_number());
  ASSERT_TRUE(whole_single.is_number());
  EXPECT_EQ(bounded.as_number(), 64.0);
  EXPECT_EQ(whole_span.as_number(), 64.0);
  EXPECT_EQ(whole_single.as_number(), 64.0);
}

TEST(ImplicitIntersection, SingleMatchesAtForWholeAxisForms) {
  // `_xlfn.SINGLE` shares the `@` projection, so neither spelling may keep
  // its own idea of which operand shapes project.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(3, 2, Value::number(34.0));  // C4
  wb.sheet(0).set_cell_value(3, 1, Value::number(24.0));  // B4

  const Value at_single = EvalSourceAt("=@C:C", wb, wb.sheet(0), 3U, 25U);
  const Value fn_single = EvalSourceAt("=_xlfn.SINGLE(C:C)", wb, wb.sheet(0), 3U, 25U);
  ASSERT_TRUE(at_single.is_number());
  ASSERT_TRUE(fn_single.is_number());
  EXPECT_EQ(fn_single.as_number(), at_single.as_number());

  const Value at_span = EvalSourceAt("=@A:C", wb, wb.sheet(0), 3U, 1U);
  const Value fn_span = EvalSourceAt("=_xlfn.SINGLE(A:C)", wb, wb.sheet(0), 3U, 1U);
  ASSERT_TRUE(at_span.is_number());
  ASSERT_TRUE(fn_span.is_number());
  EXPECT_EQ(fn_span.as_number(), at_span.as_number());
}

TEST(ImplicitIntersection, BoundedSingleRefStaysScalar) {
  // `=@A1` is not a rectangle to project: it must read A1 from any formula
  // position rather than demanding the formula sit in row 1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  const Value v = EvalSourceAt("=@A1", wb, wb.sheet(0), 9U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11.0);
}

// ---------------------------------------------------------------------------
// Intersect operator: a genuinely multi-cell overlap is a range and
// materialises wherever `RangeOp` does.
// ---------------------------------------------------------------------------

TEST(IntersectOp, MultiCellOverlapSpills) {
  // C1:D3 and D1:E2 overlap on D1:D2 — a 2x1 rectangle, not a scalar.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 3, Value::number(41.0));  // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(42.0));  // D2
  const Value v = EvalSourceIn("=C1:D3 D1:E2", wb, wb.sheet(0));
  ExpectNumberArray(v, 2U, 1U, {41.0, 42.0});
}

TEST(IntersectOp, MultiCellOverlapSpillsFromAnchoredFormula) {
  // The rectangle does not depend on where the formula sits: an anchored
  // formula inside the overlap's row band no longer collapses onto it.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));  // B1
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));  // B2
  wb.sheet(0).set_cell_value(2, 1, Value::number(32.0));  // B3
  // Formula at Z2 -> row 1, col 25. A1:C3 intersect B1:B5 = B1:B3.
  const Value v = EvalSourceAt("=A1:C3 B1:B5", wb, wb.sheet(0), 1U, 25U);
  ExpectNumberArray(v, 3U, 1U, {12.0, 22.0, 32.0});
}

TEST(IntersectOp, SingleCellOverlapStaysScalar) {
  // A 1x1 overlap degrades to its scalar exactly as `A1:A1` does.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(55.0));  // B2
  const Value v = EvalSourceAt("=A1:C3 B2:B2", wb, wb.sheet(0), 0U, 25U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 55.0);
}

TEST(IntersectOp, DisjointOperandsAreNull) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceAt("=A1:A3 C1:C3", wb, wb.sheet(0), 0U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Null);
}

TEST(IntersectOp, MultiCellOverlapBroadcastsThroughOperator) {
  // The rectangle materialises in a nested position too, so an operator
  // broadcasts over it instead of receiving a collapsed scalar.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 3, Value::number(1.0));  // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(2.0));  // D2
  const Value v = EvalSourceIn("=(C1:D3 D1:E2)*10", wb, wb.sheet(0));
  ExpectNumberArray(v, 2U, 1U, {10.0, 20.0});
}

TEST(IntersectOp, WholeAxisOverlapBehindOperatorKeepsAnchorProjection) {
  // `A1:B1048576 B:B` overlaps on the whole of column B. Only the spilling
  // position materialises a grid-axis rectangle; an operand keeps the
  // top-left anchor projection, exactly as a bare `A:C` does, so an operator
  // cannot conjure a million cells per side.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 1, Value::number(7.0));  // B1
  const Value v = EvalSourceIn("=(A1:B1048576 B:B)+1", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 8.0);
}

TEST(IntersectOp, BlockedFootprintIsSpill) {
  // The spilling position weighs the rectangle against the anchor footprint,
  // so an occupied cell below the anchor blocks it like any other spill.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 3, Value::number(41.0));   // D1
  wb.sheet(0).set_cell_value(1, 3, Value::number(42.0));   // D2
  wb.sheet(0).set_cell_value(1, 25, Value::number(99.0));  // Z2 blocks Z1:Z2.
  const Value v = EvalSourceAt("=C1:D3 D1:E2", wb, wb.sheet(0), 0U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
}

TEST(ImplicitIntersection, DynamicArrayCallCollapsesToTopLeft) {
  Workbook wb = Workbook::create();
  const Value at = EvalSourceAt("=@SEQUENCE(3,1)", wb, wb.sheet(0), 0U, 0U);
  ASSERT_TRUE(at.is_number());
  EXPECT_EQ(at.as_number(), 1.0);

  const Value single = EvalSourceAt("=_xlfn.SINGLE(SEQUENCE(3,1))", wb, wb.sheet(0), 0U, 0U);
  ASSERT_TRUE(single.is_number());
  EXPECT_EQ(single.as_number(), 1.0);
}

TEST(ImplicitIntersection, ArrayLiteralCollapsesToTopLeft) {
  Workbook wb = Workbook::create();
  const Value value = EvalSourceAt("=@{1,2;3,4}", wb, wb.sheet(0), 0U, 0U);
  ASSERT_TRUE(value.is_number());
  EXPECT_DOUBLE_EQ(value.as_number(), 1.0);
}

TEST(ImplicitIntersection, ArrayLiteralSpillsAndBroadcastsOutsideAt) {
  Workbook wb = Workbook::create();
  const Value value = EvalSourceIn("={1,2;3,4}+10", wb, wb.sheet(0));
  ExpectNumberArray(value, 2U, 2U, {11.0, 12.0, 13.0, 14.0});
}

TEST(ImplicitIntersection, SingleMatchesAtForTwoDimensionalRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));  // B2

  const Value at = EvalSourceAt("=@A1:B5", wb, wb.sheet(0), 1U, 1U);
  const Value single = EvalSourceAt("=_xlfn.SINGLE(A1:B5)", wb, wb.sheet(0), 1U, 1U);
  ASSERT_TRUE(at.is_number());
  ASSERT_TRUE(single.is_number());
  EXPECT_EQ(at.as_number(), 22.0);
  EXPECT_EQ(single.as_number(), at.as_number());
}

TEST(IterativeEvaluation, AppliesTopLevelSurfaceContractsAfterFixedPointLoop) {
  Workbook wb = Workbook::create();
  IterativeOptions options;
  options.enabled = true;
  options.max_iterations = 1U;
  wb.set_iterative_options(options);

  // An uninvoked LAMBDA must remain a top-level #CALC! result even when
  // iterative calculation is on.
  const Value lambda = EvalSourceAt("=LAMBDA(x,x+1)", wb, wb.sheet(0), 0U, 0U);
  ASSERT_TRUE(lambda.is_error());
  EXPECT_EQ(lambda.as_error(), ErrorCode::Calc);

  // The read-only spill collision check is likewise a top-level contract.
  wb.sheet(0).set_cell_value(1U, 0U, Value::number(99.0));  // A2 blocks A1:A3.
  const Value spill = EvalSourceAt("=SEQUENCE(3,1)", wb, wb.sheet(0), 0U, 0U);
  ASSERT_TRUE(spill.is_error());
  EXPECT_EQ(spill.as_error(), ErrorCode::Spill);
}

// ---------------------------------------------------------------------------
// Hybrid implicit intersection on bare RangeOp in scalar context.
// When the formula cell is bound and falls inside a 1D range, the
// row/col-aligned cell is returned (legacy II semantics, observationally
// matches IronCalc's cached behavior). Otherwise (out of range, 2D, or
// no formula cell bound) we fall back to the top-left, matching Mac's
// spill anchor.
// Verified Mac semantics: tests/oracle/cases/implicit_intersection.yaml.
// ---------------------------------------------------------------------------

TEST(RangeOp, SpillsRegardlessOfFormulaRow) {
  // Excel 365 spills a bare range independent of the formula cell's row;
  // a formula bound inside the range span no longer projects onto it.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(5.0));
  // Formula at L3 -> row=2, col=11; row 2 is inside [0..4] yet still spills.
  const Value v = EvalSourceAt("=A1:A5", wb, wb.sheet(0), 2U, 11U);
  ExpectNumberArray(v, 5U, 1U, {1.0, 2.0, 3.0, 4.0, 5.0});
}

TEST(RangeOp, SpillsRegardlessOfFormulaColumn) {
  // Symmetric to the row case: a single-row range spills regardless of the
  // formula cell's column falling inside the span.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(40.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(50.0));
  // Formula at C9 -> row=8, col=2; col 2 is inside [0..4] yet still spills.
  const Value v = EvalSourceAt("=A1:E1", wb, wb.sheet(0), 8U, 2U);
  ExpectNumberArray(v, 1U, 5U, {10.0, 20.0, 30.0, 40.0, 50.0});
}

TEST(RangeOp, SpillsWithFormulaCellOutsideRange) {
  // Single-column range A3:A8 with the formula bound outside the span
  // (row 0). Legacy implicit intersection would #VALUE!; 365 spills.
  Workbook wb = test::mac_workbook();
  wb.sheet(0).set_cell_value(2, 0, Value::number(33.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(44.0));
  wb.sheet(0).set_cell_value(7, 0, Value::number(88.0));
  // Formula at L1 -> row=0, col=11.
  const Value v = EvalSourceAt("=A3:A8", wb, wb.sheet(0), 0U, 11U);
  ExpectNumberArray(v, 6U, 1U, {33.0, 44.0, 0.0, 0.0, 0.0, 88.0});
}

TEST(RangeOp, Spills2DRegardlessOfFormulaCell) {
  // 2D range A1:B5 spills the full rectangle; the formula cell position is
  // irrelevant. Unset tail cells surface as 0.
  Workbook wb = test::mac_workbook();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));
  // Formula at Z1 -> row=0, col=25.
  const Value v = EvalSourceAt("=A1:B5", wb, wb.sheet(0), 0U, 25U);
  ExpectNumberArray(v, 5U, 2U, {11.0, 12.0, 21.0, 22.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0});
}

TEST(RangeOp, SpillsWithoutFormulaCell) {
  // No formula cell bound (parser-driven smoke path): still spills.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(99.0));
  const Value v = EvalSourceIn("=A1:A5", wb, wb.sheet(0));
  ExpectNumberArray(v, 5U, 1U, {42.0, 0.0, 99.0, 0.0, 0.0});
}

TEST(RangeOp, ScalarAtPrefixStrictUnchanged) {
  // The @-prefix path must remain strict: row 0 is outside [2..7], so
  // implicit intersection on @A3:A8 still returns #VALUE!. This guards
  // against any accidental coupling of the new RangeOp branch with the
  // ImplicitIntersection branch.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 0, Value::number(33.0));
  wb.sheet(0).set_cell_value(7, 0, Value::number(88.0));
  // Formula at L1 -> row=0, col=11.
  const Value v = EvalSourceAt("=@A3:A8", wb, wb.sheet(0), 0U, 11U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
