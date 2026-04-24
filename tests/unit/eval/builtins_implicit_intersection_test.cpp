// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for Excel 365 dynamic-array spill semantics on bare ranges
// (`=A1:A5`) and implicit intersection on `@`-prefixed ranges
// (`=@A1:A5`). Verified Mac semantics:
// `tests/oracle/cases/implicit_intersection.yaml` and corresponding
// `tests/oracle/golden/implicit_intersection.golden.json`.
//
// Bare range as a scalar: returns the top-left endpoint (the spill
// anchor visible to a single-cell reader such as xlwings).
// `@`-prefixed range: projects the range onto the formula cell's row or
// column. Single-column range requires the formula row to fall inside
// the range; single-row range requires the formula column to fall
// inside the range; 2D ranges return `#VALUE!`.

#include <cstdint>
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
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Evaluate `src` against a bound workbook anchored at the formula cell
// (`row`, `col`) on the current sheet. Used for the `@`-prefixed
// implicit-intersection cases.
Value EvalSourceAt(std::string_view src, const Workbook& wb,
                   const Sheet& current, std::uint32_t row, std::uint32_t col) {
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
  const EvalContext ctx = EvalContext(wb, current, state).with_formula_cell(row, col);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Bare range -> dynamic-array spill anchor (top-left of the range).
// ---------------------------------------------------------------------------

TEST(RangeOp, SpillAnchorTopLeftSingleCol) {
  // =A1:A5 with A1=42 -> 42 (anchor = top-left).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(99.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(100.0));
  const Value v = EvalSourceIn("=A1:A5", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(RangeOp, SpillAnchorTopLeftSingleRow) {
  // =A1:E1 with A1=1, E1=5 -> 1 (anchor = top-left).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(5.0));
  const Value v = EvalSourceIn("=A1:E1", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(RangeOp, SpillAnchorTopLeft2D) {
  // =A1:B5 with A1=11 -> 11 (anchor = top-left of 2D rectangle).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));
  const Value v = EvalSourceIn("=A1:B5", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11.0);
}

TEST(RangeOp, SpillAnchorReverseOrder) {
  // =A5:A1 -- endpoints swapped. Should still resolve to A1 (top-left
  // after normalisation).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(99.0));
  const Value v = EvalSourceIn("=A5:A1", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
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

TEST(ImplicitIntersection, TwoDReturnsValue) {
  // 2D range with implicit intersection -- conservative #VALUE!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(11.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(12.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(21.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(22.0));
  // Z3 -> row=2, col=25; range spans rows 0..4 cols 0..1, neither axis
  // aligns under our conservative 2D rule.
  const Value v = EvalSourceAt("=@A1:B5", wb, wb.sheet(0), 2U, 25U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
