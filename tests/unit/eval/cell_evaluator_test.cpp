// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `evaluate_cell_for_recalc`. The function is exercised
// indirectly through `Workbook::recalc` (the recalc engine and the
// scheduler both route through it), but a handful of direct tests guard
// the small surface — formula-text empty, parser failure, error
// propagation, and the array-spill commit branch — that the indirect
// coverage exercises only by happy accident.

#include "eval/cell_evaluator.h"

#include <cstdint>

#include "cell.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Builds a workbook with a single sheet named "Sheet1" and exposes it
// for direct mutation. Mirrors the helper used by `recalc_engine_test`.
Workbook MakeWorkbookWithSheet() {
  Workbook wb = Workbook::create();
  return wb;
}

TEST(EvaluateCellForRecalc, BarePrefixIsStripped) {
  // Confirms `strip_formula_prefix` is wired through the helper: a
  // formula stored with the leading `=` parses identically to one
  // stored without (the writer is inconsistent across paths).
  Workbook wb = MakeWorkbookWithSheet();
  Sheet& sheet = wb.sheet(0U);
  Cell cell;
  cell.formula_text = "=1+2";

  Arena arena;
  const Value v = evaluate_cell_for_recalc(wb, sheet, cell, /*row=*/0U, /*col=*/0U, default_registry(), arena);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(EvaluateCellForRecalc, ParseFailureSurfacesNameError) {
  // An empty formula text is the documented parser-failure case (no
  // expression to consume). Should surface #NAME? to mirror
  // `EvalContext::resolve_ref`.
  Workbook wb = MakeWorkbookWithSheet();
  Sheet& sheet = wb.sheet(0U);
  Cell cell;
  cell.formula_text = "=";  // strip leaves an empty string

  Arena arena;
  const Value v = evaluate_cell_for_recalc(wb, sheet, cell, 0U, 0U, default_registry(), arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(EvaluateCellForRecalc, ErrorPropagatesFromExpression) {
  // Division by zero inside the expression must propagate verbatim;
  // the helper should not swallow or rewrite Excel-visible errors.
  Workbook wb = MakeWorkbookWithSheet();
  Sheet& sheet = wb.sheet(0U);
  Cell cell;
  cell.formula_text = "=1/0";

  Arena arena;
  const Value v = evaluate_cell_for_recalc(wb, sheet, cell, 0U, 0U, default_registry(), arena);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(EvaluateCellForRecalc, ArrayResultSpillsAndReturnsAnchor) {
  // SEQUENCE returns an Array at the top level; the helper should
  // commit the spill via `dispatch_array_result` and surface the
  // anchor scalar (the first cell of the spill region).
  Workbook wb = test::mac_workbook();
  Sheet& sheet = wb.sheet(0U);
  Cell cell;
  cell.formula_text = "=SEQUENCE(3)";

  Arena arena;
  const Value v = evaluate_cell_for_recalc(wb, sheet, cell, /*row=*/0U, /*col=*/0U, default_registry(), arena);
  // Anchor scalar is the first SEQUENCE element (1).
  ASSERT_TRUE(v.is_number()) << "SEQUENCE anchor should be the first element";
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);

  // The spill region must exist on the sheet anchored at (0, 0).
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 1U);
}

TEST(EvaluateCellForRecalc, IterativeModeReadsCachedValueInsteadOfRecursing) {
  // In iterative mode the EvalContext is built without an EvalState,
  // so a self-referential formula must read the cell's cached value
  // instead of recursing. Set A1's cached value to 7 and evaluate
  // `=A1+1` at A2 in iterative mode — should yield 8.
  Workbook wb = MakeWorkbookWithSheet();
  Sheet& sheet = wb.sheet(0U);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(7.0))));

  Cell cell;
  cell.formula_text = "=A1+1";

  EvaluateCellOptions opts;
  opts.iterative_mode = true;
  Arena arena;
  const Value v = evaluate_cell_for_recalc(wb, sheet, cell, /*row=*/1U, /*col=*/0U, default_registry(), arena, opts);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 8.0);
}

}  // namespace
}  // namespace formulon::eval
