// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `RecalcEngine::partial_recalc`. The tests drive the
// engine through the workbook public API so the dep graph state is
// always populated through the same mutation paths a real client uses.

#include <cstdint>
#include <string>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Returns the cached value at `(row, col)` on `sheet_index`, or
// `Value::blank()` when the cell is absent.
Value CellValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& sheet = wb.sheet(sheet_index);
  if (const Cell* c = sheet.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

// Builds a viewport rectangle covering a single cell at (row, col) on sheet 0.
SheetCellRange SingleCell(std::uint32_t row, std::uint32_t col) {
  SheetCellRange r;
  r.sheet_id = 0;
  r.first_row = row;
  r.last_row = row;
  r.first_col = col;
  r.last_col = col;
  return r;
}

TEST(PartialRecalc, EmptyViewportIsNoop) {
  // first_row > last_row collapses the viewport to empty; the engine
  // should return immediately without touching any cell.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1+1")));

  SheetCellRange empty;
  empty.sheet_id = 0;
  empty.first_row = 5;
  empty.last_row = 4;  // collapsed
  empty.first_col = 0;
  empty.last_col = 0;

  auto stats = wb.partial_recalc(default_registry(), empty);
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 0u);
  EXPECT_EQ(stats.value().volatile_cells, 0u);
  // A2 was never evaluated, so its cached value is still blank.
  Value v = CellValue(wb, 0U, 1U, 0U);
  EXPECT_TRUE(v.is_blank());
  // The dirty flag remains so a later full recalc still picks A2 up.
  auto full = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(full));
  EXPECT_EQ(full.value().cells_evaluated, 1u);
  v = CellValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(PartialRecalc, SingleCellViewportRecomputesAncestorsOnly) {
  // A1 = 10, B1 = =A1+1 (not in viewport), C1 = =A1*2 (not in
  // viewport). Viewport asks for B1 only — only A1 (literal) and B1
  // are in the closure. C1 should remain unevaluated this pass.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));   // B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=A1*2")));   // C1

  // Viewport: B1 only.
  auto stats = wb.partial_recalc(default_registry(), SingleCell(0U, 1U));
  ASSERT_TRUE(static_cast<bool>(stats));
  // Only B1 was a formula in the closure: 1 evaluation expected.
  EXPECT_EQ(stats.value().cells_evaluated, 1u);
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 1U).is_number());
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 11.0);
  // C1 is OUTSIDE the closure: still blank, still dirty.
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 2U).is_blank());

  // A subsequent full recalc must visit C1.
  auto full = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(full));
  // C1 is still dirty; B1 is no longer dirty so only C1 evaluates.
  EXPECT_EQ(full.value().cells_evaluated, 1u);
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 2U).is_number());
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 20.0);
}

TEST(PartialRecalc, MultiCellViewportUnionsClosures) {
  // A1 = 1, A2 = 2.
  // B1 = =A1, B2 = =A2, C1 = =B1+1, C2 = =B2+1, D1 = =A1*100 (untouched).
  // Viewport: B1:B2. Closure must include {B1, A1, B2, A2}; C1 / C2 /
  // D1 are NOT in the closure (they read FROM B1/B2 but the viewport
  // only asks for B1/B2's values, not their dependents).
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1")));    // B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=A2")));    // B2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=B1+1")));  // C1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 2U, "=B2+1")));  // C2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=A1*100")));  // D1

  // Viewport: B1:B2.
  SheetCellRange vp;
  vp.sheet_id = 0;
  vp.first_row = 0;
  vp.last_row = 1;
  vp.first_col = 1;
  vp.last_col = 1;
  auto stats = wb.partial_recalc(default_registry(), vp);
  ASSERT_TRUE(static_cast<bool>(stats));
  // B1 and B2 are formulas in the closure: 2 evaluations expected.
  // C1, C2, D1 are NOT in the closure, so they stay dirty / blank.
  EXPECT_EQ(stats.value().cells_evaluated, 2u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 1U).as_number(), 2.0);
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 2U).is_blank());
  EXPECT_TRUE(CellValue(wb, 0U, 1U, 2U).is_blank());
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 3U).is_blank());
}

TEST(PartialRecalc, OutsideClosureStaysDirty) {
  // A1 = 1, B1 = =A1+1 (in viewport), C1 = =A1*5 (outside viewport).
  // After partial_recalc(B1), C1 must remain dirty so the next
  // recalc() picks it up.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=A1*5")));

  // First partial recalc visits only B1.
  auto stats = wb.partial_recalc(default_registry(), SingleCell(0U, 1U));
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 4.0);
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 2U).is_blank());

  // Subsequent full recalc must visit C1.
  auto full = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(full));
  EXPECT_EQ(full.value().cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 15.0);
}

TEST(PartialRecalc, VolatileInsideClosureRecomputed) {
  // B1 = =NOW() (volatile, inside viewport). Each partial_recalc must
  // re-execute B1 because it is a volatile cell within the closure.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=NOW()")));
  // First pass: evaluates the volatile.
  auto stats1 = wb.partial_recalc(default_registry(), SingleCell(0U, 1U));
  ASSERT_TRUE(static_cast<bool>(stats1));
  EXPECT_EQ(stats1.value().volatile_cells, 1u);
  EXPECT_EQ(stats1.value().cells_evaluated, 1u);

  // Second pass with no other mutation: still re-executes the volatile.
  auto stats2 = wb.partial_recalc(default_registry(), SingleCell(0U, 1U));
  ASSERT_TRUE(static_cast<bool>(stats2));
  EXPECT_EQ(stats2.value().volatile_cells, 1u);
  EXPECT_EQ(stats2.value().cells_evaluated, 1u);
}

TEST(PartialRecalc, VolatileOutsideClosureNotForced) {
  // A1 = =NOW() (volatile, OUTSIDE viewport). B1 = =42 (inside).
  // partial_recalc(B1) must NOT force A1 to re-evaluate.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=NOW()")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=42")));

  auto stats = wb.partial_recalc(default_registry(), SingleCell(0U, 1U));
  ASSERT_TRUE(static_cast<bool>(stats));
  // A1 is volatile but outside the closure: must NOT count.
  EXPECT_EQ(stats.value().volatile_cells, 0u);
  // Only B1 should evaluate.
  EXPECT_EQ(stats.value().cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 42.0);
  // A1 stays blank (never evaluated this pass).
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 0U).is_blank());
}

TEST(PartialRecalc, OutOfRangeSheetIsNoop) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=42")));
  SheetCellRange vp;
  vp.sheet_id = 99;  // out of range
  vp.first_row = 0;
  vp.last_row = 0;
  vp.first_col = 0;
  vp.last_col = 0;
  auto stats = wb.partial_recalc(default_registry(), vp);
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 0u);
  // A1 stays blank (no recalc happened).
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 0U).is_blank());
}

TEST(PartialRecalc, ViewportClampedToSheetExtent) {
  // Viewport rectangle extends beyond populated cells; the engine
  // tolerates the over-extent (no out-of-bounds reads) and only
  // recomputes formulas within the closure.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1*2")));

  // Viewport stretches well beyond the sheet's stored extent.
  SheetCellRange vp;
  vp.sheet_id = 0;
  vp.first_row = 0;
  vp.last_row = 50;
  vp.first_col = 0;
  vp.last_col = 50;
  auto stats = wb.partial_recalc(default_registry(), vp);
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 14.0);
}

}  // namespace
}  // namespace formulon::eval
