// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end integration tests for the workbook recalc loop. These tests
// drive the public Workbook API (`set_cell_value`, `set_cell_formula`,
// `recalc`) and observe results via `Sheet::cell_at` and
// `Sheet::resolve_cell_value`. Unlike the unit tests for the engine, these
// scenarios mix multiple cells, dynamic-array spills, and spill collisions
// to exercise the full pipeline.

#include <cstdint>
#include <string>
#include <vector>

#include "cell.h"
#include "eval/compat.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Convenience: returns the sheet's stored cached_value (does not consult
// the spill table). For spill phantoms / anchors that need to surface the
// post-spill value, use `Sheet::resolve_cell_value` directly.
Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

TEST(WorkbookRecalc, SmallChainEndToEnd) {
  // A simple linear chain: A1 = 10, A2 = =A1*2, A3 = =A2+5.
  //
  // After recalc, A2 == 20 and A3 == 25. Also verify dirty propagation:
  // mutating A1 to 20 yields A2 == 40 and A3 == 45.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1*2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 0U, "=A2+5")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value a2 = StoredValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 20.0);

  Value a3 = StoredValue(wb, 0U, 2U, 0U);
  ASSERT_TRUE(a3.is_number());
  EXPECT_DOUBLE_EQ(a3.as_number(), 25.0);

  // Mutate A1 and recalc again.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  a2 = StoredValue(wb, 0U, 1U, 0U);
  a3 = StoredValue(wb, 0U, 2U, 0U);
  ASSERT_TRUE(a2.is_number());
  ASSERT_TRUE(a3.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 40.0);
  EXPECT_DOUBLE_EQ(a3.as_number(), 45.0);
}

TEST(WorkbookRecalc, ValidPrefixWithTrailingGarbageIsNameError) {
  // A formula that parses to a valid prefix but leaves unconsumed trailing
  // tokens must NOT evaluate as the prefix — that would silently change the
  // cell's value and dependency set from what was entered. The strict parse
  // gate surfaces #NAME? and registers no dependency on the prefix's refs.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(5.0))));  // A1 = 5
  // "=A1 3" parses to the prefix `A1` with a trailing `3` the parser cannot
  // consume. Pre-fix this evaluated to 5 (the prefix) and wired a dep on A1.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1 3")));  // B1

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_error()) << "expected #NAME?, got a non-error value from a recovered prefix";
  EXPECT_EQ(b1.as_error(), ErrorCode::Name);

  // No stale dependency: mutating A1 must not resurrect the prefix value.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  b1 = StoredValue(wb, 0U, 0U, 1U);
  EXPECT_TRUE(b1.is_error()) << "trailing-garbage formula must stay #NAME? after upstream edit";
}

TEST(WorkbookRecalc, SumAcrossRange) {
  // B1 = =SUM(A1:A3) — exercises the range-flattening path in the dep
  // extractor. Filling A1..A3 with 1..3 should produce B1 == 6.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 0U, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 6.0);

  // Mutate A2 and recalc; B1 must reflect the new sum.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 24.0);
}

TEST(WorkbookRecalc, DynamicArraySpillsToAdjacentCells) {
  // A1 = =SEQUENCE(1,3) emits a 1x3 row array that should spill into A1,
  // B1, C1 with values 1, 2, 3. The anchor cached_value is the first
  // array cell; phantoms surface through `Sheet::resolve_cell_value`.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // Anchor cell (A1) holds the first array cell.
  Value a1 = wb.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1.is_number()) << "A1 expected to be a number, got error="
                              << (a1.is_error() ? static_cast<int>(a1.as_error()) : -1);
  EXPECT_DOUBLE_EQ(a1.as_number(), 1.0);

  // Phantoms B1, C1 surface via the spill table.
  Value b1 = wb.sheet(0).resolve_cell_value(0U, 1U);
  Value c1 = wb.sheet(0).resolve_cell_value(0U, 2U);
  ASSERT_TRUE(b1.is_number());
  ASSERT_TRUE(c1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c1.as_number(), 3.0);
}

TEST(WorkbookRecalc, SpillCollisionSurfacesSpillError) {
  // B1 = 99 (literal). A1 = =SEQUENCE(1,3) would normally spill into B1,
  // but the collision must surface as #SPILL! at the anchor.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_error()) << "A1 expected to be #SPILL!";
  EXPECT_EQ(a1.as_error(), ErrorCode::Spill);

  // The colliding literal at B1 must be preserved.
  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 99.0);
}

}  // namespace
}  // namespace formulon
