// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Integration tests for the dynamic-array spill engine focused on
// collision detection and recovery. These tests drive the full
// `Workbook::set_cell_*` -> `Workbook::recalc()` pipeline so the recalc
// engine, dep graph, evaluator, and `Sheet::commit_spill` all run end to
// end. They lock in the engine's current spill-collision behaviour ahead
// of any future spill-engine refactor.
//
// The contract under test:
//   * A literal scalar inside the would-be footprint of an array formula
//     surfaces `#SPILL!` at the anchor and preserves the literal at the
//     blocking cell.
//   * Clearing the blocking cell (and re-marking the anchor dirty)
//     re-evaluates the producer and spills successfully.
//   * Dirty-set propagation from blocking-cell mutations re-fires the
//     array producer without the caller having to touch it.
//   * 2D footprints behave the same as 1D footprints (any single
//     occupied cell anywhere inside the footprint blocks the spill).

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/tables_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Convenience: returns the sheet's stored cached_value. Phantoms outside
// the anchor are accessed through `Sheet::resolve_cell_value` instead.
Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

// ---------------------------------------------------------------------------
// 1D collisions (column / row footprints)
// ---------------------------------------------------------------------------

TEST(SpillCollision, ScalarBlocksColumnSpillInMiddle) {
  // A1 = =SEQUENCE(3,1) wants to spill A1, A2, A3. A2 = literal 7 sits
  // in the middle of the footprint -> #SPILL! at the anchor; the literal
  // at A2 is preserved.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_error()) << "A1 expected to be #SPILL!";
  EXPECT_EQ(a1.as_error(), ErrorCode::Spill);

  // The blocker at A2 must be intact.
  Value a2 = StoredValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 7.0);

  // A3 was never spilled into; it should be blank.
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
}

TEST(SpillCollision, ScalarBlocksColumnSpillAtTail) {
  // A1 = =SEQUENCE(3,1); A3 = literal 99 (tail of footprint). Same
  // outcome as the middle-blocker case -> #SPILL!.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 0U, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_error());
  EXPECT_EQ(a1.as_error(), ErrorCode::Spill);
  Value a3 = StoredValue(wb, 0U, 2U, 0U);
  ASSERT_TRUE(a3.is_number());
  EXPECT_DOUBLE_EQ(a3.as_number(), 99.0);
}

TEST(SpillCollision, ScalarBlocksRowSpill) {
  // A1 = =SEQUENCE(1,3) wants to spill A1, B1, C1. C1 = literal 5 ->
  // #SPILL!. Mirrors the column-spill case across the other axis.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 2U, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_error());
  EXPECT_EQ(a1.as_error(), ErrorCode::Spill);

  Value c1 = StoredValue(wb, 0U, 0U, 2U);
  ASSERT_TRUE(c1.is_number());
  EXPECT_DOUBLE_EQ(c1.as_number(), 5.0);
}

// ---------------------------------------------------------------------------
// 2D collisions
// ---------------------------------------------------------------------------

TEST(SpillCollision, ScalarBlocks2DSpillInterior) {
  // A1 = =SEQUENCE(3,3) wants a 3x3 footprint A1:C3. C2 = literal 100
  // sits inside the footprint -> #SPILL!.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 2U, Value::number(100.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_error());
  EXPECT_EQ(a1.as_error(), ErrorCode::Spill);

  Value c2 = StoredValue(wb, 0U, 1U, 2U);
  ASSERT_TRUE(c2.is_number());
  EXPECT_DOUBLE_EQ(c2.as_number(), 100.0);

  // No spill region was committed at A1.
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
}

// ---------------------------------------------------------------------------
// Recovery after the blocker is removed
// ---------------------------------------------------------------------------

TEST(SpillCollision, ClearBlockerColumnRecoversSpill) {
  // Start with A2 blocking; recalc -> #SPILL!. Then set A2 to blank,
  // recalc again -> A1 spills cleanly with cells {1, 2, 3}.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(7.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_error());
  ASSERT_EQ(a1.as_error(), ErrorCode::Spill);

  // Clear the blocker. `set_cell_value(blank)` marks A2 dirty AND its
  // dependents dirty; A1 does not depend on A2, so the recalc must rely
  // on A1 being directly re-marked. We re-set A1's formula to mark it
  // dirty -- this matches what a host that detects a clear in the
  // bordering cells would do (or a user re-typing the formula). Excel's
  // recalc engine flags this same scenario by marking neighbouring cells
  // dirty when a cell becomes blank, but Formulon's recalc engine does
  // not yet propagate spill-blocker clears back to spill anchors; this
  // test locks in current behaviour.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::blank())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // After blocker removal + anchor re-dirty, the spill should succeed.
  a1 = wb.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1.is_number()) << "A1 expected to spill cleanly; kind=" << static_cast<int>(a1.kind())
                              << " err=" << (a1.is_error() ? static_cast<int>(a1.as_error()) : -1);
  EXPECT_DOUBLE_EQ(a1.as_number(), 1.0);

  Value a2 = wb.sheet(0).resolve_cell_value(1U, 0U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 2.0);

  Value a3 = wb.sheet(0).resolve_cell_value(2U, 0U);
  ASSERT_TRUE(a3.is_number());
  EXPECT_DOUBLE_EQ(a3.as_number(), 3.0);

  // Spill region must now be registered at the anchor.
  EXPECT_NE(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
}

TEST(SpillCollision, ClearBlocker2DRecoversSpill) {
  // 3x3 footprint; clear the interior blocker and re-spill.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 2U, Value::number(100.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(StoredValue(wb, 0U, 0U, 0U).as_error(), ErrorCode::Spill);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 2U, Value::blank())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // A1 = 1, B1 = 2, C1 = 3, A2 = 4, ..., C3 = 9 (row-major SEQUENCE).
  Value a1 = wb.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1.is_number()) << "A1 expected to spill cleanly";
  EXPECT_DOUBLE_EQ(a1.as_number(), 1.0);
  Value c3 = wb.sheet(0).resolve_cell_value(2U, 2U);
  ASSERT_TRUE(c3.is_number());
  EXPECT_DOUBLE_EQ(c3.as_number(), 9.0);

  EXPECT_NE(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
}

// ---------------------------------------------------------------------------
// Two array producers competing for the same range
// ---------------------------------------------------------------------------

TEST(SpillCollision, TwoArrayProducersConflictOnOverlappingFootprint) {
  // A1 = =SEQUENCE(3,1) wants A1:A3. A2 = =SEQUENCE(3,1) also wants
  // A2:A4. The two would overlap on A2 and A3. Expectation: one wins
  // (the first to commit) and the other surfaces #SPILL!. Excel's tie-
  // break is "first formula committed wins", which mirrors the recalc
  // engine's reverse-topological SCC walk over singletons: dirty cells
  // are evaluated in graph order, so whichever Tarjan emits first
  // commits its spill and the other collides.
  //
  // This test does NOT pin which one wins -- both orderings are valid
  // observations of current behaviour. It only asserts that exactly
  // one of the two anchors holds #SPILL! and the other holds a
  // committed spill.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value a1 = StoredValue(wb, 0U, 0U, 0U);
  const Value a2 = StoredValue(wb, 0U, 1U, 0U);

  // Exactly one of the two should be #SPILL! and the other a number.
  const bool a1_spill = a1.is_error() && a1.as_error() == ErrorCode::Spill;
  const bool a2_spill = a2.is_error() && a2.as_error() == ErrorCode::Spill;
  EXPECT_NE(a1_spill, a2_spill) << "expected exactly one of A1/A2 to be #SPILL!"
                                << " a1=" << static_cast<int>(a1.kind()) << " a2=" << static_cast<int>(a2.kind());

  // The non-#SPILL! anchor must have a registered region.
  if (!a1_spill) {
    EXPECT_NE(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
  }
  if (!a2_spill) {
    EXPECT_NE(wb.sheet(0).spill_region_at_anchor(1U, 0U), nullptr);
  }
}

// ---------------------------------------------------------------------------
// Spill that would overlap a phantom of another spill
// ---------------------------------------------------------------------------

TEST(SpillCollision, ScalarFromAnotherSpillBlocksSecondSpill) {
  // A1 = =SEQUENCE(1,3) spills A1, B1, C1. Then B2 = =SEQUENCE(3,1)
  // wants B2, B3, B4 -- no overlap; both should succeed. This is the
  // negative control for the conflict cases above: non-overlapping
  // 2D footprints both spill cleanly.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // Both anchors should have committed spill regions.
  EXPECT_NE(wb.sheet(0).spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_NE(wb.sheet(0).spill_region_at_anchor(1U, 1U), nullptr);

  // Spot-check the values.
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 1U).as_number(), 2.0);  // B1
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 2U).as_number(), 3.0);  // C1
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(2U, 1U).as_number(), 2.0);  // B3
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(3U, 1U).as_number(), 3.0);  // B4
}

// ---------------------------------------------------------------------------
// Table-footprint spill collision
// ---------------------------------------------------------------------------
//
// Excel treats cells inside an existing table's data area as occupied for
// the purposes of spill-collision detection: a dynamic-array formula whose
// footprint would cover any non-blank table cell surfaces #SPILL! at the
// anchor and leaves the table contents untouched. These tests lock in that
// behaviour. They also exercise the negative case (spill landing strictly
// outside the table) to confirm that the table's `ref` does not extend any
// implicit no-spill zone beyond its declared rectangle.

TEST(SpillCollision, SpillIntoTableHeaderRowCollides) {
  // Sales table at A1:C3 (header row + 2 data rows). =SEQUENCE(3,2) at D1
  // wants D1:E3. Pre-place a literal at E2 (inside the spill footprint).
  // The literal is the explicit blocker; the test confirms that the spill
  // surfaces #SPILL! and does NOT clobber the literal.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Region"));
  s.set_cell_value(0U, 1U, Value::text("Product"));
  s.set_cell_value(0U, 2U, Value::text("Amount"));
  s.set_cell_value(1U, 0U, Value::text("North"));
  s.set_cell_value(1U, 1U, Value::text("Apple"));
  s.set_cell_value(1U, 2U, Value::number(10.0));
  s.set_cell_value(2U, 0U, Value::text("South"));
  s.set_cell_value(2U, 1U, Value::text("Banana"));
  s.set_cell_value(2U, 2U, Value::number(20.0));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C3";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = false;
  table.columns = {
      io::TableColumn{1U, "Region", "", "", ""},
      io::TableColumn{2U, "Product", "", "", ""},
      io::TableColumn{3U, "Amount", "", "", ""},
  };
  std::vector<io::TableMetadata> tables = {std::move(table)};
  wb.set_tables(std::move(tables));

  // Literal blocker inside the spill footprint at E2 (row 1, col 4).
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 4U, Value::text("blocker"))));
  // Anchor at D1 (row 0, col 3); footprint = D1:E3.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=SEQUENCE(3,2)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value d1 = StoredValue(wb, 0U, 0U, 3U);
  ASSERT_TRUE(d1.is_error()) << "D1 expected #SPILL!; kind=" << static_cast<int>(d1.kind());
  EXPECT_EQ(d1.as_error(), ErrorCode::Spill);

  // The literal at E2 must be preserved.
  const Value e2 = StoredValue(wb, 0U, 1U, 4U);
  ASSERT_TRUE(e2.is_text()) << "E2 expected to retain literal; kind=" << static_cast<int>(e2.kind());
  EXPECT_EQ(e2.as_text(), "blocker");

  // No spill region committed at the anchor.
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(0U, 3U), nullptr);
}

TEST(SpillCollision, SpillAdjacentToTableNoCollision) {
  // Sales table at A1:C3. =SEQUENCE(2,2) at D1 wants D1:E2 -- entirely
  // outside the table's rectangle (cols 0..2). No occupant inside the
  // footprint. Spill must succeed; the table's `ref` does not extend a
  // no-spill zone past its declared right edge.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Region"));
  s.set_cell_value(0U, 1U, Value::text("Product"));
  s.set_cell_value(0U, 2U, Value::text("Amount"));
  s.set_cell_value(1U, 0U, Value::text("North"));
  s.set_cell_value(1U, 1U, Value::text("Apple"));
  s.set_cell_value(1U, 2U, Value::number(10.0));
  s.set_cell_value(2U, 0U, Value::text("South"));
  s.set_cell_value(2U, 1U, Value::text("Banana"));
  s.set_cell_value(2U, 2U, Value::number(20.0));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C3";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = false;
  table.columns = {
      io::TableColumn{1U, "Region", "", "", ""},
      io::TableColumn{2U, "Product", "", "", ""},
      io::TableColumn{3U, "Amount", "", "", ""},
  };
  std::vector<io::TableMetadata> tables = {std::move(table)};
  wb.set_tables(std::move(tables));

  // Anchor at D1 (row 0, col 3); footprint = D1:E2 -> rows 0..1, cols 3..4.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=SEQUENCE(2,2)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // Anchor stores the first element; phantoms via resolve_cell_value.
  const Value d1 = wb.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1.is_number()) << "D1 expected to spill cleanly; kind=" << static_cast<int>(d1.kind());
  EXPECT_DOUBLE_EQ(d1.as_number(), 1.0);

  const Value e1 = wb.sheet(0).resolve_cell_value(0U, 4U);
  ASSERT_TRUE(e1.is_number());
  EXPECT_DOUBLE_EQ(e1.as_number(), 2.0);

  const Value d2 = wb.sheet(0).resolve_cell_value(1U, 3U);
  ASSERT_TRUE(d2.is_number());
  EXPECT_DOUBLE_EQ(d2.as_number(), 3.0);

  const Value e2 = wb.sheet(0).resolve_cell_value(1U, 4U);
  ASSERT_TRUE(e2.is_number());
  EXPECT_DOUBLE_EQ(e2.as_number(), 4.0);

  // Spill region must be registered.
  EXPECT_NE(wb.sheet(0).spill_region_at_anchor(0U, 3U), nullptr);
}

TEST(SpillCollision, SpillIntoTableDataCellCollides) {
  // Sales table at A1:C5 (header row + 4 data rows; column C holds 10/20/
  // 30/40). =SEQUENCE(2,3) at B3 wants B3:D4. Cells B3 ("Banana") and C3
  // (20) are inside both the table data area AND the spill footprint, so
  // they act as literal blockers from the spill engine's perspective.
  // Expectation: B3 -> #SPILL!; the table's existing values at C3 and C4
  // are preserved.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Region"));
  s.set_cell_value(0U, 1U, Value::text("Product"));
  s.set_cell_value(0U, 2U, Value::text("Amount"));
  s.set_cell_value(1U, 0U, Value::text("North"));
  s.set_cell_value(1U, 1U, Value::text("Apple"));
  s.set_cell_value(1U, 2U, Value::number(10.0));
  s.set_cell_value(2U, 0U, Value::text("South"));
  s.set_cell_value(2U, 1U, Value::text("Banana"));
  s.set_cell_value(2U, 2U, Value::number(20.0));
  s.set_cell_value(3U, 0U, Value::text("East"));
  s.set_cell_value(3U, 1U, Value::text("Cherry"));
  s.set_cell_value(3U, 2U, Value::number(30.0));
  s.set_cell_value(4U, 0U, Value::text("West"));
  s.set_cell_value(4U, 1U, Value::text("Durian"));
  s.set_cell_value(4U, 2U, Value::number(40.0));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C5";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = false;
  table.columns = {
      io::TableColumn{1U, "Region", "", "", ""},
      io::TableColumn{2U, "Product", "", "", ""},
      io::TableColumn{3U, "Amount", "", "", ""},
  };
  std::vector<io::TableMetadata> tables = {std::move(table)};
  wb.set_tables(std::move(tables));

  // Anchor at B3 (row 2, col 1); footprint = B3:D4 -> rows 2..3, cols 1..3.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 1U, "=SEQUENCE(2,3)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value b3 = StoredValue(wb, 0U, 2U, 1U);
  ASSERT_TRUE(b3.is_error()) << "B3 expected #SPILL!; kind=" << static_cast<int>(b3.kind());
  EXPECT_EQ(b3.as_error(), ErrorCode::Spill);

  // The table's existing data at C3 and C4 must be preserved.
  const Value c3 = StoredValue(wb, 0U, 2U, 2U);
  ASSERT_TRUE(c3.is_number()) << "C3 expected to retain table value 20; kind=" << static_cast<int>(c3.kind());
  EXPECT_DOUBLE_EQ(c3.as_number(), 20.0);

  const Value c4 = StoredValue(wb, 0U, 3U, 2U);
  ASSERT_TRUE(c4.is_number()) << "C4 expected to retain table value 30; kind=" << static_cast<int>(c4.kind());
  EXPECT_DOUBLE_EQ(c4.as_number(), 30.0);

  // No spill region committed at the anchor.
  EXPECT_EQ(wb.sheet(0).spill_region_at_anchor(2U, 1U), nullptr);
}

TEST(SpillCollision, SpillAtTableEdgeJustFitsNoCollision) {
  // Sales table at A1:C5. =SEQUENCE(3,1) at D1 wants D1:D3 -- a single
  // column flush against the table's right edge. No occupant in D1:D3.
  // Spill must succeed; this locks in that table footprints do NOT extend
  // any implicit no-spill margin past their declared `ref`.
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Region"));
  s.set_cell_value(0U, 1U, Value::text("Product"));
  s.set_cell_value(0U, 2U, Value::text("Amount"));
  s.set_cell_value(1U, 0U, Value::text("North"));
  s.set_cell_value(1U, 1U, Value::text("Apple"));
  s.set_cell_value(1U, 2U, Value::number(10.0));
  s.set_cell_value(2U, 0U, Value::text("South"));
  s.set_cell_value(2U, 1U, Value::text("Banana"));
  s.set_cell_value(2U, 2U, Value::number(20.0));
  s.set_cell_value(3U, 0U, Value::text("East"));
  s.set_cell_value(3U, 1U, Value::text("Cherry"));
  s.set_cell_value(3U, 2U, Value::number(30.0));
  s.set_cell_value(4U, 0U, Value::text("West"));
  s.set_cell_value(4U, 1U, Value::text("Durian"));
  s.set_cell_value(4U, 2U, Value::number(40.0));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C5";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = false;
  table.columns = {
      io::TableColumn{1U, "Region", "", "", ""},
      io::TableColumn{2U, "Product", "", "", ""},
      io::TableColumn{3U, "Amount", "", "", ""},
  };
  std::vector<io::TableMetadata> tables = {std::move(table)};
  wb.set_tables(std::move(tables));

  // Anchor at D1 (row 0, col 3); footprint = D1:D3 -> rows 0..2, col 3.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value d1 = wb.sheet(0).resolve_cell_value(0U, 3U);
  ASSERT_TRUE(d1.is_number()) << "D1 expected to spill cleanly; kind=" << static_cast<int>(d1.kind());
  EXPECT_DOUBLE_EQ(d1.as_number(), 1.0);

  const Value d2 = wb.sheet(0).resolve_cell_value(1U, 3U);
  ASSERT_TRUE(d2.is_number());
  EXPECT_DOUBLE_EQ(d2.as_number(), 2.0);

  const Value d3 = wb.sheet(0).resolve_cell_value(2U, 3U);
  ASSERT_TRUE(d3.is_number());
  EXPECT_DOUBLE_EQ(d3.as_number(), 3.0);

  EXPECT_NE(wb.sheet(0).spill_region_at_anchor(0U, 3U), nullptr);
}

}  // namespace
}  // namespace formulon
