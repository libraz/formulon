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

TEST(WorkbookRecalc, ThousandCellCumulativeChainUsesDependencyOrderedCaches) {
  Workbook wb = Workbook::create();
  constexpr std::uint32_t kRows = 1000U;
  for (std::uint32_t row = 0; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(1.0))));
    const std::string formula = (row == 0U) ? "=A1" : "=B" + std::to_string(row) + "+A" + std::to_string(row + 1U);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, formula)));
  }

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t row = 0; row < kRows; ++row) {
    const Cell* cell = wb.sheet(0).cell_at(row, 1U);
    ASSERT_NE(cell, nullptr);
    ASSERT_TRUE(cell->cached_value.is_number()) << "row=" << row + 1U;
    EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), static_cast<double>(row + 1U)) << "row=" << row + 1U;
  }
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

TEST(WorkbookRecalc, SpillOffGridEdgeSurfacesSpillError) {
  // A dynamic array anchored near the last row cannot fit its footprint on
  // the sheet; Excel surfaces #SPILL!. The committer must set that error
  // deterministically rather than leave the anchor's prior value in place.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  // Anchor on the final row: =SEQUENCE(3,1) would spill down into two rows
  // that do not exist.
  const std::uint32_t last_row = Sheet::kMaxRows - 1U;
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, last_row, 0U, "=SEQUENCE(3,1)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  Value anchor = wb.sheet(0).resolve_cell_value(last_row, 0U);
  ASSERT_TRUE(anchor.is_error()) << "off-grid spill expected to be #SPILL!";
  EXPECT_EQ(anchor.as_error(), ErrorCode::Spill);
}

TEST(WorkbookRecalc, DirectLambdaCallTracksBodyCellDependency) {
  // A directly-invoked lambda evaluates its body, so a cell ref inside the
  // body (A1) is a real dependency. Editing A1 must re-evaluate the caller.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));      // A1 = 10
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=LAMBDA(x, x+A1)(5)")));  // B1
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 1U).as_number(), 15.0);

  // Change A1; the lambda-call cell must reflect the new value.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(100.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 1U).as_number(), 105.0)
      << "direct lambda-call body dependency on A1 was not tracked";
}

TEST(WorkbookRecalc, RedefiningDefinedNameInvalidatesDependents) {
  // A formula referencing a defined name must re-resolve after the name is
  // retargeted; the stale value from the old definition must not survive.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));  // A1 = 1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(2.0))));  // A2 = 2
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Target", "=A1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=Target+10")));  // B1 = Target+10
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 1U).as_number(), 11.0);

  // Retarget Target from A1 to A2; B1 must recompute to 2 + 10 = 12.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Target", "=A2")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 1U).as_number(), 12.0)
      << "defined-name dependent kept the stale value after retarget";
}

TEST(WorkbookRecalc, AddingDefinedNameRecalculatesNameErrorDependents) {
  // An initially unresolved name produces #NAME?. Adding its definition must
  // rebuild the formula graph so the already-entered formula becomes live.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Rate*10")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(StoredValue(wb, 0U, 0U, 0U).is_error());
  EXPECT_EQ(StoredValue(wb, 0U, 0U, 0U).as_error(), ErrorCode::Name);

  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Rate", "=2")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_number());
  EXPECT_DOUBLE_EQ(a1.as_number(), 20.0);
}

TEST(WorkbookRecalc, FormattingLiveSpillPhantomPreservesDynamicArray) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  // B1 is a live spill phantom. Applying an xf is metadata-only and must
  // retain its value and the complete spill region.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(0U, 0U, 1U, 7U)));
  const Cell* b1 = wb.sheet(0).cell_at(0U, 1U);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(b1->xf_index, 7U);

  const Value a1 = wb.sheet(0).resolve_cell_value(0U, 0U);
  const Value spill_b1 = wb.sheet(0).resolve_cell_value(0U, 1U);
  const Value c1 = wb.sheet(0).resolve_cell_value(0U, 2U);
  ASSERT_TRUE(a1.is_number());
  ASSERT_TRUE(spill_b1.is_number());
  ASSERT_TRUE(c1.is_number());
  EXPECT_DOUBLE_EQ(a1.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(spill_b1.as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c1.as_number(), 3.0);
}

TEST(WorkbookRecalc, WritingIntoLiveSpillPhantomResurfacesSpillError) {
  // A1 spills into A1:A3. Writing a literal into the phantom A2 blocks the
  // footprint; the anchor A1 must re-evaluate to #SPILL! rather than keep
  // its old spilled value.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(eval::mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(3,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.sheet(0).resolve_cell_value(0U, 0U).is_number());  // A1 spilled

  // Block the spill by writing into phantom A2.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value a1 = wb.sheet(0).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(a1.is_error()) << "anchor kept its stale spilled value after a phantom write";
  EXPECT_EQ(a1.as_error(), ErrorCode::Spill);
}

}  // namespace
}  // namespace formulon
