//
// Unit tests for `RecalcEngine::partial_recalc`. The tests drive the
// engine through the workbook public API so the dep graph state is
// always populated through the same mutation paths a real client uses.

#include <cstdint>
#include <string>

#include "cell.h"
#include "eval/array_alloc.h"
#include "eval/builtins.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
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

Value CustomSpill(const Value* /*args*/, std::uint32_t /*arity*/, Arena& arena) {
  Value* cells = nullptr;
  ArrayValue* array = allocate_array_value(1U, 2U, arena, cells);
  if (array == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  cells[0] = Value::number(7.0);
  cells[1] = Value::number(8.0);
  return Value::array(array);
}

Value CustomArrayAbs(const Value* /*args*/, std::uint32_t /*arity*/, Arena& arena) {
  Value* cells = nullptr;
  ArrayValue* array = allocate_array_value(1U, 2U, arena, cells);
  if (array == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  cells[0] = Value::number(7.0);
  cells[1] = Value::number(8.0);
  return Value::array(array);
}

Value CustomSum(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  double sum = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (args[i].is_number()) {
      sum += args[i].as_number();
    }
  }
  return Value::number(sum);
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

TEST(PartialRecalc, OversizedViewportIsNoop) {
  // A near-full-grid viewport would seed billions of coordinates before
  // any dependency work. The engine treats it like the other invalid
  // shapes: no-op, dirty set preserved, so callers fall back to recalc().
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1+1")));

  SheetCellRange full_grid;
  full_grid.sheet_id = 0;
  full_grid.first_row = 0;
  full_grid.last_row = Sheet::kMaxRows - 1U;
  full_grid.first_col = 0;
  full_grid.last_col = Sheet::kMaxCols - 1U;

  auto stats = wb.partial_recalc(default_registry(), full_grid);
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 0u);
  // A2 was never evaluated and stays dirty for the fallback full recalc.
  EXPECT_TRUE(CellValue(wb, 0U, 1U, 0U).is_blank());
  auto full = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(full));
  EXPECT_EQ(full.value().cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 2.0);
}

TEST(PartialRecalc, SingleCellViewportRecomputesAncestorsOnly) {
  // A1 = 10, B1 = =A1+1 (not in viewport), C1 = =A1*2 (not in
  // viewport). Viewport asks for B1 only — only A1 (literal) and B1
  // are in the closure. C1 should remain unevaluated this pass.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));  // B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=A1*2")));  // C1

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
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1")));      // B1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=A2")));      // B2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=B1+1")));    // C1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 2U, "=B2+1")));    // C2
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

TEST(PartialRecalc, SpillReleaseRetriesBlockedAnchorInClosureNextWave) {
  // B1 owns the committed B1:C3 region and A2 owns a blocked A2:B4
  // footprint. Shrinking B1:C3 to B1:C1 releases A2:B4 while the viewport
  // includes both producers, so the release must be retried in a second,
  // dependency-ordered partial wave rather than lost at pass end.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 4U, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SEQUENCE(E1,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_EQ(CellValue(wb, 0U, 1U, 0U).as_error(), ErrorCode::Spill);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 4U, Value::number(1.0))));
  SheetCellRange viewport;
  viewport.sheet_id = 0U;
  viewport.first_row = 0U;
  viewport.last_row = 1U;
  viewport.first_col = 0U;
  viewport.last_col = 1U;
  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), viewport)));

  const Value a2 = CellValue(wb, 0U, 1U, 0U);
  const Value b4 = wb.sheet(0).resolve_cell_value(3U, 1U);
  ASSERT_TRUE(a2.is_number()) << "A2 kind=" << static_cast<int>(a2.kind());
  ASSERT_TRUE(b4.is_number()) << "B4 kind=" << static_cast<int>(b4.kind());
  EXPECT_DOUBLE_EQ(a2.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(b4.as_number(), 6.0);
  EXPECT_TRUE(wb.sheet(0).blocked_spill_anchors().empty());
}

TEST(PartialRecalc, InitialPhantomOnlySpillDependencyAndRemoteDirtyContract) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  wb.add_sheet("Remote");

  // The watcher is registered first and reads compact B:B. The producer's
  // anchor is A2, outside that range, while its committed phantom values
  // land in B2:B4. There is no derived edge before the first spill commit.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2)")));

  // Keep unrelated remote work dirty; a partial pass over Sheet1 must not
  // force or clear it.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 0U, "=NOW()")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 1U, "=42")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(0U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 12.0);
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{1U, 0U, 0U}));
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{1U, 0U, 1U}));

  const CellNodeId watcher{0U, 0U, 2U};
  const CellNodeId producer{0U, 1U, 0U};
  EXPECT_TRUE(wb.recalc_engine().dep_graph().has_dependency_source(watcher, producer,
                                                                   DepGraph::DependencySource::kSpillFootprint));

  // A changed producer is reached through the derived edge and refreshes the
  // aggregate in the same partial call.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2,10)")));
  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(0U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 39.0);

  // The producer may be rewritten to depend on the aggregate itself. The
  // partial wave must remove the stale spill edge before SCC construction and
  // evaluate the scalar dependency order in this same call.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=C1+1")));
  SheetCellRange c1_a2;
  c1_a2.sheet_id = 0U;
  c1_a2.first_row = 0U;
  c1_a2.last_row = 1U;
  c1_a2.first_col = 0U;
  c1_a2.last_col = 2U;
  auto rewrite_stats = wb.partial_recalc(default_registry(), c1_a2);
  ASSERT_TRUE(static_cast<bool>(rewrite_stats));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);

  // A second overlapping partial call has no dirty or volatile formulas left
  // in the closure and must preserve the scalar result without reevaluation.
  auto repeat_stats = wb.partial_recalc(default_registry(), c1_a2);
  ASSERT_TRUE(static_cast<bool>(repeat_stats));
  EXPECT_EQ(repeat_stats.value().cells_evaluated, 0U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);

  // Removing the spill leaves B:B empty. The old derived edge is active for
  // this wave, so the watcher sees the shrink before the edge is removed.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=1")));
  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(0U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_FALSE(wb.recalc_engine().dep_graph().has_dependency_source(watcher, producer,
                                                                    DepGraph::DependencySource::kSpillFootprint));
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{1U, 0U, 0U}));
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{1U, 0U, 1U}));
}

TEST(PartialRecalc, DefinedNameReindexClearsStaleSpillInViewport) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MAKE", "SEQUENCE(3,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=MAKE")));      // A2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));  // C1

  SheetCellRange c1_a2;
  c1_a2.sheet_id = 0U;
  c1_a2.first_row = 0U;
  c1_a2.last_row = 1U;
  c1_a2.first_col = 0U;
  c1_a2.last_col = 2U;
  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), c1_a2)));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 12.0);

  // Reindexing must invalidate the committed spill before rebuilding the
  // graph. The partial closure then evaluates C1 before A2=C1+1 without a
  // stale phantom-derived SCC.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MAKE", "C1+1")));
  auto rewrite_stats = wb.partial_recalc(default_registry(), c1_a2);
  ASSERT_TRUE(static_cast<bool>(rewrite_stats));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);
  EXPECT_TRUE(wb.sheet(0).committed_spill_footprints().empty());

  auto repeat_stats = wb.partial_recalc(default_registry(), c1_a2);
  ASSERT_TRUE(static_cast<bool>(repeat_stats));
  EXPECT_EQ(repeat_stats.value().cells_evaluated, 0U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);
}

TEST(PartialRecalc, ModeMultSpillReachesWholeRowWatcher) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());

  // MODE.MULT returns the tied modes as a vertical array. Its second value
  // lands in A100, where the row-100 watcher must observe it even though
  // the producer anchor is A99 and there is no initial derived edge. Keep
  // the watcher on row 101 so SUM(100:100) does not self-reference.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 98U, 0U, "=MODE.MULT(1,2,1,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 100U, 2U, "=SUM(100:100)")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(100U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 100U, 2U).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(98U, 0U).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(99U, 0U).as_number(), 2.0);
}

TEST(PartialRecalc, CustomPotentialProducerUsesDownRightGeometry) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  FunctionRegistry registry;
  register_builtins(registry);
  ASSERT_TRUE(registry.register_function(FunctionDef{"CUSTOMSPILL", 0U, kVariadic, &CustomSpill}));

  // C1 watches B:B. A100 can spill right into B100 despite its anchor being
  // outside the range; Z100 starts too far right and must remain untouched.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 0U, "=CUSTOMSPILL()")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 25U, "=CUSTOMSPILL()")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(registry, SingleCell(0U, 2U))));
  const Value aggregate = CellValue(wb, 0U, 0U, 2U);
  ASSERT_TRUE(aggregate.is_number()) << "aggregate kind=" << static_cast<int>(aggregate.kind());
  EXPECT_DOUBLE_EQ(aggregate.as_number(), 8.0);
  const Value phantom = wb.sheet(0).resolve_cell_value(99U, 1U);
  ASSERT_TRUE(phantom.is_number()) << "phantom kind=" << static_cast<int>(phantom.kind());
  EXPECT_DOUBLE_EQ(phantom.as_number(), 8.0);
  EXPECT_TRUE(CellValue(wb, 0U, 99U, 25U).is_blank());
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{0U, 99U, 25U}));
}

TEST(PartialRecalc, BoundedCompactWatcherReachesInteriorWorkOutsideViewport) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  FunctionRegistry registry;
  register_builtins(registry);
  ASSERT_TRUE(registry.register_function(FunctionDef{"CUSTOMSPILL", 0U, kVariadic, &CustomSpill}));

  // C1 watches a bounded rectangle wide enough to be registered compactly,
  // so it owns no per-cell edge into the rectangle. Neither the interior
  // formula (B2) nor the spill producer (A100, which spills right into B100)
  // is inside the viewport, and both must still be evaluated before C1 can
  // be correct.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B1:B60000)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 3U, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=D1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 0U, "=CUSTOMSPILL()")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(registry, SingleCell(0U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 1U).as_number(), 4.0);
  const Value phantom = wb.sheet(0).resolve_cell_value(99U, 1U);
  ASSERT_TRUE(phantom.is_number()) << "phantom kind=" << static_cast<int>(phantom.kind());
  EXPECT_DOUBLE_EQ(phantom.as_number(), 8.0);
  const Value aggregate = CellValue(wb, 0U, 0U, 2U);
  ASSERT_TRUE(aggregate.is_number()) << "aggregate kind=" << static_cast<int>(aggregate.kind());
  EXPECT_DOUBLE_EQ(aggregate.as_number(), 12.0);
}

TEST(PartialRecalc, RuntimeRegistryReclassifiesKnownEagerName) {
  // The potential index is built without the caller's registry. A host may
  // intentionally override an eager built-in name with an array-producing
  // UDF, so the partial candidate must be classified again at admission time.
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  FunctionRegistry registry;
  FunctionDef sum_def{"SUM", 1U, kVariadic, &CustomSum};
  sum_def.accepts_ranges = true;
  sum_def.result_shape = FunctionDef::ResultShape::kReduce;
  ASSERT_TRUE(registry.register_function(sum_def));
  FunctionDef abs_def{"ABS", 1U, 1U, &CustomArrayAbs};
  abs_def.result_shape = FunctionDef::ResultShape::kArray;
  ASSERT_TRUE(registry.register_function(abs_def));

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=ABS(1)")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(registry, SingleCell(0U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 8.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(1U, 1U).as_number(), 8.0);
}

TEST(PartialRecalc, RuntimeRegistryCanExcludeScalarIntrinsicOverride) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  FunctionRegistry registry;
  FunctionDef sum_def{"SUM", 1U, kVariadic, &CustomSum};
  sum_def.accepts_ranges = true;
  sum_def.result_shape = FunctionDef::ResultShape::kReduce;
  ASSERT_TRUE(registry.register_function(sum_def));

  // The host deliberately replaces SEQUENCE with a scalar implementation.
  // A potential index built without this runtime registry must retain the
  // formula as registry-sensitive, then reject it during partial admission.
  FunctionDef scalar_sequence{"SEQUENCE", 1U, kVariadic, &CustomSum};
  scalar_sequence.result_shape = FunctionDef::ResultShape::kScalar;
  ASSERT_TRUE(registry.register_function(scalar_sequence));

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(1,2)")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(registry, SingleCell(0U, 2U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_TRUE(CellValue(wb, 0U, 1U, 0U).is_blank());
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{0U, 1U, 0U}));
}

TEST(PartialRecalc, CleanIntermediateSpillChainCompletesInOneCall) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());

  // C100 is initially clean and scalar because B:B is empty. E100 is also
  // clean. A later A100 write must discover the clean intermediate C100 from
  // E100's range watcher and complete both derived links in one call.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 2U, "=IF(SUM(B:B)>0,SEQUENCE(1,2),0)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 4U, "=SUM(D:D)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 99U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 99U, 4U).as_number(), 0.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 0U, "=SEQUENCE(1,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 25U, "=40+2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 25U, "=SEQUENCE(1,2)")));

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(99U, 4U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 99U, 2U).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(99U, 3U).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 99U, 4U).as_number(), 2.0);

  // Neither unrelated formula is in the E100 closure. Both retain their
  // blank cache and dirty state for a later full/overlapping pass.
  EXPECT_TRUE(CellValue(wb, 0U, 0U, 25U).is_blank());
  EXPECT_TRUE(CellValue(wb, 0U, 99U, 25U).is_blank());
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{0U, 0U, 25U}));
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{0U, 99U, 25U}));
}

TEST(PartialRecalc, WholeRowSpillGeometryPreservesUnrelatedDirtyCells) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 3U, 3U, "=SUM(3:3)")));        // D4
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(2,3)")));   // A2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 25U, "=40+2")));           // Z1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 99U, 0U, "=SEQUENCE(1,2)")));  // A100

  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(3U, 3U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 3U, 3U).as_number(), 15.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(2U, 0U).as_number(), 4.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(2U, 2U).as_number(), 6.0);

  EXPECT_TRUE(CellValue(wb, 0U, 0U, 25U).is_blank());
  EXPECT_TRUE(CellValue(wb, 0U, 99U, 0U).is_blank());
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{0U, 0U, 25U}));
  EXPECT_TRUE(wb.recalc_engine().dirty().contains(CellNodeId{0U, 99U, 0U}));
}

TEST(PartialRecalc, MultiStageSpillChainCompletesInOneCall) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(1.0))));           // A2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=IF(A2,SEQUENCE(1,2),0)")));  // C1 -> D1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 4U, "=SUM(D:D)")));                // E1

  // E1 initially has no committed edge to C1. Candidate closure expansion
  // admits C1, evaluates its array, reconciles E1->C1, and retries E1 before
  // this one partial call returns.
  ASSERT_TRUE(static_cast<bool>(wb.partial_recalc(default_registry(), SingleCell(0U, 4U))));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 4U).as_number(), 2.0);
  EXPECT_DOUBLE_EQ(wb.sheet(0).resolve_cell_value(0U, 3U).as_number(), 2.0);
}

}  // namespace
}  // namespace formulon::eval
