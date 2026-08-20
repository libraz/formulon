//
// Unit tests for the workbook recalc engine. The tests drive the engine
// indirectly through `Workbook::set_cell_*` / `Workbook::recalc` so the
// public mutation API and the engine stay in sync.

#include "eval/recalc_engine.h"

#include <algorithm>
#include <cstdint>
#include <string>

#include "cell.h"
#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "eval/scheduler.h"
#include "eval/volatile_tracker.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Returns the cached value at `(row, col)` on `sheet_index`, or
// `Value::blank()` when the cell is absent. Test-only convenience.
Value CellValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& sheet = wb.sheet(sheet_index);
  if (const Cell* c = sheet.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

TEST(RecalcEngine, SimpleLinearChainEvaluatesInOrder) {
  Workbook wb = Workbook::create();
  // A1 = 1, A2 = =A1+1. After recalc, A2 should hold 2.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1+1")));

  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  // A2 was the only formula; one evaluation expected. (No volatile cells.)
  EXPECT_EQ(stats.value().cells_evaluated, 1u);
  EXPECT_EQ(stats.value().volatile_cells, 0u);
  EXPECT_EQ(stats.value().cycle_cells, 0u);

  Value v = CellValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(v.is_number()) << "A2 expected to be a number";
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(RecalcEngine, UpstreamLiteralChangePropagates) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));

  // Now mutate A1 -> 2; A2 must re-evaluate to 3 on the next pass.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));
  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  // Only A2 is a formula cell; the literal A1 is not counted in
  // cells_evaluated (it is not evaluated, only re-stored).
  EXPECT_EQ(stats.value().cells_evaluated, 1u);

  Value v = CellValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(RecalcEngine, WholeColumnDependencyTracksNewValuesAndFormulas) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A:A)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 1.0);

  // A2 did not exist when B1 was registered. Its write must still dirty B1
  // without promoting B1 to the real volatile set.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(2.0))));
  auto value_stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(value_stats));
  EXPECT_EQ(value_stats.value().volatile_cells, 0U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 3.0);

  // A3 is a newly added formula inside the range. Registration creates an
  // ordering edge so B1 observes its fresh value in the same recalc pass.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 0U, "=A1+A2")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 2U, 0U).as_number(), 3.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 6.0);
}

// Drives the ordering contract for a whole-column reference — registered as
// one compact rectangle spanning every row — in one of the two registration
// orders. `D1` is the literal both the interior formula and the aggregate
// ultimately read. Registering the aggregate first is the case where the
// ordering edge has to be acquired after the fact, when the interior formula
// appears inside a rectangle that is already being watched.
void AssertWholeColumnOrdering(bool aggregate_first) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 3U, Value::number(10.0))));  // D1

  const auto write_aggregate = [&wb]() {
    return static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A:A)"));  // B1
  };
  const auto write_interior = [&wb]() {
    return static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=D1+1"));  // A1
  };
  if (aggregate_first) {
    ASSERT_TRUE(write_aggregate());
    ASSERT_TRUE(write_interior());
  } else {
    ASSERT_TRUE(write_interior());
    ASSERT_TRUE(write_aggregate());
  }

  auto initial = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(initial));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 0U).as_number(), 11.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 11.0);

  const auto dependencies = wb.recalc_engine().dep_graph().dependencies_of(CellNodeId{0U, 0U, 1U});
  EXPECT_NE(std::find(dependencies.begin(), dependencies.end(), CellNodeId{0U, 0U, 0U}), dependencies.end());

  // One pass must refresh A1 before B1 reads it; a missing ordering edge
  // leaves B1 on the previous value for a whole recalc.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 3U, Value::number(20.0))));
  auto updated = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(updated));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 0U).as_number(), 21.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 21.0);
}

TEST(RecalcEngine, WholeColumnOrderingEdgeRefreshesFormulaBeforeAggregateInBothRegistrationOrders) {
  AssertWholeColumnOrdering(/*aggregate_first=*/false);
  AssertWholeColumnOrdering(/*aggregate_first=*/true);
}

TEST(RecalcEngine, MultiColumnWholeRangeDirtiesOnFarColumnWrite) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 999U, 2U, Value::number(4.0))));  // C1000
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=SUM(A:C)")));         // D1

  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 3U).as_number(), 4.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 999U, 2U, Value::number(9.0))));
  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 1U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 3U).as_number(), 9.0);
}

TEST(RecalcEngine, WholeRowDependencyTracksNewValues) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SUM(1:1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 2U, Value::number(4.0))));
  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().volatile_cells, 0U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 5.0);
}

// Drives the ordering contract for a bounded rectangle wide enough to be
// registered compactly, in one of the two registration orders. `D1` is the
// literal both the interior formula and the aggregate ultimately read.
void AssertLargeRangeOrdering(bool aggregate_first) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 3U, Value::number(10.0))));  // D1

  const auto write_aggregate = [&wb]() {
    return static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A60000)"));  // B1
  };
  const auto write_interior = [&wb]() {
    return static_cast<bool>(wb.set_cell_formula(0U, 5000U, 0U, "=D1+1"));  // A5001
  };
  if (aggregate_first) {
    ASSERT_TRUE(write_aggregate());
    ASSERT_TRUE(write_interior());
  } else {
    ASSERT_TRUE(write_interior());
    ASSERT_TRUE(write_aggregate());
  }

  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 5000U, 0U).as_number(), 11.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 11.0);

  // The ordering edge exists in the graph even though the rectangle's cells
  // do not: Tarjan needs it to refresh A5001 before B1 in one pass.
  const auto dependencies = wb.recalc_engine().dep_graph().dependencies_of(CellNodeId{0U, 0U, 1U});
  EXPECT_NE(std::find(dependencies.begin(), dependencies.end(), CellNodeId{0U, 5000U, 0U}), dependencies.end());

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 3U, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 5000U, 0U).as_number(), 21.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 21.0);
}

TEST(RecalcEngine, LargeRangeAggregateRecalcsWhenAnInteriorLiteralChanges) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));   // A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A60000)")));  // B1
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 1.0);

  // A40001 did not exist when B1 was registered, and the rectangle owns no
  // per-cell edge to it. The compact watcher must still wake B1 — without
  // making it volatile.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 40000U, 0U, Value::number(5.0))));
  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().volatile_cells, 0U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 6.0);
}

TEST(RecalcEngine, LargeRangeAggregateOrdersAgainstInteriorFormulaInBothRegistrationOrders) {
  AssertLargeRangeOrdering(/*aggregate_first=*/true);
  AssertLargeRangeOrdering(/*aggregate_first=*/false);
}

TEST(RecalcEngine, LargeRangeAggregateDropsWatcherWhenOwnerBecomesLiteral) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));   // A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A60000)")));  // B1
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 1.0);

  // Overwriting the watcher with a literal must retire its rectangle;
  // otherwise every later write inside the range would keep dirtying a cell
  // that no longer reads it.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 40000U, 0U, Value::number(5.0))));
  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 0U);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 99.0);
}

TEST(RecalcEngine, LargeRangeGraphFootprintIsIndependentOfArea) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));      // A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 59999U, 1U, Value::number(2.0))));  // B60000
  const std::size_t before = wb.recalc_engine().dep_graph().node_count();

  // 120,000 cells, none of which may become a graph node: registration cost
  // is bounded by the formula text plus one rectangle.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 3U, "=SUM(A1:B60000)")));  // D1
  EXPECT_EQ(wb.recalc_engine().dep_graph().node_count(), before);

  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 3U).as_number(), 3.0);
  EXPECT_EQ(wb.recalc_engine().dep_graph().node_count(), before);
}

void AssertPhantomOnlySpillSequence(Workbook& wb, bool producer_first) {
  wb.set_excel_profile(mac_365_ja_jp_profile());
  auto install_watcher = [&] { ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)"))); };
  auto install_producer = [&] { ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2)"))); };
  if (producer_first) {
    install_producer();
    install_watcher();
  } else {
    install_watcher();
    install_producer();
  }

  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 12.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2,10)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 39.0);

  // Replacing the producer with a scalar formula that reads the aggregate
  // must clear the old spill-derived edge before SCC construction. The
  // aggregate sees an empty B column and the producer then reads its fresh
  // zero, without a false #REF! cycle.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=C1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_FALSE(wb.recalc_engine().dep_graph().has_dependency_source(CellNodeId{0U, 0U, 2U}, CellNodeId{0U, 1U, 0U},
                                                                    DepGraph::DependencySource::kSpillFootprint));
}

TEST(RecalcEngine, PhantomOnlySpillRangeWatcherWorksWatcherFirst) {
  Workbook wb = Workbook::create();
  AssertPhantomOnlySpillSequence(wb, false);
}

TEST(RecalcEngine, PhantomOnlySpillRangeWatcherWorksProducerFirst) {
  Workbook wb = Workbook::create();
  AssertPhantomOnlySpillSequence(wb, true);
}

TEST(RecalcEngine, CrossSheetPhantomOnlySpillRangeWatcherUpdates) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  wb.add_sheet("Producer");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SUM(Producer!B:B)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 1U, 0U, "=SEQUENCE(3,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 0U).as_number(), 12.0);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 1U, 0U, "=SEQUENCE(3,2,10)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 0U).as_number(), 39.0);
}

TEST(RecalcEngine, PhantomOnlySpillWholeRowWatcherUpdates) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  // A2 is outside row 3, while the second row of its 2x3 spill contributes
  // three phantom-only cells to the compact whole-row dependency.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(2,3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 3U, 3U, "=SUM(3:3)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 3U, 3U).as_number(), 15.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(2,3,10)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 3U, 3U).as_number(), 42.0);
}

TEST(RecalcEngine, UnregisteringSpillProducerDirtiesRangeWatcherBeforeEdgeRemoval) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=SEQUENCE(3,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "bad formula")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
}

TEST(RecalcEngine, DefinedNameReindexClearsStaleSpillFootprint) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MAKE", "SEQUENCE(3,2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=MAKE")));      // A2
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(B:B)")));  // C1

  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 12.0);

  // Retargeting the name from an array producer to a scalar formula must
  // invalidate its old committed spill before graph re-registration. The
  // aggregate then sees an empty B column, while A2 observes C1's fresh zero.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("MAKE", "C1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 2U).as_number(), 0.0);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 1U, 0U).as_number(), 1.0);
  EXPECT_TRUE(wb.sheet(0).committed_spill_footprints().empty());
}

TEST(RecalcEngine, IterationLimitRetainsLastApproximation) {
  Workbook wb = Workbook::create();
  // A1 converges to 2 but cannot do so in one sweep. The engine must retain
  // that first finite approximation instead of overwriting it with #NUM!.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=(A1+2)/2")));
  IterativeOptions options;
  options.enabled = true;
  options.max_iterations = 1U;
  options.max_change = 1e-12;
  wb.set_iterative_options(options);

  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cycle_cells, 1U);
  const Value value = CellValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(value.is_number());
  EXPECT_DOUBLE_EQ(value.as_number(), 1.0);
}

TEST(RecalcEngine, DirectCycleProducesRefError) {
  Workbook wb = Workbook::create();
  // A1 = =A2, A2 = =A1. Both cells participate in a 2-element cycle SCC
  // and should receive #REF!.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=A2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1")));

  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cycle_cells, 2u);
  EXPECT_EQ(stats.value().cells_evaluated, 0u);

  Value a1 = CellValue(wb, 0U, 0U, 0U);
  Value a2 = CellValue(wb, 0U, 1U, 0U);
  ASSERT_TRUE(a1.is_error());
  EXPECT_EQ(a1.as_error(), ErrorCode::Ref);
  ASSERT_TRUE(a2.is_error());
  EXPECT_EQ(a2.as_error(), ErrorCode::Ref);
}

TEST(RecalcEngine, VolatileCellReexecutesEveryRecalc) {
  Workbook wb = Workbook::create();
  // NOW() is a volatile function: every recalc pass should re-execute
  // the cell, regardless of upstream changes. We do not check the cell's
  // value (NOW returns the current time), only that the engine reports
  // the volatile cell was evaluated.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=NOW()")));
  auto stats1 = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats1));
  EXPECT_EQ(stats1.value().volatile_cells, 1u);
  EXPECT_EQ(stats1.value().cells_evaluated, 1u);

  // Second recalc with no mutations: NOW() must still re-execute.
  auto stats2 = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats2));
  EXPECT_EQ(stats2.value().volatile_cells, 1u);
  EXPECT_EQ(stats2.value().cells_evaluated, 1u);
}

TEST(RecalcEngine, RegistrationRecordsTheVolatilityClass) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=RAND()")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=INDIRECT(\"B1\")")));
  // Both classes in one formula: the dynamic read is the binding one.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 2U, 0U, "=RAND()+OFFSET(B1,1,0)")));

  const VolatileTracker& volatiles = wb.recalc_engine().volatiles();
  EXPECT_EQ(volatiles.size(), 3u);
  EXPECT_TRUE(volatiles.contains(CellNodeId{0U, 0U, 0U}));
  EXPECT_FALSE(volatiles.contains_dynamic_reference(CellNodeId{0U, 0U, 0U}));
  EXPECT_TRUE(volatiles.contains_dynamic_reference(CellNodeId{0U, 1U, 0U}));
  EXPECT_TRUE(volatiles.contains_dynamic_reference(CellNodeId{0U, 2U, 0U}));
}

TEST(RecalcEngine, RewritingAwayADynamicReferenceDropsTheClass) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=INDIRECT(\"B1\")")));
  ASSERT_TRUE(wb.recalc_engine().volatiles().contains_dynamic_reference(CellNodeId{0U, 0U, 0U}));

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=RAND()")));
  EXPECT_TRUE(wb.recalc_engine().volatiles().contains(CellNodeId{0U, 0U, 0U}));
  EXPECT_FALSE(wb.recalc_engine().volatiles().contains_dynamic_reference(CellNodeId{0U, 0U, 0U}));

  // A rewrite to a non-volatile formula drops the cell entirely.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=1+1")));
  EXPECT_FALSE(wb.recalc_engine().volatiles().contains(CellNodeId{0U, 0U, 0U}));
}

TEST(RecalcEngine, CrossSheetReferencePropagatesValue) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");

  // Sheet2!A1 = 42; Sheet1!A1 = =Sheet2!A1.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Sheet2!A1")));

  auto stats = wb.recalc(default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cells_evaluated, 1u);

  Value v = CellValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);

  // Mutate the upstream cross-sheet cell; the local formula must recompute.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(99.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));

  v = CellValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 99.0);
}

TEST(RecalcEngine, DepGraphAccessibleForInspection) {
  // Lightweight sanity: the engine exposes its dep graph via
  // `recalc_engine().dep_graph()` and that graph reports the registered
  // edge after `set_cell_formula`.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1+1")));

  const DepGraph& graph = wb.recalc_engine().dep_graph();
  // A2 (row=1, col=0) reads A1 (row=0, col=0).
  std::vector<CellNodeId> deps = graph.dependencies_of(CellNodeId{0U, 1U, 0U});
  ASSERT_EQ(deps.size(), 1u);
  EXPECT_EQ(deps[0], (CellNodeId{0U, 0U, 0U}));
}

TEST(RecalcEngine, FormulaUpdateRewritesDependencies) {
  // Re-registering a formula must drop the old edges. Start with =A1+1, then
  // overwrite with =B1+1 and verify the dep graph no longer points at A1.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=A1+1")));
  {
    const DepGraph& g = wb.recalc_engine().dep_graph();
    auto deps = g.dependencies_of(CellNodeId{0U, 1U, 0U});
    ASSERT_EQ(deps.size(), 1u);
    EXPECT_EQ(deps[0], (CellNodeId{0U, 0U, 0U}));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, "=B1+1")));
  {
    const DepGraph& g = wb.recalc_engine().dep_graph();
    auto deps = g.dependencies_of(CellNodeId{0U, 1U, 0U});
    ASSERT_EQ(deps.size(), 1u);
    EXPECT_EQ(deps[0], (CellNodeId{0U, 0U, 1U}));  // B1 == col 1
  }
}

// ---------------------------------------------------------------------------
// Parallel scheduler contracts shared with the serial engine
// ---------------------------------------------------------------------------

TEST(SchedulerContract, ArenaExhaustionAbortsThePassWithOutOfMemory) {
  // An allocation failure during evaluation is a resource error, not a
  // formula result: the pass must surface `kOutOfMemory` instead of
  // committing the degraded value, exactly as `RecalcEngine::recalc` does.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1000)")));

  SchedulerConfig tiny;
  tiny.num_threads = 1U;
  tiny.max_arena_bytes = 4096U;  // far below the 1,000-cell result
  SchedulerStats tiny_stats;
  auto exhausted = recalc_parallel(wb, tiny, &tiny_stats);
  ASSERT_FALSE(static_cast<bool>(exhausted)) << "a pass that ran out of arena must not report success";
  EXPECT_EQ(exhausted.error().code, FormulonErrorCode::kOutOfMemory);

  // The same workbook under the default ceiling completes and commits.
  SchedulerConfig roomy;
  roomy.num_threads = 1U;
  SchedulerStats roomy_stats;
  ASSERT_TRUE(static_cast<bool>(recalc_parallel(wb, roomy, &roomy_stats)));
  EXPECT_EQ(roomy_stats.cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 0U).as_number(), 1.0);
}

TEST(SchedulerContract, SccPhaseWalksOnlyTheDirtySubgraph) {
  // The SCC phase is scoped to the dirty induced subgraph, like the serial
  // engine's `tarjan_scc_subset` call. A one-cell edit to a workbook with
  // many registered formulas must submit one node to Tarjan, not the whole
  // registered graph.
  Workbook wb = Workbook::create();
  constexpr std::uint32_t kFormulaCount = 200U;
  for (std::uint32_t row = 0; row < kFormulaCount; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(static_cast<double>(row)))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=A" + std::to_string(row + 1U) + "+1")));
  }
  SchedulerConfig cfg;
  cfg.num_threads = 1U;
  // The first pass is dirty everywhere: each literal and each formula.
  SchedulerStats warmup;
  ASSERT_TRUE(static_cast<bool>(recalc_parallel(wb, cfg, &warmup)));
  EXPECT_EQ(warmup.scc_nodes_considered, 2U * kFormulaCount);

  // Touch one input: the dirty closure is that literal plus the one formula
  // reading it, whatever the registered graph's size.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(42.0))));
  SchedulerStats incremental;
  ASSERT_TRUE(static_cast<bool>(recalc_parallel(wb, cfg, &incremental)));
  EXPECT_EQ(incremental.scc_nodes_considered, 2u) << "Tarjan must not walk the whole registered graph";
  EXPECT_EQ(incremental.cells_evaluated, 1u);
  EXPECT_DOUBLE_EQ(CellValue(wb, 0U, 0U, 1U).as_number(), 43.0);
}

}  // namespace
}  // namespace formulon::eval
