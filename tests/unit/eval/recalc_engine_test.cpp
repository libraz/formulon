//
// Unit tests for the workbook recalc engine. The tests drive the engine
// indirectly through `Workbook::set_cell_*` / `Workbook::recalc` so the
// public mutation API and the engine stay in sync.

#include "eval/recalc_engine.h"

#include <cstdint>
#include <string>

#include "cell.h"
#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
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

}  // namespace
}  // namespace formulon::eval
