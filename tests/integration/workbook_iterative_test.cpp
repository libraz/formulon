//
// Integration tests for the workbook-level iterative-calc surface.
// Drives the public Workbook API (`set_cell_formula`, `set_cell_value`,
// `set_iterative_options`, `recalc`) against real cyclic dep-graph
// configurations and verifies the recalc engine forwards SCCs to the
// iterative solver under the expected conditions:
//
//   * iterative calc disabled (default) -> cyclic SCC surfaces #REF! and
//     bumps `RecalcStats::cycle_cells` (legacy behaviour preserved).
//   * iterative calc enabled, convergent recurrence -> SCC members hold
//     the converged numeric values, `RecalcStats::iterative_cells`
//     reports per-member counts.
//   * iterative calc enabled, divergent recurrence -> SCC members hold
//     #NUM! and `RecalcStats::cycle_cells` accounts for the failure.
//   * `max_iterations` is honoured: a tight cap forces the iteration
//     limit to fire and the engine writes #NUM! on exhaustion.

#include <cmath>
#include <cstdint>
#include <limits>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

TEST(WorkbookIterative, DefaultDisabledStillSurfacesRefForCycle) {
  // Default iterative options: enabled = false. A cycle between A1 and
  // B1 must surface #REF! on both cells, and `RecalcStats::cycle_cells`
  // must report 2; `iterative_cells` stays at 0.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));

  EXPECT_FALSE(wb.iterative_options().enabled);

  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cycle_cells, 2U);
  EXPECT_EQ(stats.value().iterative_cells, 0U);

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(a1.is_error());
  ASSERT_TRUE(b1.is_error());
  EXPECT_EQ(a1.as_error(), ErrorCode::Ref);
  EXPECT_EQ(b1.as_error(), ErrorCode::Ref);
}

TEST(WorkbookIterative, EnabledIdentityCycleObservedBehaviour) {
  // `=A1+1` and `=B1-1`-style "non-fixed-point" identity cycle:
  // A1 = B1 + 1, B1 = A1 - 1. Substituting into either equation yields
  // a tautology (any pair (a, a-1) satisfies both), so there is no
  // unique fixed point — but Gauss-Seidel-style iteration finds
  // *some* pair quickly because each pass commits A then B, and the
  // commit order means B always recomputes from the freshly-committed
  // A. Empirically: pass 1 commits A=1 (B was Blank, treated as 0),
  // then B = A - 1 = 0. Pass 2 commits A = B + 1 = 1, B = A - 1 = 0.
  // Delta is 0 -> converged. We pin this observable shape.
  //
  // Documents the design decision noted in the bundle plan:
  // "non-fixed-point identity cycles are observed and the test pins the
  // outcome".
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1-1")));

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 100U;
  opts.max_change = 0.001;
  wb.set_iterative_options(opts);
  EXPECT_TRUE(wb.iterative_options().enabled);

  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  // The Gauss-Seidel pass converges on the (1, 0) pair.
  EXPECT_EQ(stats.value().iterative_cells, 2U);
  EXPECT_EQ(stats.value().cells_evaluated, 2U);
  EXPECT_EQ(stats.value().cycle_cells, 0U);

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(a1.is_number()) << "A1 expected numeric, got " << (a1.is_error() ? static_cast<int>(a1.as_error()) : -1);
  ASSERT_TRUE(b1.is_number()) << "B1 expected numeric, got " << (b1.is_error() ? static_cast<int>(b1.as_error()) : -1);
  EXPECT_DOUBLE_EQ(a1.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(b1.as_number(), 0.0);
}

TEST(WorkbookIterative, EnabledAveragingCycleConvergesToSharedValue) {
  // A1 = B1 / 2 + 10, B1 = A1.
  // Substituting B1 = A1 into A1 = B1 / 2 + 10 -> A1 = A1 / 2 + 10
  // -> A1 / 2 = 10 -> A1 = 20. The system has a true fixed point
  // (A1, B1) = (20, 20), reached geometrically.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1/2+10")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1")));

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 200U;
  opts.max_change = 0.0001;
  wb.set_iterative_options(opts);

  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().iterative_cells, 2U);
  EXPECT_EQ(stats.value().cycle_cells, 0U);

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(a1.is_number());
  ASSERT_TRUE(b1.is_number());
  EXPECT_NEAR(a1.as_number(), 20.0, 0.001);
  EXPECT_NEAR(b1.as_number(), 20.0, 0.001);
}

TEST(WorkbookIterative, WholeColumnAggregateCycleConvergesAndCountsBothCells) {
  Workbook wb = Workbook::create();
  // A1 = C1 * 0.5 + 1 and C1 = SUM(A:A) form a cycle through a compact
  // whole-column dependency. The ordering edge must preserve the fresh A1
  // value before C1 is evaluated on each Gauss-Seidel sweep.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=C1*0.5+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SUM(A:A)")));

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 200U;
  opts.max_change = 1e-9;
  wb.set_iterative_options(opts);

  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().iterative_cells, 2U);
  EXPECT_EQ(stats.value().cycle_cells, 0U);

  const Value a1 = StoredValue(wb, 0U, 0U, 0U);
  const Value c1 = StoredValue(wb, 0U, 0U, 2U);
  ASSERT_TRUE(a1.is_number());
  ASSERT_TRUE(c1.is_number());
  EXPECT_NEAR(a1.as_number(), 2.0, 1e-6);
  EXPECT_NEAR(c1.as_number(), 2.0, 1e-6);
}

TEST(WorkbookIterative, EnabledGrowingCycleKeepsFiniteApproximation) {
  // A1 = 2 * A1 + 1: a single-cell self-referential SCC whose recurrence
  // grows without bound. Excel has no residual-growth cutoff, so 100
  // iterations from a blank seed simply leave `2^100 - 1` in the cell.
  // Only an actual non-finite result would surface #NUM!.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=2*A1+1")));

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 100U;
  opts.max_change = 0.001;
  wb.set_iterative_options(opts);

  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  // The solve never met `max_change`, so the members are tallied as
  // unresolved rather than iterative.
  EXPECT_EQ(stats.value().cycle_cells, 1U);
  EXPECT_EQ(stats.value().iterative_cells, 0U);

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_number()) << "iteration-limit exhaustion must retain the last approximation";
  EXPECT_DOUBLE_EQ(a1.as_number(), std::pow(2.0, 100.0) - 1.0);
}

TEST(WorkbookIterative, MaxIterationsHonoured) {
  // Convergent-but-slow recurrence: A1 = (A1 + 1000) / 2, fixed point at
  // A1 = 1000. With max_iterations = 3 and a tight max_change the solver
  // cannot converge in time. Excel leaves the last approximation in the
  // cell, so three halvings from a blank seed give 500, 750, 875.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=(A1+1000)/2")));

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 3U;
  opts.max_change = 1e-9;
  wb.set_iterative_options(opts);

  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(stats.value().cycle_cells, 1U);
  EXPECT_EQ(stats.value().iterative_cells, 0U);

  Value a1 = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(a1.is_number()) << "iteration-limit exhaustion must retain the last approximation";
  EXPECT_DOUBLE_EQ(a1.as_number(), 875.0);
}

TEST(WorkbookIterative, OptionsRoundTrip) {
  // Sanity: `set_iterative_options` round-trips through
  // `iterative_options()` and the defaults match the documented Excel
  // values.
  Workbook wb = Workbook::create();
  EXPECT_FALSE(wb.iterative_options().enabled);
  EXPECT_EQ(wb.iterative_options().max_iterations, 100U);
  EXPECT_DOUBLE_EQ(wb.iterative_options().max_change, 0.001);

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 42U;
  opts.max_change = 0.5;
  wb.set_iterative_options(opts);
  EXPECT_TRUE(wb.iterative_options().enabled);
  EXPECT_EQ(wb.iterative_options().max_iterations, 42U);
  EXPECT_DOUBLE_EQ(wb.iterative_options().max_change, 0.5);
}

// The iteration budget is bounded where it enters the model, so every
// reader of `iterative_options()` sees a value the solver can exhaust in
// bounded time — whether or not the path it took clamped on the way in.
TEST(WorkbookIterative, OptionsClampTheIterationBudgetOnTheWayIn) {
  Workbook wb = Workbook::create();

  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_change = 0.0;  // unsatisfiable, so the count is the only bound
  opts.max_iterations = std::numeric_limits<std::uint32_t>::max();
  wb.set_iterative_options(opts);
  EXPECT_EQ(wb.iterative_options().max_iterations, eval::kMaxIterationsCap);

  opts.max_iterations = eval::kMaxIterationsCap + 1U;
  wb.set_iterative_options(opts);
  EXPECT_EQ(wb.iterative_options().max_iterations, eval::kMaxIterationsCap);

  // The cap itself and everything under it are stored verbatim; the clamp
  // must not round a legal request down.
  opts.max_iterations = eval::kMaxIterationsCap;
  wb.set_iterative_options(opts);
  EXPECT_EQ(wb.iterative_options().max_iterations, eval::kMaxIterationsCap);

  opts.max_iterations = 1U;
  wb.set_iterative_options(opts);
  EXPECT_EQ(wb.iterative_options().max_iterations, 1U);

  // The low end is the solver's contract (`0` means one pass), not this
  // setter's, so it is passed through rather than raised to 1 here.
  opts.max_iterations = 0U;
  wb.set_iterative_options(opts);
  EXPECT_EQ(wb.iterative_options().max_iterations, 0U);
}

}  // namespace
}  // namespace formulon
