// Concurrency scheduler tests grouped by fixture size and execution mode.
// Shared probes and workbook-pair helpers live in scheduler_test_helpers.h.

#include "scheduler_test_helpers.h"

namespace formulon::eval {
namespace {

using namespace scheduler_test;

TEST(Scheduler, EmptyWorkbookProducesZeroStats) {
  Workbook wb = Workbook::create();
  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  auto result = wb.recalc_parallel(default_registry(), cfg, &stats);
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_EQ(stats.cells_evaluated, 0U);
  EXPECT_EQ(stats.sccs_processed, 0U);
  EXPECT_EQ(stats.parallel_steps, 0U);
  EXPECT_EQ(stats.serial_fallback_steps, 0U);
  EXPECT_EQ(stats.cycle_recoveries, 0U);
}

TEST(Scheduler, SingleConstantCellSingleThread) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(42.0))));
  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 1U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cells_evaluated, 0U);  // No formula to execute.
  EXPECT_EQ(stats.worker_threads_started, 0U);
  EXPECT_EQ(stats.worker_threads_used, 0U);
  // A literal mark may seed the dirty set but nothing in the SCC graph;
  // the standalone-dirty sweep will skip it because formula_text is empty.
}

TEST(Scheduler, SingleConstantCellEightThreads) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(42.0))));
  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cells_evaluated, 0U);
}

TEST(Scheduler, ParallelIterativeProgressCallbackCanAbort) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=(A1+1000)/2")));

  IterativeOptions options;
  options.enabled = true;
  options.max_iterations = 100U;
  options.max_change = 1e-12;
  wb.set_iterative_options(options);

  ProgressAbortAfter progress;
  progress.abort_after = 3U;
  wb.recalc_engine().set_iterative_progress(&AbortIterativeSolve, &progress);

  SchedulerConfig config;
  config.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), config, nullptr)));
  EXPECT_EQ(progress.calls.load(std::memory_order_relaxed), 3U);

  // Aborting is a "stop here", not a failure: the cell keeps whatever the
  // last completed pass committed, exactly as iteration-limit exhaustion
  // does. Three halvings toward the fixed point 1000 give 500, 750, 875.
  const Value value = StoredValue(wb, 0U, 0U, 0U);
  ASSERT_TRUE(value.is_number()) << "an aborted solve must retain the last approximation";
  EXPECT_DOUBLE_EQ(value.as_number(), 875.0);
}

TEST(Scheduler, IterativeProgressCallbackRunsOnCallerThread) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=(A1+1000)/2")));
  // A second, disjoint self-cycle makes the cyclic layer genuinely wide.
  // Before callback affinity was enforced, the scheduler dispatched these
  // two SCCs to separate workers and invoked the progress callback there.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=(C1+2000)/2")));

  IterativeOptions options;
  options.enabled = true;
  options.max_iterations = 100U;
  options.max_change = 1e-6;
  wb.set_iterative_options(options);

  ProgressThreadProbe probe;
  probe.caller = std::this_thread::get_id();
  wb.recalc_engine().set_iterative_progress(&RecordIterativeProgressThread, &probe);

  SchedulerConfig config;
  config.num_threads = 4U;
  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), config, &stats)));
  EXPECT_EQ(stats.cycle_recoveries, 2U);
  EXPECT_EQ(stats.cells_evaluated, 2U);
  EXPECT_EQ(stats.parallel_steps, 0U);
  EXPECT_EQ(stats.serial_fallback_steps, 1U);

  {
    std::lock_guard<std::mutex> guard(probe.mutex);
    ASSERT_FALSE(probe.callback_threads.empty());
    for (const std::thread::id callback_thread : probe.callback_threads) {
      EXPECT_EQ(callback_thread, probe.caller);
    }
  }

  const Value a1 = StoredValue(wb, 0U, 0U, 0U);
  const Value c1 = StoredValue(wb, 0U, 0U, 2U);
  ASSERT_TRUE(a1.is_number());
  ASSERT_TRUE(c1.is_number());
  EXPECT_NEAR(a1.as_number(), 1000.0, 1e-4);
  EXPECT_NEAR(c1.as_number(), 2000.0, 1e-4);
}

}  // namespace
}  // namespace formulon::eval
