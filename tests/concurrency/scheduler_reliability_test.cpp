// Concurrency scheduler tests grouped by fixture size and execution mode.
// Shared probes and workbook-pair helpers live in scheduler_test_helpers.h.

#include "scheduler_test_helpers.h"

namespace formulon::eval {
namespace {

using namespace scheduler_test;

// Repeats the wide-independent-layer scenario in a tight loop with the
// maximum auto-detected pool to give ThreadSanitizer the maximum chance
// of catching a racy `next_index` claim. Each iteration must produce
// exactly the same per-cell results — a missed task (e.g. from a
// premature `relaxed` fetch_add allowing a worker to skip a slot) would
// surface as a non-numeric / wrong-numeric cached value.
TEST(SchedulerParallelClaim, WideLayerNoMissedTasksUnderRepeatedRuns) {
  constexpr std::uint32_t kRows = 64U;
  constexpr int kIterations = 16;
  for (int iter = 0; iter < kIterations; ++iter) {
    Workbook wb = Workbook::create();
    for (std::uint32_t r = 0; r < kRows; ++r) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, r, 0U, Value::number(static_cast<double>(r) + 1.0))));
    }
    // Column B: every B<r> reads A<r>. One wide layer of `kRows` super-nodes
    // all draining through the same `next_index` atomic.
    for (std::uint32_t r = 0; r < kRows; ++r) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, r, 1U, "=A" + std::to_string(r + 1) + "+1")));
    }

    SchedulerStats stats;
    SchedulerConfig cfg;
    cfg.num_threads = 8U;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
    // Every formula must have been claimed exactly once.
    EXPECT_EQ(stats.cells_evaluated, static_cast<std::uint64_t>(kRows)) << "iteration " << iter;
    EXPECT_GE(stats.parallel_steps, 1U) << "iteration " << iter;
    EXPECT_EQ(stats.serial_fallback_steps, 0U) << "iteration " << iter;

    for (std::uint32_t r = 0; r < kRows; ++r) {
      const Value v = StoredValue(wb, 0U, r, 1U);
      ASSERT_TRUE(v.is_number()) << "iter " << iter << " row " << r;
      EXPECT_DOUBLE_EQ(v.as_number(), static_cast<double>(r) + 2.0) << "iter " << iter << " row " << r;
    }
  }
}

// ---------------------------------------------------------------------------
// Degradation when the OS refuses worker threads
// ---------------------------------------------------------------------------

TEST(Scheduler, NoWorkerThreadsFallsBackToSerialEvaluation) {
  // A host at its thread limit must still complete the recalc. Every
  // launch is refused here, so the pool starts empty and each layer takes
  // the calling-thread path; the numbers have to match what the fully
  // parallel diamond produces.
  ThreadLaunchInjection injection;
  WorkbookPair wp;
  wp.set_value(0, 0, 0, Value::number(10.0));
  wp.set_formula(0, 0, 1, "=A1*2");
  wp.set_formula(0, 0, 2, "=A1+5");
  wp.set_formula(0, 0, 3, "=B1+C1");

  set_thread_launch_failure_after(0U);
  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, &stats)));
  clear_thread_launch_failure_injection();

  EXPECT_EQ(stats.parallel_steps, 0U) << "no worker was started, so no layer can have been dispatched to the pool";
  EXPECT_GE(stats.serial_fallback_steps, 1U);
  EXPECT_EQ(stats.cells_evaluated, 3U);
  EXPECT_EQ(stats.worker_threads_started, 0U);
  EXPECT_EQ(stats.worker_threads_used, 0U);

  Value d1 = StoredValue(wp.parallel, 0, 0, 3);
  ASSERT_TRUE(d1.is_number());
  EXPECT_DOUBLE_EQ(d1.as_number(), 35.0);
}

TEST(Scheduler, PartialWorkerLaunchStillDispatchesWideLayers) {
  // Two workers out of the eight requested: the pass keeps using the pool
  // for layers wide enough to benefit, just with less of it.
  ThreadLaunchInjection injection;
  WorkbookPair wp;
  wp.set_value(0, 0, 0, Value::number(10.0));
  wp.set_formula(0, 0, 1, "=A1*2");
  wp.set_formula(0, 0, 2, "=A1+5");
  wp.set_formula(0, 0, 3, "=B1+C1");

  set_thread_launch_failure_after(2U);
  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, &stats)));
  clear_thread_launch_failure_injection();

  EXPECT_GE(stats.parallel_steps, 1U) << "the B/C layer should still reach the two workers that did start";
  EXPECT_EQ(stats.cells_evaluated, 3U);
  EXPECT_EQ(stats.worker_threads_started, 2U);
  EXPECT_GE(stats.worker_threads_used, 1U);
  EXPECT_LE(stats.worker_threads_used, stats.worker_threads_started);

  Value d1 = StoredValue(wp.parallel, 0, 0, 3);
  ASSERT_TRUE(d1.is_number());
  EXPECT_DOUBLE_EQ(d1.as_number(), 35.0);

  // And the degraded pass agrees with a plain serial recalc cell for cell.
  RecalcBothAndExpectEqual(wp, 8U);
}

TEST(Scheduler, MixedValueKindsMatchSerialRecalc) {
  WorkbookPair wp;
  wp.set_value(0, 0, 0, Value::text("hello"));
  wp.set_value(0, 0, 1, Value::boolean(true));
  wp.set_value(0, 0, 2, Value::error(ErrorCode::Name));
  wp.set_formula(0, 1, 0, "=A1&\" world\"");
  wp.set_formula(0, 1, 1, "=NOT(B1)");
  wp.set_formula(0, 1, 2, "=C1");

  RecalcBothAndExpectEqual(wp, 4U);
}

TEST(Scheduler, CycleRecoveryViaIterativeSolver) {
  // A1 = (B1 + 10) / 2, B1 = (A1 + 20) / 2. Averaging cycle that
  // converges. Iterative-calc enabled.
  Workbook wb = Workbook::create();
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 100U;
  opts.max_change = 1e-6;
  wb.set_iterative_options(opts);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=(B1+10)/2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=(A1+20)/2")));

  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_GE(stats.cycle_recoveries, 1U);
  EXPECT_EQ(stats.cells_evaluated, 2U);
}

TEST(Scheduler, ParallelIterativeTextCycleSurvivesArenaResets) {
  // The SCC is evaluated by one scheduler worker, but each member resets
  // that worker's Arena before evaluating. Both formulas return the same
  // Text value while depending on the other cell, so the second sweep must
  // compare an owned convergence snapshot rather than an arena string_view.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=IF(B1=1,\"stable\",\"stable\")")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=IF(A1=1,\"stable\",\"stable\")")));

  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 4U;
  opts.max_change = 0.001;
  wb.set_iterative_options(opts);

  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cycle_recoveries, 1U);
  EXPECT_EQ(stats.cells_evaluated, 2U);

  const Value a1 = StoredValue(wb, 0U, 0U, 0U);
  const Value b1 = StoredValue(wb, 0U, 0U, 1U);
  ASSERT_TRUE(a1.is_text());
  ASSERT_TRUE(b1.is_text());
  EXPECT_EQ(a1.as_text(), "stable");
  EXPECT_EQ(b1.as_text(), "stable");
}

TEST(Scheduler, ZeroThreadsAutoDetectsAndExecutes) {
  // num_threads = 0 means "auto-detect, capped at 8". The scheduler must
  // still resolve a sensible worker count and successfully complete.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1*A1")));

  SchedulerStats stats;
  SchedulerConfig cfg;  // num_threads = 0 → auto.
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cells_evaluated, 1U);

  Value b1 = StoredValue(wb, 0, 0, 1);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 4.0);
}

TEST(Scheduler, NullStatsDoesNotCrash) {
  // Passing nullptr for `stats` is supported and must not crash.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+10")));

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, /*stats=*/nullptr)));

  Value b1 = StoredValue(wb, 0, 0, 1);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 15.0);
}

TEST(Scheduler, RepeatedRecalcIsIdempotent) {
  // Running `recalc_parallel` twice in a row with no intervening edits
  // must produce the same values and zero counters on the second pass
  // (dirty set was cleared by the first pass).
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1*7")));

  SchedulerStats stats1;
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats1)));
  EXPECT_EQ(stats1.cells_evaluated, 1U);

  SchedulerStats stats2;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats2)));
  // Second pass: nothing dirty, no work.
  EXPECT_EQ(stats2.cells_evaluated, 0U);
  EXPECT_EQ(stats2.sccs_processed, 0U);

  Value b1 = StoredValue(wb, 0, 0, 1);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 21.0);
}

TEST(Scheduler, IncrementalEditPropagates) {
  // Edit-then-recalc semantics with the parallel engine. Mutate A1 after
  // the first recalc and confirm B1 reflects the new value.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  Value b1 = StoredValue(wb, 0, 0, 1);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 3.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  b1 = StoredValue(wb, 0, 0, 1);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 11.0);
}

TEST(Scheduler, ParallelMatchesSerialOnSimpleWorkbooks) {
  // Five workbook shapes, each driven through both `recalc()` and
  // `recalc_parallel()`. Per-cell value equality is asserted by
  // `RecalcBothAndExpectEqual`.

  {
    SCOPED_TRACE("shape: linear chain");
    WorkbookPair wp;
    wp.set_value(0, 0, 0, Value::number(3.0));
    wp.set_formula(0, 1, 0, "=A1*2");
    wp.set_formula(0, 2, 0, "=A2-1");
    RecalcBothAndExpectEqual(wp, 8U);
  }
  {
    SCOPED_TRACE("shape: range sum");
    WorkbookPair wp;
    for (std::uint32_t r = 0; r < 5; ++r) {
      wp.set_value(0, r, 0, Value::number(static_cast<double>(r + 1)));
    }
    wp.set_formula(0, 0, 1, "=SUM(A1:A5)");
    RecalcBothAndExpectEqual(wp, 4U);
  }
  {
    SCOPED_TRACE("shape: diamond");
    WorkbookPair wp;
    wp.set_value(0, 0, 0, Value::number(7.0));
    wp.set_formula(0, 0, 1, "=A1+1");
    wp.set_formula(0, 0, 2, "=A1*3");
    wp.set_formula(0, 0, 3, "=B1+C1");
    RecalcBothAndExpectEqual(wp, 4U);
  }
  {
    SCOPED_TRACE("shape: text concat chain");
    WorkbookPair wp;
    wp.set_value(0, 0, 0, Value::number(5.0));
    wp.set_formula(0, 1, 0, "=A1+10");
    wp.set_formula(0, 2, 0, "=IF(A2>10, 1, 0)");
    RecalcBothAndExpectEqual(wp, 8U);
  }
  {
    SCOPED_TRACE("shape: wide layer");
    WorkbookPair wp;
    for (std::uint32_t r = 0; r < 20; ++r) {
      wp.set_value(0, r, 0, Value::number(static_cast<double>(r)));
      wp.set_formula(0, r, 1, "=A" + std::to_string(r + 1) + "+1");
    }
    RecalcBothAndExpectEqual(wp, 8U);
  }
}

}  // namespace
}  // namespace formulon::eval
