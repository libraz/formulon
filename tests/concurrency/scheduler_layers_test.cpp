// Concurrency scheduler tests grouped by fixture size and execution mode.
// Shared probes and workbook-pair helpers live in scheduler_test_helpers.h.

#include "scheduler_test_helpers.h"

namespace formulon::eval {
namespace {

using namespace scheduler_test;

TEST(Scheduler, TwoCellChainSerialLayers) {
  // A1 = 1, B1 = =A1+1. Two layers, each of size 1 -> serial dispatch on
  // both. parallel_steps must remain 0.
  WorkbookPair wp;
  wp.set_value(0, 0, 0, Value::number(1.0));
  wp.set_formula(0, 0, 1, "=A1+1");

  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.parallel_steps, 0U);
  // The single dirty SCC (B1) is dispatched serially.
  EXPECT_GE(stats.serial_fallback_steps, 1U);

  Value b1 = StoredValue(wp.parallel, 0, 0, 1);
  ASSERT_TRUE(b1.is_number());
  EXPECT_DOUBLE_EQ(b1.as_number(), 2.0);
}

TEST(Scheduler, WideIndependentLayer) {
  // 100 independent literals A1..A100. No dependencies among them, so a
  // single layer of 100 super-nodes is dispatched in parallel. There are
  // no formulas to evaluate (the cells are all literals), so this case
  // primarily exercises the dirty-set / scheduler bookkeeping with a
  // wide layer.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 100; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, r, 0U, Value::number(static_cast<double>(r)))));
  }
  // Now add 100 formulas in column B that each read one A cell — this
  // gives a layer of 100 independent super-nodes.
  for (std::uint32_t r = 0; r < 100; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, r, 1U, "=A" + std::to_string(r + 1) + "*2")));
  }

  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cells_evaluated, 100U);
  EXPECT_GE(stats.parallel_steps, 1U);
  EXPECT_GE(stats.worker_threads_started, 2U);
  EXPECT_GE(stats.worker_threads_used, 1U);

  for (std::uint32_t r = 0; r < 100; ++r) {
    Value v = StoredValue(wb, 0U, r, 1U);
    ASSERT_TRUE(v.is_number()) << "row " << r;
    EXPECT_DOUBLE_EQ(v.as_number(), static_cast<double>(r) * 2.0);
  }
}

TEST(Scheduler, WideLayerExecutesOnDistinctWorkerThreads) {
  Workbook wb = Workbook::create();
  FunctionRegistry registry;
  FunctionDef def{};
  def.canonical_name = "RECORD_WORKER_THREAD";
  def.min_arity = 1U;
  def.max_arity = 1U;
  def.impl = &RecordWorkerThreadImpl;
  ASSERT_TRUE(registry.register_function(def));

  ResetWorkerThreadProbe(/*synchronize=*/true);

  constexpr std::uint32_t kRows = 16U;
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(static_cast<double>(row + 1U)))));
    ASSERT_TRUE(static_cast<bool>(
        wb.set_cell_formula(0U, row, 1U, "=RECORD_WORKER_THREAD(A" + std::to_string(row + 1U) + ")")));
  }

  SchedulerConfig config;
  config.num_threads = 4U;
  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(recalc_parallel(wb, registry, config, &stats)));
  EXPECT_GE(stats.worker_threads_started, 2U);
  EXPECT_GE(stats.worker_threads_used, 2U);
  EXPECT_GE(stats.parallel_steps, 1U);

  std::set<std::thread::id> observed_threads;
  bool timed_out = false;
  {
    std::lock_guard<std::mutex> guard(g_worker_thread_probe.mutex);
    observed_threads = g_worker_thread_probe.ids;
    timed_out = g_worker_thread_probe.timed_out;
  }
  EXPECT_FALSE(timed_out);
  EXPECT_GE(observed_threads.size(), 2U);
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const Value value = StoredValue(wb, 0U, row, 1U);
    ASSERT_TRUE(value.is_number()) << "row " << row;
    EXPECT_DOUBLE_EQ(value.as_number(), static_cast<double>(row + 1U));
  }
}

// ---------------------------------------------------------------------------
// Volatility classes and the layer split
// ---------------------------------------------------------------------------

TEST(Scheduler, ValueVolatileLayerStaysOnThePool) {
  // `RAND` re-fires every pass but reads nothing, so the layering derived
  // from the dependency edges is complete and every one of these cells is
  // poolable. A pass must therefore dispatch the layer to the workers and
  // leave no serial tail behind: if the scheduler ever went back to
  // isolating value volatility, `pooled` would be empty here and the whole
  // layer would fall to the caller, flipping both counters.
  constexpr std::uint32_t kRows = 64U;
  Workbook wb = Workbook::create();
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 0U, "=RAND()*0+1")));
  }

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));

  EXPECT_EQ(stats.cells_evaluated, kRows);
  EXPECT_EQ(stats.parallel_steps, 1U);
  EXPECT_EQ(stats.serial_fallback_steps, 0U);
  EXPECT_GE(stats.worker_threads_used, 1U);
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const Value v = StoredValue(wb, 0U, row, 0U);
    ASSERT_TRUE(v.is_number()) << "row " << row;
    EXPECT_DOUBLE_EQ(v.as_number(), 1.0) << "row " << row;
  }
}

TEST(Scheduler, DynamicReferenceLayerLeavesThePoolButValueVolatileDoesNot) {
  // One layer holding both classes and nothing else: the `INDIRECT` half
  // must be the serial tail, the `RAND` half must be the pooled body.
  // Column A is settled before the measured pass so it contributes no
  // super-node — re-merging the classes would leave the pool with nothing
  // to do and drop `parallel_steps` to zero. The tracker assertions pin
  // which cell landed in which class.
  constexpr std::uint32_t kRows = 32U;
  Workbook wb = Workbook::create();
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(static_cast<double>(row)))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=RAND()*0+1")));
    ASSERT_TRUE(
        static_cast<bool>(wb.set_cell_formula(0U, row, 2U, "=INDIRECT(\"A" + std::to_string(row + 1U) + "\")")));
  }

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));

  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cells_evaluated, 2U * kRows);
  EXPECT_EQ(stats.sccs_processed, 2U * kRows);
  EXPECT_EQ(stats.parallel_steps, 1U);
  EXPECT_EQ(stats.serial_fallback_steps, 1U);

  const VolatileTracker& volatiles = wb.recalc_engine().volatiles();
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const CellNodeId value_volatile{0U, row, 1U};
    const CellNodeId dynamic_volatile{0U, row, 2U};
    EXPECT_TRUE(volatiles.contains(value_volatile)) << "row " << row;
    EXPECT_FALSE(volatiles.contains_dynamic_reference(value_volatile)) << "row " << row;
    EXPECT_TRUE(volatiles.contains(dynamic_volatile)) << "row " << row;
    EXPECT_TRUE(volatiles.contains_dynamic_reference(dynamic_volatile)) << "row " << row;

    const Value value_result = StoredValue(wb, 0U, row, 1U);
    ASSERT_TRUE(value_result.is_number()) << "row " << row;
    EXPECT_DOUBLE_EQ(value_result.as_number(), 1.0) << "row " << row;
    const Value dynamic_result = StoredValue(wb, 0U, row, 2U);
    ASSERT_TRUE(dynamic_result.is_number()) << "row " << row;
    EXPECT_DOUBLE_EQ(dynamic_result.as_number(), static_cast<double>(row)) << "row " << row;
  }
}

TEST(Scheduler, MixedVolatilityClassesMatchSerialRecalc) {
  // Both classes in one workbook, with every value-volatile result folded
  // back to a fixed number so an ordering difference cannot hide behind a
  // legitimately changing value. Column A is the target the dynamic
  // references resolve to at evaluation time.
  constexpr std::uint32_t kRows = 48U;
  WorkbookPair wp;
  wp.set_value(0, 0, 4, Value::number(0.0));  // E1: the OFFSET base.
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string row_1based = std::to_string(row + 1U);
    wp.set_formula(0, row, 0, "=RAND()*0+" + row_1based);
    wp.set_formula(0, row, 1, "=INDIRECT(\"A" + row_1based + "\")");
    wp.set_formula(0, row, 2, "=IF(NOW()>0,A" + row_1based + ",-1)");
    wp.set_formula(0, row, 3, "=OFFSET($E$1," + std::to_string(row) + ",-4)");
  }

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  // Settling pass: no edge orders a dynamic reference against the cell it
  // reads, so only the steady state is comparable across engines.
  ASSERT_TRUE(static_cast<bool>(wp.serial.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, nullptr)));

  for (std::uint32_t pass = 0U; pass < 4U; ++pass) {
    ASSERT_TRUE(static_cast<bool>(wp.serial.recalc(default_registry())));
    ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, nullptr)));
    for (std::uint32_t row = 0U; row < kRows; ++row) {
      for (std::uint32_t col = 0U; col < 4U; ++col) {
        const Value expected = StoredValue(wp.serial, 0U, row, col);
        ASSERT_TRUE(expected.is_number()) << "pass " << pass << " at (" << row << ", " << col << ")";
        EXPECT_DOUBLE_EQ(expected.as_number(), static_cast<double>(row + 1U))
            << "pass " << pass << " at (" << row << ", " << col << ")";
        EXPECT_EQ(expected, StoredValue(wp.parallel, 0U, row, col))
            << "pass " << pass << " value mismatch at (" << row << ", " << col << ")";
      }
    }
  }
}

TEST(Scheduler, SingleThreadConfigKeepsEvaluationOnCaller) {
  Workbook wb = Workbook::create();
  FunctionRegistry registry;
  FunctionDef def{};
  def.canonical_name = "RECORD_CALLER_THREAD";
  def.min_arity = 1U;
  def.max_arity = 1U;
  def.impl = &RecordWorkerThreadImpl;
  ASSERT_TRUE(registry.register_function(def));

  ResetWorkerThreadProbe(/*synchronize=*/false);
  const std::thread::id caller = std::this_thread::get_id();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=RECORD_CALLER_THREAD(A1)")));

  SchedulerConfig config;
  config.num_threads = 1U;
  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(recalc_parallel(wb, registry, config, &stats)));
  EXPECT_EQ(stats.worker_threads_started, 0U);
  EXPECT_EQ(stats.worker_threads_used, 0U);

  std::set<std::thread::id> observed;
  {
    std::lock_guard<std::mutex> guard(g_worker_thread_probe.mutex);
    observed = g_worker_thread_probe.ids;
  }
  ASSERT_EQ(observed.size(), 1U);
  EXPECT_EQ(*observed.begin(), caller);
}

TEST(Scheduler, NestedRecalcFromWorkerReturnsReentrant) {
  Workbook wb = Workbook::create();
  FunctionRegistry registry;
  FunctionDef def{};
  def.canonical_name = "NESTED_WORKER_RECALC";
  def.min_arity = 1U;
  def.max_arity = 1U;
  def.impl = &WorkerReentryImpl;
  ASSERT_TRUE(registry.register_function(def));

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=NESTED_WORKER_RECALC(A1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=NESTED_WORKER_RECALC(A2)")));

  g_worker_reentry_workbook = &wb;
  g_worker_reentry_code.store(0, std::memory_order_release);
  SchedulerConfig config;
  config.num_threads = 4U;
  const auto outer = recalc_parallel(wb, registry, config, nullptr);
  g_worker_reentry_workbook = nullptr;

  ASSERT_TRUE(static_cast<bool>(outer));
  EXPECT_EQ(g_worker_reentry_code.load(std::memory_order_acquire),
            static_cast<int>(FormulonErrorCode::kGraphRecalcReentrant));
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 0U, 1U).as_number(), 10.0);
  EXPECT_DOUBLE_EQ(StoredValue(wb, 0U, 1U, 1U).as_number(), 20.0);
}

TEST(Scheduler, DeepChainNoParallelism) {
  // 100-cell deep chain: A1=1, A2=A1+1, ..., A100=A99+1. Every cell is
  // its own layer (size 1) so parallel_steps must stay 0.
  WorkbookPair wp;
  wp.set_value(0, 0, 0, Value::number(1.0));
  for (std::uint32_t r = 1; r < 100; ++r) {
    wp.set_formula(0, r, 0, "=A" + std::to_string(r) + "+1");
  }

  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.parallel_steps, 0U);
  EXPECT_EQ(stats.cells_evaluated, 99U);

  Value last = StoredValue(wp.parallel, 0, 99, 0);
  ASSERT_TRUE(last.is_number());
  EXPECT_DOUBLE_EQ(last.as_number(), 100.0);

  RecalcBothAndExpectEqual(wp, 8U);
}

TEST(Scheduler, DiamondParallelLayer) {
  // A1 -> B1, A1 -> C1, B1+C1 -> D1.
  // Layer 0: A1 (literal — not a formula, so no dirty SCC for it).
  // Layer 1: B1 and C1 (parallel).
  // Layer 2: D1.
  WorkbookPair wp;
  wp.set_value(0, 0, 0, Value::number(10.0));
  wp.set_formula(0, 0, 1, "=A1*2");
  wp.set_formula(0, 0, 2, "=A1+5");
  wp.set_formula(0, 0, 3, "=B1+C1");

  SchedulerStats stats;
  SchedulerConfig cfg;
  cfg.num_threads = 8U;
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_EQ(stats.cells_evaluated, 3U);
  EXPECT_GE(stats.parallel_steps, 1U);  // The B/C layer.

  Value d1 = StoredValue(wp.parallel, 0, 0, 3);
  ASSERT_TRUE(d1.is_number());
  EXPECT_DOUBLE_EQ(d1.as_number(), 35.0);

  RecalcBothAndExpectEqual(wp, 8U);
}

}  // namespace
}  // namespace formulon::eval
