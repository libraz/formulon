//
// Concurrency tests for the parallel SCC-layered recalc scheduler.
//
// These tests build small workbook fixtures that exercise the layering
// algorithm (linear chains, wide independent layers, diamonds, cycles)
// and assert that:
//
//   * `recalc_parallel` produces the same per-cell results as the
//     single-threaded `recalc()` engine;
//   * the per-pass `SchedulerStats` counters are consistent with the
//     dep-graph topology (parallel layers vs serial fallback);
//   * cyclic SCCs route through the iterative solver and bump the
//     `cycle_recoveries` counter on convergence.
//
// The slow stress test (`StressRandomDag`, label `SLOW`) drives a
// larger random DAG many times and is the primary safety net for race
// detection under ThreadSanitizer.

#include "eval/scheduler.h"

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstdint>
#include <mutex>
#include <random>
#include <set>
#include <string>
#include <thread>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/volatile_tracker.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/thread_launch.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

/// Clears any injected thread-launch failure when the test leaves scope.
/// Without this an assertion that aborts a degradation test mid-way would
/// leave every later test unable to start a worker.
struct ThreadLaunchInjection {
  ~ThreadLaunchInjection() { clear_thread_launch_failure_injection(); }
};

Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

struct ProgressAbortAfter {
  std::atomic<std::uint32_t> calls{0U};
  std::uint32_t abort_after = 0U;
};

struct ProgressThreadProbe {
  std::thread::id caller;
  std::mutex mutex;
  std::vector<std::thread::id> callback_threads;
};

bool RecordIterativeProgressThread(std::uint32_t /*iteration*/, double /*max_residual*/,
                                   std::uint32_t /*max_iterations*/, void* user_data) {
  auto* probe = static_cast<ProgressThreadProbe*>(user_data);
  std::lock_guard<std::mutex> guard(probe->mutex);
  probe->callback_threads.push_back(std::this_thread::get_id());
  return true;
}

struct WorkerThreadProbe {
  std::condition_variable condition;
  std::mutex mutex;
  std::set<std::thread::id> ids;
  bool synchronize = false;
  bool released = true;
  bool timed_out = false;
};

WorkerThreadProbe g_worker_thread_probe;

void ResetWorkerThreadProbe(bool synchronize) {
  std::lock_guard<std::mutex> guard(g_worker_thread_probe.mutex);
  g_worker_thread_probe.ids.clear();
  g_worker_thread_probe.synchronize = synchronize;
  g_worker_thread_probe.released = !synchronize;
  g_worker_thread_probe.timed_out = false;
}

Value RecordWorkerThreadImpl(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity != 1U) {
    return Value::error(ErrorCode::Value);
  }
  std::unique_lock<std::mutex> lock(g_worker_thread_probe.mutex);
  g_worker_thread_probe.ids.insert(std::this_thread::get_id());
  if (g_worker_thread_probe.synchronize && !g_worker_thread_probe.released) {
    // Keep the first worker occupied until a sibling SCC reaches this same
    // rendezvous. The bounded deadline only prevents a deadlock if the host
    // unexpectedly launches fewer workers than the fixture requests.
    const auto deadline = std::chrono::steady_clock::now() + std::chrono::seconds(1);
    while (!g_worker_thread_probe.released && g_worker_thread_probe.ids.size() < 2U) {
      if (g_worker_thread_probe.condition.wait_until(lock, deadline) == std::cv_status::timeout) {
        g_worker_thread_probe.timed_out = true;
        g_worker_thread_probe.released = true;
        g_worker_thread_probe.condition.notify_all();
        break;
      }
    }
    if (g_worker_thread_probe.ids.size() >= 2U && !g_worker_thread_probe.released) {
      g_worker_thread_probe.released = true;
      g_worker_thread_probe.condition.notify_all();
    }
  }
  return args[0];
}

Workbook* g_worker_reentry_workbook = nullptr;
std::atomic<int> g_worker_reentry_code{0};

Value WorkerReentryImpl(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  if (arity != 1U || g_worker_reentry_workbook == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  SchedulerConfig config;
  config.num_threads = 1U;
  const auto nested = g_worker_reentry_workbook->recalc_parallel(default_registry(), config, nullptr);
  g_worker_reentry_code.store(nested ? static_cast<int>(FormulonErrorCode::kOk) : static_cast<int>(nested.error().code),
                              std::memory_order_release);
  return args[0];
}

bool AbortIterativeSolve(std::uint32_t /*iteration*/, double /*max_residual*/, std::uint32_t /*max_iterations*/,
                         void* user_data) {
  auto* progress = static_cast<ProgressAbortAfter*>(user_data);
  return progress->calls.fetch_add(1U, std::memory_order_relaxed) + 1U < progress->abort_after;
}

// Builds two sibling workbooks and applies the same edits to both. Used
// by tests that run a recalc on one and `recalc_parallel` on the other,
// then compare cell values to confirm bit-for-bit equality.
struct WorkbookPair {
  Workbook serial;
  Workbook parallel;

  WorkbookPair() : serial(Workbook::create()), parallel(Workbook::create()) {}

  void set_value(std::size_t sheet, std::uint32_t row, std::uint32_t col, Value v) {
    ASSERT_TRUE(static_cast<bool>(serial.set_cell_value(sheet, row, col, v)));
    ASSERT_TRUE(static_cast<bool>(parallel.set_cell_value(sheet, row, col, v)));
  }

  void set_formula(std::size_t sheet, std::uint32_t row, std::uint32_t col, std::string formula) {
    ASSERT_TRUE(static_cast<bool>(serial.set_cell_formula(sheet, row, col, formula)));
    ASSERT_TRUE(static_cast<bool>(parallel.set_cell_formula(sheet, row, col, formula)));
  }
};

// Recalcs both members and asserts every cell in the populated rows of
// the serial sheet matches the parallel one. Cell-level equality uses
// `Value::operator==` (defined for the Value variant); arrays / spills
// are compared via `resolve_cell_value`.
void RecalcBothAndExpectEqual(WorkbookPair& wp, std::uint32_t threads = 4U) {
  ASSERT_TRUE(static_cast<bool>(wp.serial.recalc(default_registry())));
  SchedulerConfig cfg;
  cfg.num_threads = threads;
  ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, nullptr)));

  ASSERT_EQ(wp.serial.sheet_count(), wp.parallel.sheet_count());
  for (std::size_t s = 0; s < wp.serial.sheet_count(); ++s) {
    const Sheet& a = wp.serial.sheet(s);
    const Sheet& b = wp.parallel.sheet(s);
    for (const auto& [row, cells] : a.rows()) {
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        Value va = cells[col].cached_value;
        Value vb = StoredValue(wp.parallel, s, row, col);
        EXPECT_EQ(va, vb) << "value mismatch at (" << row << ", " << col << ")";
      }
    }
    (void)b;
  }
}

// ---------------------------------------------------------------------------
// Empty / trivial workbooks
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Topology-specific tests
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// SLOW: stress test — a larger random DAG run repeatedly to surface races
// under ThreadSanitizer.
// ---------------------------------------------------------------------------

TEST(SchedulerSlow, StressRandomDag) {
  // 200-cell random DAG, run 50 times. Each run rebuilds the workbook from
  // scratch and asserts every cell matches the single-threaded recalc.
  // Kept comfortably under the SLOW timeout budget (120 s).
  constexpr std::uint32_t kCellCount = 200U;
  constexpr int kIterations = 50;

  std::mt19937 rng(0xCAFEBABEU);

  for (int iter = 0; iter < kIterations; ++iter) {
    WorkbookPair wp;
    // Seed cells: rows 0..9 are numeric literals 1..10.
    for (std::uint32_t r = 0; r < 10; ++r) {
      wp.set_value(0, r, 0, Value::number(static_cast<double>(r + 1)));
    }
    // Subsequent cells reference up to 3 earlier rows in column A.
    for (std::uint32_t r = 10; r < kCellCount; ++r) {
      std::uniform_int_distribution<std::uint32_t> pick(0U, r - 1U);
      const std::uint32_t a = pick(rng);
      const std::uint32_t b = pick(rng);
      const std::string formula = "=A" + std::to_string(a + 1) + "+A" + std::to_string(b + 1);
      wp.set_formula(0, r, 0, formula);
    }

    SchedulerConfig cfg;
    cfg.num_threads = 8U;
    ASSERT_TRUE(static_cast<bool>(wp.serial.recalc(default_registry())));
    ASSERT_TRUE(static_cast<bool>(wp.parallel.recalc_parallel(default_registry(), cfg, nullptr)));

    for (std::uint32_t r = 0; r < kCellCount; ++r) {
      Value vs = StoredValue(wp.serial, 0, r, 0);
      Value vp = StoredValue(wp.parallel, 0, r, 0);
      ASSERT_EQ(vs.kind(), vp.kind()) << "iter=" << iter << " r=" << r;
      if (vs.is_number()) {
        ASSERT_DOUBLE_EQ(vs.as_number(), vp.as_number()) << "iter=" << iter << " r=" << r;
      }
    }
  }
}

// Two independent dynamic-array formulas on the same sheet in the same
// parallel layer, with disjoint spill footprints, evaluated repeatedly
// under a 4-thread pool. Each formula spills through `Sheet::commit_spill`
// / `Sheet::resolve_cell_value`, which mutate and read the sheet's spill
// table and row store. Run under ThreadSanitizer this is the race-detection
// fixture for concurrent spill commits on a shared sheet; the value
// assertions also catch a lost / torn spill (a phantom reading back as
// #SPILL! or blank).
TEST(SchedulerSlow, ParallelSpillNoDataRace) {
  constexpr int kIterations = 40;
  for (int iter = 0; iter < kIterations; ++iter) {
    Workbook wb = Workbook::create();
    // A1 spills A1:A4 = {1, 2, 3, 4}.
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(4,1)"))) << "iter " << iter;
    // C1 spills C1:C4 = {2, 4, 6, 8}. Disjoint footprint from column A.
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=SEQUENCE(4,1)*2"))) << "iter " << iter;

    SchedulerConfig cfg;
    cfg.num_threads = 4U;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))) << "iter " << iter;

    const Sheet& s = wb.sheet(0);
    for (std::uint32_t r = 0; r < 4U; ++r) {
      const Value a = s.resolve_cell_value(r, 0U);
      ASSERT_TRUE(a.is_number()) << "iter " << iter << " A row " << r << " kind=" << static_cast<int>(a.kind());
      EXPECT_DOUBLE_EQ(a.as_number(), static_cast<double>(r + 1U)) << "iter " << iter << " A row " << r;

      const Value c = s.resolve_cell_value(r, 2U);
      ASSERT_TRUE(c.is_number()) << "iter " << iter << " C row " << r << " kind=" << static_cast<int>(c.kind());
      EXPECT_DOUBLE_EQ(c.as_number(), static_cast<double>((r + 1U) * 2U)) << "iter " << iter << " C row " << r;
    }
  }
}

}  // namespace
}  // namespace formulon::eval
