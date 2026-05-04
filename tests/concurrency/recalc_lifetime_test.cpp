// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Lifetime / re-entrancy / external-reader concurrency tests for the
// parallel recalc scheduler.
//
// These tests complement `scheduler_test.cpp` (which exercises the
// layering algorithm under TSan via the `SchedulerSlow.StressRandomDag`
// fixture) by covering the *contract* edges that the algorithm tests do
// not touch:
//
//   * Workbook destruction after a `recalc_parallel` call returns: the
//     scheduler joins every worker before returning, so a subsequent
//     `unique_ptr<Workbook>::reset()` is safe.
//   * Concurrent reads on a workbook while `recalc_parallel` is running.
//     The scheduler holds an internal mutex only across `set_cell_*`
//     writes; const observers (`Workbook::sheet`, `Sheet::cell_at`,
//     `Sheet::resolve_cell_value`) take no lock, and the contract says
//     reads on cells DISJOINT from the dirty set are race-free.
//   * Sequential `recalc_parallel` invocations: 50 back-to-back passes on
//     the same workbook must stay TSan-clean.
//   * Re-entrant recalc detection: a UDF that calls
//     `Workbook::recalc_parallel` from inside an evaluator callback must
//     surface `kGraphRecalcReentrant` and not corrupt the outer pass.
//
// Every test joins all spawned threads. Detached threads are forbidden by
// project policy and would defeat the TSan invariant of "every event
// observed before exit".

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <thread>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Helper: read the cached_value for `(row, col)` on `sheet_index`.
Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

// Builds a small chain workbook: A1 = 1, B1 = A1*2, C1 = B1+5.
// Sufficient to exercise a non-trivial dep graph without inflating the
// TSan event budget.
void SeedTinyWorkbook(Workbook& wb) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1*2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 2U, "=B1+5")));
}

// ---------------------------------------------------------------------------
// Workbook destruction after recalc completes
// ---------------------------------------------------------------------------

TEST(RecalcLifetime, DestroyWorkbookAfterRecalcReturns) {
  // The scheduler joins all workers before returning, so the workbook is
  // safe to destroy on the calling thread once `recalc_parallel` has
  // returned. This test exercises that path: spawn `recalc_parallel`,
  // wait for it via thread join, then drop the unique_ptr. Under TSan
  // any worker still holding a reference would surface as a use-after-free.
  auto wb = std::make_unique<Workbook>(Workbook::create());
  SeedTinyWorkbook(*wb);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb->recalc_parallel(default_registry(), cfg, nullptr)));

  // Read a value to confirm the recalc actually committed before destroy.
  EXPECT_TRUE(StoredValue(*wb, 0, 0, 1).is_number());

  // Destroy. If a worker were still alive it would race with the dtor.
  wb.reset();
  // Reaching here without TSan complaint is the assertion.
  SUCCEED();
}

// ---------------------------------------------------------------------------
// Sequential recalc_parallel calls — TSan-clean steady state
// ---------------------------------------------------------------------------

TEST(RecalcLifetime, FiftySequentialRecalcParallelCallsTSanClean) {
  // 50 back-to-back recalc passes on the same workbook. Each pass should
  // be a no-op after the first (nothing is dirty) but the scheduler still
  // walks its layering bookkeeping. TSan must observe no race across the
  // entire sequence.
  Workbook wb = Workbook::create();
  SeedTinyWorkbook(wb);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  for (int i = 0; i < 50; ++i) {
    SchedulerStats stats;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats))) << "pass " << i;
    if (i == 0) {
      EXPECT_GE(stats.cells_evaluated, 2U);
    } else {
      // Subsequent passes have nothing dirty.
      EXPECT_EQ(stats.cells_evaluated, 0U);
    }
  }
}

// ---------------------------------------------------------------------------
// Sequential recalc, mutate-and-recalc, mutate-and-recalc loop
// ---------------------------------------------------------------------------

TEST(RecalcLifetime, EditRecalcLoopTSanClean) {
  // Edit + recalc, 20 iterations. Each edit dirties B1 and C1
  // transitively; the scheduler must produce consistent results without
  // any race.
  Workbook wb = Workbook::create();
  SeedTinyWorkbook(wb);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  for (int i = 0; i < 20; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i)))));
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
    Value c1 = StoredValue(wb, 0, 0, 2);
    ASSERT_TRUE(c1.is_number()) << "iter " << i;
    EXPECT_DOUBLE_EQ(c1.as_number(), static_cast<double>(i) * 2.0 + 5.0);
  }
}

// ---------------------------------------------------------------------------
// Concurrent reads on disjoint cells while recalc is in flight
// ---------------------------------------------------------------------------

TEST(RecalcLifetime, ConcurrentReadOnDisjointCellsDuringRecalc) {
  // Build a workbook with two disjoint sub-graphs: column A is the
  // "active" set being recalc'd, column D contains static literals that
  // the reader thread observes. The reader only ever touches column D,
  // which the scheduler does not write because no formula in A reads
  // from D and the literals were committed before the recalc started.
  // Reads on a Sheet::cell_at pointer therefore race only with the
  // const-only iteration internals of the scheduler — which the parallel
  // engine does not mutate while its workers are running.
  Workbook wb = Workbook::create();

  // Active column A: 50 cells of dependent formulas.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  for (std::uint32_t r = 1; r < 50U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, r, 0U, "=A" + std::to_string(r) + "+1")));
  }
  // Static column D: 50 literals. NOT marked dirty (already up to date).
  for (std::uint32_t r = 0; r < 50U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, r, 3U, Value::number(static_cast<double>(r * 100)))));
  }
  // Run an initial recalc to clear the dirty set entirely.
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  // Re-dirty column A by mutating the seed cell.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));

  std::atomic<bool> reader_observed_mismatch{false};
  std::atomic<bool> reader_should_run{true};
  // Reader thread loops on column D. Every read must observe the literal
  // value written before recalc_parallel started.
  std::thread reader([&] {
    while (reader_should_run.load(std::memory_order_acquire)) {
      for (std::uint32_t r = 0; r < 50U; ++r) {
        Value v = StoredValue(wb, 0U, r, 3U);
        if (!v.is_number() || v.as_number() != static_cast<double>(r * 100)) {
          reader_observed_mismatch.store(true, std::memory_order_release);
          return;
        }
      }
    }
  });

  // Drive the recalc on the main thread.
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));

  // Stop the reader and join.
  reader_should_run.store(false, std::memory_order_release);
  reader.join();

  EXPECT_FALSE(reader_observed_mismatch.load());

  // Confirm column A converged to the expected values.
  Value last = StoredValue(wb, 0, 49U, 0U);
  ASSERT_TRUE(last.is_number());
  EXPECT_DOUBLE_EQ(last.as_number(), 51.0);  // 2 + 49 * 1
}

// ---------------------------------------------------------------------------
// Concurrent reads on completely separate workbooks (independence test)
// ---------------------------------------------------------------------------

TEST(RecalcLifetime, IndependentWorkbooksRecalcInParallel) {
  // Eight worker threads, each owning its own Workbook, each driving
  // recalc_parallel concurrently. Workbooks share no state (the function
  // registry is read-only post-init), so this must be TSan-clean even
  // though the workers run truly simultaneously.
  constexpr int kThreads = 8;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<int> failures{0};

  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&failures, i] {
      Workbook wb = Workbook::create();
      if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i + 1))))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1*A1"))) {
        ++failures;
        return;
      }
      SchedulerConfig cfg;
      cfg.num_threads = 2U;
      for (int pass = 0; pass < 10; ++pass) {
        if (!static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))) {
          ++failures;
          return;
        }
      }
      const Sheet& s = wb.sheet(0);
      const Cell* c = s.cell_at(0U, 1U);
      if (c == nullptr || !c->cached_value.is_number()) {
        ++failures;
        return;
      }
      const double expected = static_cast<double>((i + 1) * (i + 1));
      if (c->cached_value.as_number() != expected) {
        ++failures;
      }
    });
  }
  for (std::thread& t : workers) {
    t.join();
  }
  EXPECT_EQ(failures.load(), 0);
}

// ---------------------------------------------------------------------------
// Re-entrant recalc detection
// ---------------------------------------------------------------------------

namespace {

// Test fixture for the re-entrancy probe: a UDF impl reaches the workbook
// through a static pointer and attempts a nested `recalc_parallel`. The
// returned error code is captured in `g_reentry_observed_code` so the
// surrounding test can inspect it after the outer recalc completes.
//
// The UDF is registered into a *test-local* `FunctionRegistry` (not the
// process-wide `default_registry()`), which means we cannot drive it from
// `Workbook::recalc_parallel(registry, cfg, stats)`'s default registry
// path. We instead invoke `eval::recalc_parallel(wb, custom_registry,
// cfg, stats)` directly. The fixture below covers the test plumbing.
Workbook* g_reentry_workbook = nullptr;
std::atomic<int> g_reentry_observed_code{0};
std::atomic<int> g_reentry_call_count{0};

Value reentry_probe_impl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  // Triggered while the outer `recalc_parallel` is still on the stack.
  // The nested call must short-circuit with `kGraphRecalcReentrant` and
  // leave the outer pass intact.
  ++g_reentry_call_count;
  if (g_reentry_workbook == nullptr) {
    return Value::number(0.0);
  }
  SchedulerConfig cfg;
  cfg.num_threads = 1U;
  auto result = g_reentry_workbook->recalc_parallel(default_registry(), cfg, nullptr);
  if (!result) {
    g_reentry_observed_code.store(static_cast<int>(result.error().code), std::memory_order_release);
  } else {
    // Treat success as the failure mode the test is meant to catch:
    // nested recalc was supposed to be rejected.
    g_reentry_observed_code.store(static_cast<int>(FormulonErrorCode::kOk), std::memory_order_release);
  }
  return Value::number(42.0);
}

}  // namespace

TEST(RecalcLifetime, NestedRecalcParallelReturnsRecalcReentrant) {
  // Build a workbook whose only formula calls our reentry-probe UDF. The
  // UDF runs on the worker thread (or the calling thread for a singleton
  // layer). `g_in_recalc` is `thread_local` and set on the same thread
  // that is dispatching, so the nested call observes the flag and returns
  // `kGraphRecalcReentrant`.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=PROBE_NESTED_RECALC()")));

  // Custom registry containing the probe.
  FunctionRegistry registry;
  FunctionDef def{};
  def.canonical_name = "PROBE_NESTED_RECALC";
  def.min_arity = 0U;
  def.max_arity = 0U;
  def.impl = &reentry_probe_impl;
  ASSERT_TRUE(registry.register_function(def));

  g_reentry_workbook = &wb;
  g_reentry_observed_code.store(-1, std::memory_order_release);
  g_reentry_call_count.store(0, std::memory_order_release);

  SchedulerConfig cfg;
  cfg.num_threads = 1U;  // Single-thread keeps the UDF on the main thread.
  auto outer = recalc_parallel(wb, registry, cfg, nullptr);
  // The outer pass succeeds; the rejection happens *inside* the UDF's
  // nested call.
  EXPECT_TRUE(static_cast<bool>(outer));
  EXPECT_GE(g_reentry_call_count.load(), 1) << "UDF was never invoked";
  EXPECT_EQ(g_reentry_observed_code.load(), static_cast<int>(FormulonErrorCode::kGraphRecalcReentrant));

  // Outer recalc still committed the UDF's return value (42).
  Value v = StoredValue(wb, 0, 0, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 42.0);

  g_reentry_workbook = nullptr;
}

TEST(RecalcLifetime, ReentryFlagClearedAfterOuterReturns) {
  // After a successful `recalc_parallel` returns, the per-thread flag
  // must be cleared so a subsequent invocation on the same thread is
  // accepted normally. Regression: forgetting to reset the flag turns
  // the engine into a one-shot resource.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));

  SchedulerConfig cfg;
  cfg.num_threads = 2U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  // Second call must succeed (flag was cleared).
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  // Third call from a freshly-spawned thread must also succeed: the flag
  // is per-thread, not global.
  std::atomic<bool> ok{false};
  std::thread t([&] { ok.store(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))); });
  t.join();
  EXPECT_TRUE(ok.load());
}

namespace {

// Serial-recalc analogue of `reentry_probe_impl`: nested call routed
// through `Workbook::recalc()` rather than `recalc_parallel()`. Without
// the shared `g_in_recalc` flag the inner serial call would deadlock on
// the engine mutex; with the guard it surfaces `kGraphRecalcReentrant`.
Value serial_reentry_probe_impl(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  ++g_reentry_call_count;
  if (g_reentry_workbook == nullptr) {
    return Value::number(0.0);
  }
  auto result = g_reentry_workbook->recalc(default_registry());
  if (!result) {
    g_reentry_observed_code.store(static_cast<int>(result.error().code), std::memory_order_release);
  } else {
    g_reentry_observed_code.store(static_cast<int>(FormulonErrorCode::kOk), std::memory_order_release);
  }
  return Value::number(7.0);
}

}  // namespace

TEST(RecalcLifetime, NestedSerialRecalcReturnsRecalcReentrant) {
  // Outer `Workbook::recalc()` evaluates a UDF that re-enters
  // `recalc()` on the same thread. The shared `g_in_recalc` flag must
  // observe the outer pass and reject the inner call instead of
  // deadlocking on the engine mutex.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=PROBE_NESTED_SERIAL_RECALC()")));

  FunctionRegistry registry;
  FunctionDef def{};
  def.canonical_name = "PROBE_NESTED_SERIAL_RECALC";
  def.min_arity = 0U;
  def.max_arity = 0U;
  def.impl = &serial_reentry_probe_impl;
  ASSERT_TRUE(registry.register_function(def));

  g_reentry_workbook = &wb;
  g_reentry_observed_code.store(-1, std::memory_order_release);
  g_reentry_call_count.store(0, std::memory_order_release);

  auto outer = wb.recalc(registry);
  EXPECT_TRUE(static_cast<bool>(outer));
  EXPECT_GE(g_reentry_call_count.load(), 1) << "UDF was never invoked";
  EXPECT_EQ(g_reentry_observed_code.load(), static_cast<int>(FormulonErrorCode::kGraphRecalcReentrant));

  Value v = StoredValue(wb, 0, 0, 0);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);

  g_reentry_workbook = nullptr;
}

// ---------------------------------------------------------------------------
// recalc_parallel followed by serial recalc on the same workbook
// ---------------------------------------------------------------------------

TEST(RecalcLifetime, AlternatingParallelAndSerialRecalcTSanClean) {
  // The serial `Workbook::recalc()` and the parallel
  // `Workbook::recalc_parallel()` share the same underlying `RecalcEngine`
  // but are never concurrent — they run sequentially on the same thread
  // here. The dirty / volatile state must transfer cleanly between them.
  Workbook wb = Workbook::create();
  SeedTinyWorkbook(wb);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;

  for (int i = 0; i < 10; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i)))));
    if ((i % 2) == 0) {
      ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
    } else {
      ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
    }
    Value c1 = StoredValue(wb, 0, 0, 2);
    ASSERT_TRUE(c1.is_number());
    EXPECT_DOUBLE_EQ(c1.as_number(), static_cast<double>(i) * 2.0 + 5.0);
  }
}

}  // namespace
}  // namespace formulon::eval
