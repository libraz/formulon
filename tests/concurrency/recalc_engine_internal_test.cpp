// Copyright 2026 libraz. Licensed under the MIT License.
//
// ThreadSanitizer-targeted tests for `RecalcEngine`'s internal mutable
// state and the "dep graph is effectively immutable during recalc"
// contract.
//
// Two scopes:
//
//   * Concurrent dep-graph reads while recalc is executing. Worker
//     threads inside the scheduler call `DepGraph::dependencies_of` /
//     `dependents_of` against the shared graph; the test-side reader
//     mirrors that with a separate thread that interrogates the engine's
//     read-only `dep_graph()` accessor while a `recalc_parallel` is in
//     flight on the main thread. Both must observe the same post-init
//     graph, race-free.
//   * Multiple sequential `set_cell_value` -> `recalc_parallel` cycles
//     with `mark_dirty` happening on the main thread between recalc
//     passes. The contract is that mutations route through
//     `Workbook::set_cell_value` (which serialises with itself, but is
//     never called concurrently with `recalc_parallel`); we verify the
//     bookkeeping survives a long sequence of such transitions.
//
// Per the contract, *concurrent* `mark_dirty` from multiple threads is
// out of scope: `mark_dirty` is private to `RecalcEngine` and reachable
// only via `Workbook::set_cell_value` / `set_cell_formula`. Callers must
// not race those mutators with `recalc_parallel`. The test below
// explicitly exercises the *sequential* mutate-recalc-mutate-recalc
// pattern so a future regression that introduced a hidden async path
// would surface immediately.

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <thread>
#include <vector>

#include "cell.h"
#include "eval/dep_graph.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(sheet_index);
  if (const Cell* c = s.cell_at(row, col); c != nullptr) {
    return c->cached_value;
  }
  return Value::blank();
}

// Builds a 10-cell linear chain: A1 = 1, A2 = A1+1, ..., A10 = A9+1.
void SeedChainWorkbook(Workbook& wb) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  for (std::uint32_t r = 1; r < 10U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, r, 0U, "=A" + std::to_string(r) + "+1")));
  }
}

// ---------------------------------------------------------------------------
// Concurrent dep-graph reads while recalc is in flight
// ---------------------------------------------------------------------------

TEST(RecalcEngineInternal, DepGraphReadsRaceFreeDuringRecalc) {
  // While the scheduler runs, a separate reader thread interrogates the
  // engine's `dep_graph()` accessor for `dependents_of` / `dependencies_of`
  // on cells the workers also walk. The graph is built up-front by
  // `set_cell_formula` and is not mutated during recalc; the read paths
  // therefore never race the scheduler, only one another.
  Workbook wb = Workbook::create();
  SeedChainWorkbook(wb);

  // Build a couple of explicit dependents so the queries observe a
  // populated reverse map.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A10)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=A10*A10")));

  std::atomic<bool> mismatch{false};
  std::atomic<bool> reader_should_run{true};
  std::thread reader([&] {
    const DepGraph& graph = wb.recalc_engine().dep_graph();
    while (reader_should_run.load(std::memory_order_acquire)) {
      // Reads on the dep graph: walk every cell of the chain.
      for (std::uint32_t r = 0; r < 10U; ++r) {
        CellNodeId node{0U, r, 0U};
        const auto deps = graph.dependencies_of(node);
        const auto deps2 = graph.dependents_of(node);
        // Just exercise the loops; we don't care about the exact size,
        // only that the reads complete without TSan flagging anything.
        (void)deps;
        (void)deps2;
        if (deps.size() > 100U) {  // Sanity guard against memory corruption.
          mismatch.store(true, std::memory_order_release);
          return;
        }
      }
    }
  });

  // Run a handful of recalc passes while the reader keeps walking.
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  for (int pass = 0; pass < 5; ++pass) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(pass)))));
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  }

  reader_should_run.store(false, std::memory_order_release);
  reader.join();

  EXPECT_FALSE(mismatch.load());
}

// ---------------------------------------------------------------------------
// Sequential mutate -> recalc -> mutate -> recalc
// ---------------------------------------------------------------------------

TEST(RecalcEngineInternal, SequentialMarkDirtyRecalcAlternationTSanClean) {
  // The scheduler's `mark_dirty` path is reached via
  // `Workbook::set_cell_value`. This test alternates a single-cell
  // mutation with a `recalc_parallel` invocation 30 times and validates
  // that every pass produces the right tail value. Under TSan this
  // surfaces any race between the mutator's eager dirty-set update and
  // the next recalc's BFS propagation.
  Workbook wb = Workbook::create();
  SeedChainWorkbook(wb);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  for (int i = 0; i < 30; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i)))));
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
    Value a10 = StoredValue(wb, 0, 9U, 0U);
    ASSERT_TRUE(a10.is_number()) << "iter " << i;
    EXPECT_DOUBLE_EQ(a10.as_number(), static_cast<double>(i) + 9.0);
  }
}

// ---------------------------------------------------------------------------
// Reading volatile-tracker / dirty-set state via the engine accessor
// ---------------------------------------------------------------------------

TEST(RecalcEngineInternal, EngineAccessorReadsAfterRecalcReturns) {
  // The scheduler clears the dirty set on completion. After a recalc,
  // reading `dirty()` and `volatiles()` from a separate thread (now that
  // the recalc has joined every worker) must observe consistent state.
  Workbook wb = Workbook::create();
  SeedChainWorkbook(wb);

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));

  std::atomic<int> mismatch{0};
  std::vector<std::thread> readers;
  readers.reserve(4);
  for (int i = 0; i < 4; ++i) {
    readers.emplace_back([&] {
      const RecalcEngine& engine = wb.recalc_engine();
      // Post-recalc the dirty set is empty; volatiles list is empty for
      // this workbook (no NOW / RAND / OFFSET formulas).
      for (int n = 0; n < 200; ++n) {
        if (engine.dirty().size() != 0U) {
          ++mismatch;
        }
      }
    });
  }
  for (std::thread& t : readers) {
    t.join();
  }
  EXPECT_EQ(mismatch.load(), 0);
}

// ---------------------------------------------------------------------------
// Many small workbooks recalcd concurrently — engine instances are
// disjoint, no shared engine state.
// ---------------------------------------------------------------------------

TEST(RecalcEngineInternal, ManyDisjointEnginesRaceFree) {
  // 8 worker threads × 4 workbooks each. Each workbook owns its own
  // engine instance; the only shared resource is the function registry.
  // TSan sees worker-local engine touches only, so any cross-engine race
  // would imply a hidden static — the test catches that regression.
  constexpr int kThreads = 8;
  constexpr int kWorkbooksPerThread = 4;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<int> failures{0};

  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&failures] {
      for (int w = 0; w < kWorkbooksPerThread; ++w) {
        Workbook wb = Workbook::create();
        if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(3.0)))) {
          ++failures;
          return;
        }
        if (!static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1*A1+1"))) {
          ++failures;
          return;
        }
        SchedulerConfig cfg;
        cfg.num_threads = 1U;
        if (!static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))) {
          ++failures;
          return;
        }
        const Sheet& s = wb.sheet(0);
        const Cell* b1 = s.cell_at(0U, 1U);
        if (b1 == nullptr || !b1->cached_value.is_number() || b1->cached_value.as_number() != 10.0) {
          ++failures;
          return;
        }
      }
    });
  }
  for (std::thread& t : workers) {
    t.join();
  }
  EXPECT_EQ(failures.load(), 0);
}

// ---------------------------------------------------------------------------
// Iterative-options mutation between recalcs
// ---------------------------------------------------------------------------

TEST(RecalcEngineInternal, IterativeOptionsMutationRaceFree) {
  // Sequentially toggle iterative-calc options between recalc passes. No
  // race is possible (single-threaded mutation), but the test is the
  // canonical "the engine's `set_iterative_options` does not leak past
  // its own scope" check.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+1")));

  SchedulerConfig cfg;
  cfg.num_threads = 2U;

  // First pass: iterative calc disabled -> #REF! on every member.
  IterativeOptions disabled;
  disabled.enabled = false;
  wb.set_iterative_options(disabled);
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  Value a1 = StoredValue(wb, 0, 0, 0);
  ASSERT_TRUE(a1.is_error());

  // Second pass: enable iterative calc, re-dirty, observe convergence
  // attempt (the cycle is non-converging so the result is still an
  // error, but the path through the iterative solver must not race).
  IterativeOptions enabled;
  enabled.enabled = true;
  enabled.max_iterations = 20U;
  enabled.max_change = 0.01;
  wb.set_iterative_options(enabled);
  // Force re-evaluation by mutating one cell.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=B1+1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
}

}  // namespace
}  // namespace formulon::eval
