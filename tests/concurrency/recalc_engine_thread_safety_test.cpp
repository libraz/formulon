//
// ThreadSanitizer-targeted tests for `RecalcEngine`'s public-API thread
// safety. The engine's internal mutex is supposed to serialise every
// mutating entry (`register_formula`, `unregister_formula`,
// `clear_cell_dependencies`, `mark_dirty`, the `recalc*` family) against
// every other mutating entry, even when callers race them from disjoint
// threads.
//
// These tests do NOT exercise the documented "no concurrent recalc on
// the same workbook" invariant — that contract is the caller's
// responsibility — but they DO cover the lower-level guarantee that the
// engine's bookkeeping survives an adversarial caller that races
// `register_formula` / `mark_dirty` from multiple threads. The lock
// keeps the dep graph and dirty set internally consistent so a future
// API surface that exposes parallel mutation paths cannot regress
// silently.
//
// Detached threads are forbidden by project policy; every thread is
// joined.

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <string>
#include <thread>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Helper: a workbook with N rows in column A laid out as
//   A1 = 1, A2 = A1+1, ..., AN = A(N-1)+1.
// Sized small enough to keep the TSan event budget low while still
// producing a non-trivial dep graph with O(N) edges.
constexpr std::uint32_t kChainRows = 32U;

void SeedChain(Workbook& wb) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  for (std::uint32_t r = 1U; r < kChainRows; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, r, 0U, "=A" + std::to_string(r) + "+1")));
  }
}

// ---------------------------------------------------------------------------
// Concurrent register_formula / mark_dirty must not crash or corrupt state.
// ---------------------------------------------------------------------------

TEST(RecalcEngineThreadSafety, ConcurrentMutatorsRaceFreeUnderLock) {
  // Two writer threads hammer the workbook's mutating recalc-engine
  // entry points in parallel: one keeps re-issuing `set_cell_formula`
  // (which routes through `register_formula` + `mark_dirty`), the other
  // keeps re-issuing `set_cell_value` on a sibling cell (which routes
  // through `mark_dirty` for every dependent). With the engine mutex in
  // place these calls serialise cleanly; without it the dep graph and
  // dirty set would race and TSan would flag the unsynchronised access.
  Workbook wb = Workbook::create();
  SeedChain(wb);

  std::atomic<bool> stop{false};
  std::atomic<int> failures{0};

  // Writer 1: rewrites A2's formula in a tight loop. Each rewrite drops
  // the cell's previous outgoing edges and adds the new one, exercising
  // both `unregister_formula_locked` and `register_formula_locked`.
  std::thread writer_a([&] {
    for (int i = 0; !stop.load(std::memory_order_acquire) && i < 200; ++i) {
      const std::string formula = (i % 2 == 0) ? "=A1+1" : "=A1+2";
      if (!static_cast<bool>(wb.set_cell_formula(0U, 1U, 0U, formula))) {
        ++failures;
        return;
      }
    }
  });

  // Writer 2: re-stamps the seed cell so every dependent gets dirtied
  // through the BFS in `set_cell_value`. The volatile / dirty bookkeeping
  // is the most TSan-visible mutable state during the race.
  std::thread writer_b([&] {
    for (int i = 0; !stop.load(std::memory_order_acquire) && i < 200; ++i) {
      if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i + 1))))) {
        ++failures;
        return;
      }
    }
  });

  writer_a.join();
  writer_b.join();
  stop.store(true, std::memory_order_release);
  EXPECT_EQ(failures.load(), 0);

  // Drive a final recalc and confirm the engine state is still healthy.
  // We do not assert specific values because the writers raced — only
  // that the post-race recalc completes successfully.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
}

// ---------------------------------------------------------------------------
// Concurrent recalc + read-only dep_graph() observer.
// ---------------------------------------------------------------------------

TEST(RecalcEngineThreadSafety, ParallelRecalcSerialisesAgainstReaderObserver) {
  // While `recalc_parallel` is in flight the engine mutex is held for
  // the entire pass. A separate observer thread that calls into the
  // mutating API (`mark_dirty` on any of the chain's cells) should
  // block until the recalc returns; once it does, the dep graph state
  // remains coherent and a subsequent recalc still succeeds.
  Workbook wb = Workbook::create();
  SeedChain(wb);

  // Initial recalc to commit values.
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));

  // Re-dirty the chain by mutating the seed.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(2.0))));

  std::atomic<int> failures{0};
  std::atomic<bool> observer_done{false};

  // Observer thread: races a `set_cell_value` against the parallel
  // recalc on the main thread. The lock makes this a happens-before
  // edge instead of a race; either the observer's mutation lands
  // before the recalc's lock acquisition (and the recalc picks it up)
  // or after (and a follow-up recalc would). Either ordering must be
  // race-free under TSan.
  std::thread observer([&] {
    if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::number(99.0)))) {
      ++failures;
    }
    observer_done.store(true, std::memory_order_release);
  });

  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
  observer.join();
  EXPECT_TRUE(observer_done.load());
  EXPECT_EQ(failures.load(), 0);

  // Final recalc must succeed regardless of the observer's interleaving.
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));
}

// ---------------------------------------------------------------------------
// Concurrent independent workbooks: per-workbook mutex must not contend.
// ---------------------------------------------------------------------------

TEST(RecalcEngineThreadSafety, IndependentWorkbookEnginesIndependentLocks) {
  // Eight threads each own a workbook and drive a tight
  // mutate-recalc-mutate loop. Per-workbook mutexes mean these never
  // contend; TSan must observe no race even though the total mutation
  // rate is high.
  constexpr int kThreads = 8;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<int> failures{0};

  for (int t = 0; t < kThreads; ++t) {
    workers.emplace_back([&failures, t] {
      Workbook wb = Workbook::create();
      if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(t + 1))))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=A1+A1"))) {
        ++failures;
        return;
      }
      for (int i = 0; i < 25; ++i) {
        if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i))))) {
          ++failures;
          return;
        }
        if (!static_cast<bool>(wb.recalc(default_registry()))) {
          ++failures;
          return;
        }
      }
    });
  }
  for (std::thread& th : workers) {
    th.join();
  }
  EXPECT_EQ(failures.load(), 0);
}

}  // namespace
}  // namespace formulon::eval
