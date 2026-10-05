#pragma once

#include <atomic>
#include <chrono>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <set>
#include <string>
#include <thread>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "eval/volatile_tracker.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/thread_launch.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval::scheduler_test {

/// Clears any injected thread-launch failure when the test leaves scope.
/// Without this an assertion that aborts a degradation test mid-way would
/// leave every later test unable to start a worker.
struct ThreadLaunchInjection {
  ~ThreadLaunchInjection() { clear_thread_launch_failure_injection(); }
};

inline Value StoredValue(const Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
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

inline bool RecordIterativeProgressThread(std::uint32_t /*iteration*/, double /*max_residual*/,
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

inline WorkerThreadProbe g_worker_thread_probe;

inline void ResetWorkerThreadProbe(bool synchronize) {
  std::lock_guard<std::mutex> guard(g_worker_thread_probe.mutex);
  g_worker_thread_probe.ids.clear();
  g_worker_thread_probe.synchronize = synchronize;
  g_worker_thread_probe.released = !synchronize;
  g_worker_thread_probe.timed_out = false;
}

inline Value RecordWorkerThreadImpl(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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

inline Workbook* g_worker_reentry_workbook = nullptr;
inline std::atomic<int> g_worker_reentry_code{0};

inline Value WorkerReentryImpl(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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

inline bool AbortIterativeSolve(std::uint32_t /*iteration*/, double /*max_residual*/, std::uint32_t /*max_iterations*/,
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
inline void RecalcBothAndExpectEqual(WorkbookPair& wp, std::uint32_t threads = 4U) {
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

}  // namespace formulon::eval::scheduler_test
