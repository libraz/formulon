// Copyright 2026 libraz. Licensed under the MIT License.
//
// ThreadSanitizer-targeted tests for the function registry's read paths.
//
// `default_registry()` is a Meyers singleton; C++11 mandates thread-safe
// initialisation, so the first-use race between concurrent callers must
// be handled by the language runtime, not the engine. The tests here
// verify two layers of that guarantee:
//
//   * The `default_registry()` accessor itself: 8 threads invoke it
//     simultaneously and read back its `size()`. All threads observe the
//     same fully-populated registry.
//   * Concurrent `lookup()` calls: with the registry already initialised,
//     8 threads each look up several thousand random function names. The
//     map is read-only post-init, so this exercises only the reader-
//     reader sharing path.
//   * Concurrent lazy dispatch: `find_lazy()` (the linear scan inside
//     tree_walker.cpp) is read-only; no separate test is required because
//     it has no shared mutable state, but we still drive an integration-
//     level `recalc_parallel` workload that hits the lazy path from
//     multiple threads to surface any future regression.
//
// The first-use test cannot prove the singleton races safely against
// initialisation in the literal sense — the singleton is already
// initialised by the time gtest enters `main()` (every other test in the
// binary touches it). That is fine: the language guarantee, the
// reader/reader test below, and the integration-level lazy-dispatch
// drive cover the practical failure modes.

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <random>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// A static list of well-known function names, sampled randomly by the
// concurrent-lookup test. Mixing hits and misses keeps the lookup loop
// honest (it must walk the whole hash bucket chain for each).
constexpr std::string_view kKnownNames[] = {
    "SUM",     "AVERAGE",  "COUNT",   "MAX",         "MIN",   "ABS",         "ROUND",  "POWER",    "SQRT",
    "EXP",     "LN",       "LOG",     "PI",          "RAND",  "RANDBETWEEN", "IF",     "IFERROR",  "IFNA",
    "AND",     "OR",       "NOT",     "TRUE",        "FALSE", "LEN",         "UPPER",  "LOWER",    "TRIM",
    "LEFT",    "RIGHT",    "MID",     "CONCATENATE", "DATE",  "YEAR",        "MONTH",  "DAY",      "TODAY",
    "NOW",     "WEEKDAY",  "VLOOKUP", "HLOOKUP",     "INDEX", "MATCH",       "OFFSET", "INDIRECT", "ISBLANK",
    "ISERROR", "ISNUMBER", "ISTEXT",  "ISLOGICAL",
};

constexpr std::string_view kUnknownNames[] = {
    "DEFINITELY_NOT_A_FUNCTION", "ZZZ_PHANTOM", "FOO", "BAR_BAZ", "EXCEL_DOES_NOT_HAVE_THIS",
};

// ---------------------------------------------------------------------------
// First-use access from multiple threads
// ---------------------------------------------------------------------------

TEST(RegistryThreadSafety, FirstUseFromMultipleThreadsObservesSameRegistry) {
  // Every thread fetches `default_registry()` and reads `size()` plus a
  // canary lookup. Under TSan the C++11 thread-safe-init machinery is
  // expected to suppress any data race on the singleton's underlying
  // storage. Note that by the time gtest runs the binary's startup code
  // has likely already constructed the singleton; this test still
  // exercises the "many concurrent readers reading the same instance"
  // path.
  constexpr int kThreads = 8;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<std::size_t> first_size{0};
  std::atomic<int> mismatch_count{0};

  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&] {
      const FunctionRegistry& reg = default_registry();
      const std::size_t sz = reg.size();
      // Compare-exchange the first observed size; subsequent threads must
      // see the same number.
      std::size_t expected = 0U;
      if (first_size.compare_exchange_strong(expected, sz)) {
        return;
      }
      if (first_size.load() != sz) {
        ++mismatch_count;
      }
      // Canary lookup: SUM must always be present.
      if (reg.lookup("SUM") == nullptr) {
        ++mismatch_count;
      }
    });
  }
  for (std::thread& t : workers) {
    t.join();
  }
  EXPECT_GT(first_size.load(), 0U);
  EXPECT_EQ(mismatch_count.load(), 0);
}

// ---------------------------------------------------------------------------
// Concurrent reads on the populated registry
// ---------------------------------------------------------------------------

TEST(RegistryThreadSafety, EightThreadsLookupOneThousandNamesEach) {
  // 8 threads each perform 1000 lookups against the global registry. A
  // subset of the names are guaranteed hits; another subset are
  // guaranteed misses. Both paths walk the underlying hash map, which
  // must remain race-free for concurrent readers.
  const FunctionRegistry& reg = default_registry();
  ASSERT_GT(reg.size(), 0U);

  constexpr int kThreads = 8;
  constexpr int kLookupsPerThread = 1000;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<std::uint64_t> hits{0};
  std::atomic<std::uint64_t> misses{0};

  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&, seed = static_cast<std::uint32_t>(i + 1)] {
      std::mt19937 rng(seed);
      std::uniform_int_distribution<std::size_t> known_pick(0U, sizeof(kKnownNames) / sizeof(kKnownNames[0]) - 1U);
      std::uniform_int_distribution<std::size_t> unknown_pick(0U,
                                                              sizeof(kUnknownNames) / sizeof(kUnknownNames[0]) - 1U);
      std::uniform_int_distribution<int> bucket(0, 9);
      for (int n = 0; n < kLookupsPerThread; ++n) {
        if (bucket(rng) < 7) {
          // 70% known names.
          std::string_view name = kKnownNames[known_pick(rng)];
          if (reg.lookup(name) != nullptr) {
            ++hits;
          }
        } else {
          // 30% unknown names.
          std::string_view name = kUnknownNames[unknown_pick(rng)];
          if (reg.lookup(name) == nullptr) {
            ++misses;
          }
        }
      }
    });
  }
  for (std::thread& t : workers) {
    t.join();
  }
  // Sanity: at least some hits and misses occurred.
  EXPECT_GT(hits.load(), 0U);
  EXPECT_GT(misses.load(), 0U);
}

// ---------------------------------------------------------------------------
// Lookup case-insensitivity from multiple threads
// ---------------------------------------------------------------------------

TEST(RegistryThreadSafety, CaseInsensitiveLookupRaceFree) {
  // The lookup path uppercases its input before hashing. Concurrent
  // callers therefore each produce an independent local string, which
  // means the only shared state is the underlying hash map. This test
  // exercises mixed-case spellings to exercise the uppercase fast path
  // for every thread.
  const FunctionRegistry& reg = default_registry();
  constexpr int kThreads = 8;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<std::uint64_t> hits{0};

  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&] {
      static const char* const kCases[] = {"sum", "Sum", "SUM", "sUm", "SUm", "sUM"};
      for (int n = 0; n < 500; ++n) {
        for (const char* spelling : kCases) {
          if (reg.lookup(spelling) != nullptr) {
            ++hits;
          }
        }
      }
    });
  }
  for (std::thread& t : workers) {
    t.join();
  }
  EXPECT_EQ(hits.load(), static_cast<std::uint64_t>(kThreads * 500 * 6));
}

// ---------------------------------------------------------------------------
// Lazy dispatch path from multiple threads
// ---------------------------------------------------------------------------

TEST(RegistryThreadSafety, LazyDispatchPathDriverFromEightThreads) {
  // Lazy-dispatched functions (IF / IFERROR / IFNA / SUMIF / ...) are
  // resolved by `find_lazy()` inside `tree_walker.cpp`. The function is
  // a linear scan over a static-storage array literal — no shared
  // mutable state — but driving it under load from 8 threads keeps any
  // future regression honest (e.g. someone introducing a cached pointer).
  //
  // Each thread owns its own workbook so the engine layer holds no
  // cross-thread state.
  constexpr int kThreads = 8;
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  std::atomic<int> failures{0};

  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&failures, i] {
      Workbook wb = Workbook::create();
      // Cells exercising several lazy functions: IF, IFERROR, SUMIF.
      if (!static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(static_cast<double>(i + 1))))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(2.0)))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, "=IF(A1>0,A1*2,-1)"))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_formula(0U, 1U, 1U, "=IFERROR(A2/A1,0)"))) {
        ++failures;
        return;
      }
      if (!static_cast<bool>(wb.set_cell_formula(0U, 2U, 1U, "=SUMIF(A1:A2,\">0\")"))) {
        ++failures;
        return;
      }

      SchedulerConfig cfg;
      cfg.num_threads = 1U;  // Already running concurrently across workbooks.
      for (int pass = 0; pass < 5; ++pass) {
        if (!static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr))) {
          ++failures;
          return;
        }
      }

      // Spot-check expected results.
      const Sheet& s = wb.sheet(0);
      const Cell* b1 = s.cell_at(0U, 1U);
      if (b1 == nullptr || !b1->cached_value.is_number() ||
          b1->cached_value.as_number() != static_cast<double>(2 * (i + 1))) {
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
// Registry size is stable post-init
// ---------------------------------------------------------------------------

TEST(RegistryThreadSafety, SizeIsStableAcrossManyConcurrentReaders) {
  // Eight threads each call `size()` 5000 times. The reported value must
  // be identical across every observation: the registry is finalised
  // before the first call and never mutated.
  const FunctionRegistry& reg = default_registry();
  const std::size_t expected = reg.size();
  ASSERT_GT(expected, 0U);

  constexpr int kThreads = 8;
  constexpr int kIterations = 5000;
  std::atomic<int> mismatches{0};
  std::vector<std::thread> workers;
  workers.reserve(kThreads);
  for (int i = 0; i < kThreads; ++i) {
    workers.emplace_back([&] {
      for (int n = 0; n < kIterations; ++n) {
        if (reg.size() != expected) {
          ++mismatches;
        }
      }
    });
  }
  for (std::thread& t : workers) {
    t.join();
  }
  EXPECT_EQ(mismatches.load(), 0);
}

}  // namespace
}  // namespace formulon::eval
