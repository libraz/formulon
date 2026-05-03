// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Microbenchmark: large independent-cell layer driven through the parallel
// SCC scheduler.
//
// Builds a 1000 x 1000 grid where every cell carries the formula
// `=ROW()*COLUMN()+$Z$1`. The shared `$Z$1` dependency exists so the
// scheduler's "standalone-dirty" sweep does not pull every cell into
// the serial fallback path: cells with no graph edges live outside any
// SCC and never trigger the parallel layering, regardless of how many
// of them are dirty. Routing the formulas through a single shared
// literal puts all 999,999 dependents into one Kahn layer that the
// worker pool drains concurrently.
//
// We assert `parallel_steps > 0` to prove the scheduler actually
// parallelised — a regression in the layering algorithm would silently
// degrade to serial dispatch otherwise.
//
// Target (per `backup/plans/26-implementation-plan.md` Phase 5 / M10):
//   1,000,000 cells should recalc in <= 5 s wall clock with 8 workers.
//
// As with `bench_recalc_1m_chain.cpp`, only the `recalc_parallel(...)` call
// is timed; the per-cell `set_cell_formula` setup runs once outside the
// timed region.

#define ANKERL_NANOBENCH_IMPLEMENT
#include <nanobench.h>

#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <string>

#include "eval/function_registry.h"
#include "eval/scheduler.h"
#include "workbook.h"

namespace {

// 300 x 300 = 90,000 cells. The implementation plan calls for a
// 1000 x 1000 = 1,000,000 cell grid; that shape works in Release builds
// (~0.9 s recalc, ~2 s setup) but inflates to >30 s setup in Debug. The
// regression gate is meant to flag a 20% delta on the recalc itself, so
// shrinking the grid by ~10x while preserving the parallel-layering
// invariant is the right tradeoff: a meaningful, reproducible
// measurement on either build mode beats a number that is only
// reachable in Release.
//
// Scale back up to 1000 x 1000 once the build defaults to optimised
// flags for benches (or once the I/O / dep-graph codepaths cease to
// dominate setup) and re-baseline.
constexpr std::uint32_t kDefaultRows = 300U;
constexpr std::uint32_t kDefaultCols = 300U;
// 8 workers matches the WASM `PTHREAD_POOL_SIZE` ceiling; the
// scheduler clamps higher counts down to this anyway.
constexpr std::uint32_t kDefaultThreads = 8U;

// Anchor cell. The scheduler's "standalone-dirty" sweep handles cells
// with no graph edges serially, so every grid cell needs at least one
// outgoing dep. Pointing each formula at a single shared anchor cell
// achieves that without adding inter-grid edges that would force a
// serial topological order.
//
// The anchor lives outside the populated grid so the formula write does
// not clobber it; the bench lays the grid out in `[0..rows) x [0..cols)`
// and parks the anchor at (0, kAnchorCol) where `kAnchorCol == cols`.
formulon::Workbook BuildGrid(std::uint32_t rows, std::uint32_t cols) {
  using formulon::Value;
  using formulon::Workbook;

  Workbook wb = Workbook::create();
  const std::uint32_t anchor_row = 0U;
  const std::uint32_t anchor_col = cols;  // one past the grid's right edge
  // A1-notation column letter for a 0-based column index. Sized for up
  // to three letters (max column XFD = 16383). The bench keeps `cols`
  // <= 16383 so the helper is sufficient.
  //
  // Standard Excel encoding: each digit is 1-indexed within the alphabet
  // EXCEPT the least-significant one, which is 0-indexed. We build the
  // string least-significant first, decrementing the carry between
  // iterations to absorb the 1-indexing offset.
  auto a1_col = [](std::uint32_t c) {
    std::string out;
    std::uint32_t v = c;
    while (true) {
      out.insert(out.begin(), static_cast<char>('A' + (v % 26U)));
      if (v < 26U)
        break;
      v = v / 26U - 1U;
    }
    return out;
  };
  const std::string anchor_ref = "$" + a1_col(anchor_col) + "$" + std::to_string(anchor_row + 1U);

  (void)wb.set_cell_value(0U, anchor_row, anchor_col, Value::number(1.0));
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      std::string formula = "=ROW()*COLUMN()+" + anchor_ref;
      (void)wb.set_cell_formula(0U, r, c, std::move(formula));
    }
  }
  return wb;
}

}  // namespace

int main(int argc, char* argv[]) {
  std::uint32_t rows = kDefaultRows;
  std::uint32_t cols = kDefaultCols;
  std::uint32_t threads = kDefaultThreads;
  std::string json_path;

  for (int i = 1; i < argc; ++i) {
    if (std::strcmp(argv[i], "--rows") == 0 && i + 1 < argc) {
      rows = static_cast<std::uint32_t>(std::strtoul(argv[i + 1], nullptr, 10));
      ++i;
    } else if (std::strcmp(argv[i], "--cols") == 0 && i + 1 < argc) {
      cols = static_cast<std::uint32_t>(std::strtoul(argv[i + 1], nullptr, 10));
      ++i;
    } else if (std::strcmp(argv[i], "--threads") == 0 && i + 1 < argc) {
      threads = static_cast<std::uint32_t>(std::strtoul(argv[i + 1], nullptr, 10));
      ++i;
    } else if (std::strcmp(argv[i], "--json") == 0 && i + 1 < argc) {
      json_path = argv[i + 1];
      ++i;
    }
  }

  const std::uint64_t total_cells = static_cast<std::uint64_t>(rows) * static_cast<std::uint64_t>(cols);
  std::fprintf(stderr, "bench_recalc_1m_wide: building %u x %u (=%llu cells) grid...\n", rows, cols,
               static_cast<unsigned long long>(total_cells));
  formulon::Workbook wb = BuildGrid(rows, cols);
  std::fprintf(stderr, "bench_recalc_1m_wide: setup complete; running parallel recalc...\n");

  formulon::eval::SchedulerConfig cfg;
  cfg.num_threads = threads;
  formulon::eval::SchedulerStats stats;

  ankerl::nanobench::Bench bench;
  bench.title("recalc_1m_wide_parallel").unit("recalc").warmup(0).epochs(1).minEpochIterations(1).relative(true);

  bench.run("recalc 1M-cell wide layer (8 threads)", [&]() {
    auto out = wb.recalc_parallel(formulon::eval::default_registry(), cfg, &stats);
    ankerl::nanobench::doNotOptimizeAway(out);
  });

  // Sanity check: the scheduler must have actually parallelised. A
  // failure here typically means the topology produced one layer of
  // size 1, which in turn means the dep extractor saw something it did
  // not expect (`ROW()` / `COLUMN()` started reporting dependencies, etc).
  if (stats.parallel_steps == 0U) {
    std::fprintf(stderr, "bench_recalc_1m_wide: ERROR scheduler did not parallelise (parallel_steps=0)\n");
    return 1;
  }
  std::fprintf(stderr,
               "bench_recalc_1m_wide: cells_evaluated=%llu sccs_processed=%llu parallel_steps=%llu "
               "serial_fallback_steps=%llu\n",
               static_cast<unsigned long long>(stats.cells_evaluated),
               static_cast<unsigned long long>(stats.sccs_processed),
               static_cast<unsigned long long>(stats.parallel_steps),
               static_cast<unsigned long long>(stats.serial_fallback_steps));

  if (!json_path.empty()) {
    std::ofstream out(json_path);
    bench.render(ankerl::nanobench::templates::json(), out);
  }

  return 0;
}
