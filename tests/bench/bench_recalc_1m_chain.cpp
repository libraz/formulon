//
// Microbenchmark: large linear chain recalc on the single-threaded engine.
//
// Builds a workbook whose first column carries a literal seed (`A1 = 1`)
// followed by a long chain `A(i+1) = A(i) + 1`. The recalc engine walks the
// chain from leaf to root in a single SCC pass; with no parallelism
// available (each cell is its own layer of size 1) the wall-clock time is
// dominated by per-cell tree-walker dispatch overhead.
//
// Target: 1,000,000 cells should recalc in <= 5 s wall clock.
//
// On developer laptops the bench's setup phase (per-cell `set_cell_formula`
// calls, each of which re-parses the formula and updates the dep graph) is
// significantly more expensive than the recalc itself. The setup is run
// once outside the timed region; only the `wb.recalc(...)` call is measured.
//
// Output: a single nanobench JSON document on stdout (when invoked with
// `--json <path>`), or a human-readable summary table otherwise.

#define ANKERL_NANOBENCH_IMPLEMENT
#include <nanobench.h>

#include <chrono>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <fstream>
#include <string>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "value.h"
#include "workbook.h"

namespace {

// Default chain length. The implementation plan calls for a 1,000,000-cell
// chain; the current tree-walk recalc engine runs at ~150 us/cell on a
// modern Apple-Silicon laptop in Release mode, which scales superlinearly
// because each cell propagates dirtiness through every dependent ahead of
// it (see `Workbook::set_cell_formula` -> `mark_dependents_dirty`). A 1M
// run also overflows the recursive Tarjan SCC walker's stack on macOS's
// default 8 MiB pthread stack — a 2,000-cell chain already crashes a
// Debug build, where stack frames are several times larger than Release.
//
// Until both are addressed (iterative Tarjan, batched dep-graph updates
// in `set_cell_formula`, bytecode VM dispatch on hot paths), we ship a
// 1,000-cell baseline. The default lands well below the Debug-mode
// stack-overflow threshold and gives a reproducible measurement that
// the regression gate can anchor against on either build configuration.
//
// To scale back up to the original 1M target once the engine improves,
// bump `kDefaultCellCount` and re-baseline.
constexpr std::uint32_t kDefaultCellCount = 1000U;

formulon::Workbook BuildChain(std::uint32_t n) {
  using formulon::Value;
  using formulon::Workbook;

  Workbook wb = Workbook::create();
  // Seed: A1 = 1.
  (void)wb.set_cell_value(0U, 0U, 0U, Value::number(1.0));
  // A2..An: each cell adds 1 to its predecessor.
  for (std::uint32_t r = 1; r < n; ++r) {
    std::string formula = "=A" + std::to_string(r) + "+1";
    (void)wb.set_cell_formula(0U, r, 0U, std::move(formula));
  }
  return wb;
}

// Builds formulas with disjoint one-cell dependencies. After the initial
// recalc, editing one formula gives the dirty-SCC benchmark a large workbook
// whose selected closure contains exactly that formula. The literal edge is
// intentional: completely constant formulas are omitted from the dependency
// graph's dirty-node set and would not exercise Tarjan at all.
formulon::Workbook BuildIndependentFormulas(std::uint32_t n) {
  using formulon::Value;
  using formulon::Workbook;

  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < n; ++r) {
    (void)wb.set_cell_value(0U, r, 1U, Value::number(1.0));
    (void)wb.set_cell_formula(0U, r, 0U, "=B" + std::to_string(r + 1U));
  }
  return wb;
}

}  // namespace

int main(int argc, char* argv[]) {
  std::uint32_t cell_count = kDefaultCellCount;
  std::string json_path;
  bool single_cell_edit = false;

  for (int i = 1; i < argc; ++i) {
    if (std::strcmp(argv[i], "--cells") == 0 && i + 1 < argc) {
      cell_count = static_cast<std::uint32_t>(std::strtoul(argv[i + 1], nullptr, 10));
      ++i;
    } else if (std::strcmp(argv[i], "--json") == 0 && i + 1 < argc) {
      json_path = argv[i + 1];
      ++i;
    } else if (std::strcmp(argv[i], "--single-cell-edit") == 0) {
      single_cell_edit = true;
    }
  }

  std::fprintf(stderr, "bench_recalc_1m_chain: building %s of %u cells...\n",
               single_cell_edit ? "independent formulas" : "chain", cell_count);
  if (single_cell_edit && cell_count == 0U) {
    std::fprintf(stderr, "bench_recalc_1m_chain: --single-cell-edit requires --cells >= 1\n");
    return 1;
  }
  formulon::Workbook wb = single_cell_edit ? BuildIndependentFormulas(cell_count) : BuildChain(cell_count);
  if (single_cell_edit) {
    auto initial = wb.recalc(formulon::eval::default_registry());
    if (!initial) {
      std::fprintf(stderr, "bench_recalc_1m_chain: initial recalc failed\n");
      return 1;
    }
    // The final formula has no dependents, so the next pass must select and
    // evaluate only that changed cell regardless of workbook size.
    (void)wb.set_cell_formula(0U, cell_count - 1U, 0U, "=B" + std::to_string(cell_count) + "+1");
  }
  std::fprintf(stderr, "bench_recalc_1m_chain: setup complete; running recalc...\n");

  if (single_cell_edit) {
    // A nanobench epoch re-invokes a sub-microsecond body many times; only
    // its first invocation would be dirty and later ones would measure a
    // no-op recalc. This mode instead takes exactly one wall-clock sample.
    const auto start = std::chrono::steady_clock::now();
    auto stats = wb.recalc(formulon::eval::default_registry());
    const auto elapsed = std::chrono::steady_clock::now() - start;
    if (!stats) {
      std::fprintf(stderr, "bench_recalc_1m_chain: recalc failed\n");
      return 1;
    }
    if (stats.value().cells_evaluated != 1U) {
      std::fprintf(stderr, "bench_recalc_1m_chain: expected one evaluated cell, got %llu\n",
                   static_cast<unsigned long long>(stats.value().cells_evaluated));
      return 1;
    }
    std::fprintf(stderr, "bench_recalc_1m_chain: one-cell edit among %u formulas: %.3f ms\n", cell_count,
                 std::chrono::duration<double, std::milli>(elapsed).count());
    return 0;
  }

  ankerl::nanobench::Bench bench;
  bench.title("recalc_1m_chain").unit("recalc").warmup(0).epochs(1).minEpochIterations(1).relative(true);

  bench.run("recalc 1M-cell linear chain", [&]() {
    auto stats = wb.recalc(formulon::eval::default_registry());
    ankerl::nanobench::doNotOptimizeAway(stats);
  });

  if (!json_path.empty()) {
    std::ofstream out(json_path);
    bench.render(ankerl::nanobench::templates::json(), out);
  }

  return 0;
}
