// Concurrency scheduler tests grouped by fixture size and execution mode.
// Shared probes and workbook-pair helpers live in scheduler_test_helpers.h.

#include <random>

#include "scheduler_test_helpers.h"

namespace formulon::eval {
namespace {

using namespace scheduler_test;

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
