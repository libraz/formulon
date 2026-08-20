//
// ThreadSanitizer-targeted tests for the cells a parallel recalc pass reads
// without a dependency edge to justify the read.
//
// `INDIRECT` and `OFFSET` compute their target at evaluation time, so
// `dep_extractor` records no edge for it. Nothing in the condensed graph
// then separates such a formula from the cell it is about to read, and the
// two land in the same topological layer — one worker resolving the
// reference while another commits the target. The fixtures below build
// exactly that shape and assert three things about it:
//
//   * the layer really is shared (the stats pin the topology, so the test
//     cannot quietly stop covering the case);
//   * the parallel result equals the single-threaded engine's and repeats
//     identically across passes;
//   * Text targets are included, because those are the ones where the
//     target's `cached_text_owned` allocation is replaced under the reader.
//
// Every target formula is chosen to recompute to the value it already had,
// so an ordering difference cannot hide behind a value difference: the
// only thing left for a scheduling bug to produce is a sanitizer report.
//
// Detached threads are forbidden by project policy; the scheduler joins
// every worker it starts.

#include <cstdint>
#include <string>
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

// Rows of the dynamic-reference block. Wide enough that the pooled half of
// the layer holds several tasks per worker, small enough to keep the TSan
// event budget bounded.
constexpr std::uint32_t kRows = 96U;

// Column layout of the fixture.
constexpr std::uint32_t kTargetCol = 0U;    // A: the cells being rewritten.
constexpr std::uint32_t kIndirectCol = 1U;  // B: =INDIRECT("A<n>")
constexpr std::uint32_t kOffsetCol = 2U;    // C: =OFFSET($D$1, <n>, -3)
constexpr std::uint32_t kAnchorCol = 3U;    // D: literal OFFSET base.

// Long enough that the target's Text payload is a heap buffer rather than
// a small-string one, so replacing it is a free the reader could observe.
std::string TargetText(std::uint32_t row) {
  return "formulon-parallel-recalc-row-" + std::to_string(row);
}

Value StoredValue(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Sheet& sheet = wb.sheet(0U);
  if (const Cell* cell = sheet.cell_at(row, col); cell != nullptr) {
    return cell->cached_value;
  }
  return Value::blank();
}

// Populates `wb` with the shared-layer fixture:
//
//   D1        literal, the OFFSET base. Never dirty, so the OFFSET cells
//             keep an in-degree of zero inside the dirty subgraph and stay
//             in the same layer as the cells they actually read.
//   A1..An    `=UPPER("...")`, a stable Text result that is nevertheless
//             recomputed (and re-stored) on every pass.
//   B1..Bn    `=INDIRECT("A<n>")` — reads A<n> with no edge to it.
//   C1..Cn    `=OFFSET($D$1, <n-1>, -3)` — likewise.
void SeedDynamicRefWorkbook(Workbook& wb) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, kAnchorCol, Value::number(0.0))));
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string row_1based = std::to_string(row + 1U);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, kTargetCol, "=UPPER(\"" + TargetText(row) + "\")")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, kIndirectCol, "=INDIRECT(\"A" + row_1based + "\")")));
    ASSERT_TRUE(
        static_cast<bool>(wb.set_cell_formula(0U, row, kOffsetCol, "=OFFSET($D$1," + std::to_string(row) + ",-3)")));
  }
}

// Re-dirties the target column. The dynamic readers are volatile and seed
// themselves; without this the targets would be clean after the first pass
// and the shared layer would evaporate.
void MarkTargetsDirty(Workbook& wb) {
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    wb.recalc_engine().mark_dirty(CellNodeId{0U, row, kTargetCol});
  }
}

std::string ExpectedText(std::uint32_t row) {
  std::string upper = TargetText(row);
  for (char& c : upper) {
    if (c >= 'a' && c <= 'z') {
      c = static_cast<char>(c - 'a' + 'A');
    }
  }
  return upper;
}

void ExpectFixtureResolved(const Workbook& wb) {
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string expected = ExpectedText(row);
    const Value target = StoredValue(wb, row, kTargetCol);
    ASSERT_TRUE(target.is_text()) << "target row " << row;
    EXPECT_EQ(target.as_text(), expected) << "target row " << row;

    const Value indirect = StoredValue(wb, row, kIndirectCol);
    ASSERT_TRUE(indirect.is_text()) << "INDIRECT row " << row;
    EXPECT_EQ(indirect.as_text(), expected) << "INDIRECT row " << row;

    const Value offset = StoredValue(wb, row, kOffsetCol);
    ASSERT_TRUE(offset.is_text()) << "OFFSET row " << row;
    EXPECT_EQ(offset.as_text(), expected) << "OFFSET row " << row;
  }
}

// The topology assertion. Once the seeding pass has settled the literal
// anchor, every super-node of the fixture is a singleton whose in-degree
// inside the dirty subgraph is zero, so a pass is exactly one layer
// holding all `3 * kRows` of them: the targets dispatched to the pool, the
// volatile readers run afterwards on the caller. If a future change
// separated the readers from their targets into different layers this
// fixture would stop covering the race, and these counters are what would
// say so.
void ExpectSingleSharedLayer(const SchedulerStats& stats) {
  EXPECT_EQ(stats.sccs_processed, 3U * kRows);
  EXPECT_EQ(stats.parallel_steps, 1U);
  EXPECT_EQ(stats.serial_fallback_steps, 1U);
  EXPECT_GE(stats.worker_threads_started, 2U);
}

TEST(ParallelDynamicRef, SharedLayerMatchesSerialRecalc) {
  Workbook serial = Workbook::create();
  Workbook parallel = Workbook::create();
  ASSERT_NO_FATAL_FAILURE(SeedDynamicRefWorkbook(serial));
  ASSERT_NO_FATAL_FAILURE(SeedDynamicRefWorkbook(parallel));

  SchedulerConfig cfg;
  cfg.num_threads = 4U;

  // Settling pass. Neither engine orders a dynamic reference against the
  // cell it reads — no edge exists to order it by — so a reader's first
  // pass may legitimately precede its target's. What both engines must
  // agree on is the steady state, which is what the measured pass below
  // observes: the targets are re-evaluated to the values they already
  // hold, so the readers see the same bytes whichever side of the write
  // they land on.
  ASSERT_TRUE(static_cast<bool>(serial.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(parallel.recalc_parallel(default_registry(), cfg, nullptr)));

  MarkTargetsDirty(serial);
  MarkTargetsDirty(parallel);

  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(serial.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(parallel.recalc_parallel(default_registry(), cfg, &stats)));

  ExpectSingleSharedLayer(stats);
  ASSERT_NO_FATAL_FAILURE(ExpectFixtureResolved(serial));

  for (std::uint32_t row = 0U; row < kRows; ++row) {
    for (const std::uint32_t col : {kTargetCol, kIndirectCol, kOffsetCol, kAnchorCol}) {
      EXPECT_EQ(StoredValue(serial, row, col), StoredValue(parallel, row, col))
          << "value mismatch at (" << row << ", " << col << ")";
    }
  }
}

TEST(ParallelDynamicRef, RepeatedPassesAreDeterministic) {
  Workbook wb = Workbook::create();
  ASSERT_NO_FATAL_FAILURE(SeedDynamicRefWorkbook(wb));

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, nullptr)));

  for (std::uint32_t pass = 0U; pass < 12U; ++pass) {
    MarkTargetsDirty(wb);
    SchedulerStats stats;
    ASSERT_TRUE(static_cast<bool>(wb.recalc_parallel(default_registry(), cfg, &stats)));
    ExpectSingleSharedLayer(stats);
    ASSERT_NO_FATAL_FAILURE(ExpectFixtureResolved(wb)) << "pass " << pass;
  }
}

}  // namespace
}  // namespace formulon::eval
