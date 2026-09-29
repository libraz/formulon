//
// Registration-scale guard for compact range dependencies.
//
// A lookup dragged down a column is the shape that turns a per-cell
// dependency graph into hundreds of megabytes: one
// `VLOOKUP(A2, Sheet2!$A$1:$F$50000, 3, FALSE)` covers 300,000 cells, and a
// thousand copies of it would be 300 million permanently resident edges
// across the graph's three indexes. Registered as one interned rectangle the
// same workbook contributes no range-derived node at all, so the graph
// footprint has to stay proportional to the formulas themselves and
// registration has to finish in seconds. Loading a workbook whose formulas
// watch a whole populated column is held to the same linear bar.
//
// Lives in its own `SLOW`-labelled executable: a thousand registrations over
// a large rectangle is far above the fast tier's per-case budget.

#include <chrono>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

TEST(RecalcScaleSlow, DraggedLookupOverLargeTableKeepsGraphFootprintFlat) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");

  constexpr std::uint32_t kFormulaCount = 1000U;
  const auto started = std::chrono::steady_clock::now();
  for (std::uint32_t row = 0; row < kFormulaCount; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(row + 1.0))));
    ASSERT_TRUE(static_cast<bool>(
        wb.set_cell_formula(0U, row, 1U, "=VLOOKUP(A" + std::to_string(row + 1U) + ",Sheet2!$A$1:$F$50000,3,FALSE)")));
  }
  const auto elapsed = std::chrono::steady_clock::now() - started;

  // Two nodes per formula: the lookup key it reads and the formula itself.
  // The 300,000-cell rectangle contributes none.
  EXPECT_EQ(wb.recalc_engine().dep_graph().node_count(), 2U * kFormulaCount);
  EXPECT_LT(std::chrono::duration_cast<std::chrono::seconds>(elapsed).count(), 10);
}

// Every row pairs a literal with a formula watching the whole literal column,
// so the reader interleaves them: each literal lands under N already
// registered watchers and each formula's rectangle spans every literal
// loaded so far. Load has to stay linear in the row count on both counts.
TEST(RecalcScaleSlow, LoadingWholeColumnWatchersStaysLinear) {
  constexpr std::uint32_t kRows = 40000U;
  std::vector<std::uint8_t> bytes;
  {
    Workbook wb = Workbook::create();
    // Literals first so building the fixture through the live API does not
    // itself pay the per-write watcher walk being measured on load.
    for (std::uint32_t row = 0; row < kRows; ++row) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(row + 1.0))));
    }
    for (std::uint32_t row = 0; row < kRows; ++row) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=ROWS(A:A)")));
    }
    auto written = io::write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(written));
    bytes = std::move(written.value());
  }

  const auto started = std::chrono::steady_clock::now();
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  const auto elapsed = std::chrono::steady_clock::now() - started;
  ASSERT_TRUE(static_cast<bool>(loaded));
  EXPECT_LT(std::chrono::duration_cast<std::chrono::seconds>(elapsed).count(), 10);

  Workbook& wb = loaded.value().workbook;
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* last = wb.sheet(0U).cell_at(kRows - 1U, 1U);
  ASSERT_NE(last, nullptr);
  ASSERT_TRUE(last->cached_value.is_number());
  EXPECT_DOUBLE_EQ(last->cached_value.as_number(), 1048576.0);
}

}  // namespace
}  // namespace formulon
