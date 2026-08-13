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
// registration has to finish in seconds.
//
// Lives in its own `SLOW`-labelled executable: a thousand registrations over
// a large rectangle is far above the fast tier's per-case budget.

#include <chrono>
#include <cstdint>
#include <string>

#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
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

}  // namespace
}  // namespace formulon
