//
// ThreadSanitizer-targeted test for the error-detail side record: parallel
// recalc workers commit details for many error cells on one sheet at once,
// each worker classifying through its own thread-local origin slot. The
// result must equal the single-threaded engine's.

#include <cstdint>
#include <optional>
#include <string>

#include "eval/function_registry.h"
#include "eval/scheduler.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

constexpr std::uint32_t kRows = 96U;

// Column A: values. B: an empty FILTER (measured origin). C: a reader of B
// (propagated). D: an empty TOCOL (unknown, plain). E: a number.
void Seed(Workbook& wb) {
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string r = std::to_string(row + 1U);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(row + 1.0))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=FILTER(A" + r + ",A" + r + ">1000)")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 2U, "=B" + r)));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 3U, "=TOCOL({1,2}/0,2)")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 4U, "=A" + r + "*2")));
  }
}

void ExpectDetails(const Workbook& wb) {
  const Sheet& sheet = wb.sheet(0U);
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    SCOPED_TRACE("row " + std::to_string(row));
    const std::optional<CellErrorDetail> origin = sheet.error_detail_at(row, 1U);
    ASSERT_TRUE(origin.has_value());
    EXPECT_EQ(origin->shape, CellErrorDetail::Shape::kSubType);
    EXPECT_EQ(origin->sub_type, 3U);
    const std::optional<CellErrorDetail> reader = sheet.error_detail_at(row, 2U);
    ASSERT_TRUE(reader.has_value());
    EXPECT_EQ(reader->shape, CellErrorDetail::Shape::kPropagated);
    EXPECT_FALSE(sheet.error_detail_at(row, 3U).has_value());
    EXPECT_FALSE(sheet.error_detail_at(row, 4U).has_value());
  }
}

TEST(ErrorDetailThread, ParallelRecalcMatchesTheSerialDetails) {
  Workbook serial = Workbook::create();
  Seed(serial);
  ASSERT_TRUE(static_cast<bool>(serial.ensure_calculated(default_registry())));
  ExpectDetails(serial);

  Workbook parallel = Workbook::create();
  Seed(parallel);
  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  ASSERT_TRUE(static_cast<bool>(parallel.recalc_parallel(default_registry(), cfg, nullptr)));
  ExpectDetails(parallel);

  // A second pass over the same cells rewrites every detail concurrently.
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(parallel.set_cell_value(0U, row, 0U, Value::number(row + 2.0))));
  }
  ASSERT_TRUE(static_cast<bool>(parallel.recalc_parallel(default_registry(), cfg, nullptr)));
  ExpectDetails(parallel);
}

}  // namespace
}  // namespace formulon::eval
