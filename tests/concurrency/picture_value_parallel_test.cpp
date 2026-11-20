//
// ThreadSanitizer-targeted test for picture comparison on recalc workers.
//
// Workers read the workbook's picture table through the scope each formula
// evaluation opens, while nothing mutates it. The fixture compares, counts,
// sorts and deduplicates pictures in several dependent layers of many
// independent rows, and asserts the parallel result equals the serial one.

#include <cstdint>
#include <optional>
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

// Rows per layer: enough pooled tasks per worker, small enough for the TSan event budget.
constexpr std::uint32_t kRows = 96U;
// Columns B..F, each layer reading the one before it.
constexpr std::uint32_t kFirstCol = 1U;
constexpr std::uint32_t kLastCol = 5U;

using Bytes = std::vector<std::uint8_t>;

void PutBe32(Bytes& out, std::uint32_t v) {
  for (int shift = 24; shift >= 0; shift -= 8) {
    out.push_back(static_cast<std::uint8_t>(v >> shift));
  }
}

// A PNG whose header probes as 4 x 4; `tag` makes otherwise equal images differ.
Bytes Png(std::uint8_t tag) {
  Bytes out = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  PutBe32(out, 4U);
  PutBe32(out, 4U);
  out.insert(out.end(), {8, 6, 0, 0, 0, tag, tag, tag, tag});
  return out;
}

Value StoredValue(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  if (const Cell* cell = wb.sheet(0U).cell_at(row, col); cell != nullptr) {
    return cell->cached_value;
  }
  return Value::blank();
}

// A1 and A3: one picture; A2: greater bytes with an alt. Per row r:
//   B  =$A$1=$A$2                                   FALSE
//   C  =IF(Br,0,COUNTIF($A$1:$A$3,$A$1))            2
//   D  =INDEX(SORT($A$1:$A$2),1+0*Cr)               A2's picture
//   E  =Dr=$A$2                                     TRUE
//   F  =ROWS(UNIQUE($A$1:$A$3))+0*Er                2
void SeedPictureWorkbook(Workbook& wb) {
  const Bytes low = Png(0x11);
  const Bytes high = Png(0x22);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_picture(0U, 0U, 0U, low.data(), low.size(), std::nullopt)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_picture(0U, 1U, 0U, high.data(), high.size(), "b")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_picture(0U, 2U, 0U, low.data(), low.size(), std::nullopt)));
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const std::string r = std::to_string(row + 1U);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=$A$1=$A$2")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 2U, "=IF(B" + r + ",0,COUNTIF($A$1:$A$3,$A$1))")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 3U, "=INDEX(SORT($A$1:$A$2),1+0*C" + r + ")")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 4U, "=D" + r + "=$A$2")));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 5U, "=ROWS(UNIQUE($A$1:$A$3))+0*E" + r)));
  }
}

void ExpectFixtureResolved(const Workbook& wb) {
  const Value high = StoredValue(wb, 1U, 0U);
  ASSERT_TRUE(high.is_picture());
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const Value b = StoredValue(wb, row, 1U);
    ASSERT_TRUE(b.is_boolean()) << "B row " << row;
    EXPECT_FALSE(b.as_boolean()) << "B row " << row;
    const Value c = StoredValue(wb, row, 2U);
    ASSERT_TRUE(c.is_number()) << "C row " << row;
    EXPECT_EQ(2.0, c.as_number()) << "C row " << row;
    const Value d = StoredValue(wb, row, 3U);
    ASSERT_TRUE(d.is_picture()) << "D row " << row;
    EXPECT_EQ(high.as_picture().picture_id, d.as_picture().picture_id) << "D row " << row;
    const Value e = StoredValue(wb, row, 4U);
    ASSERT_TRUE(e.is_boolean()) << "E row " << row;
    EXPECT_TRUE(e.as_boolean()) << "E row " << row;
    const Value f = StoredValue(wb, row, 5U);
    ASSERT_TRUE(f.is_number()) << "F row " << row;
    EXPECT_EQ(2.0, f.as_number()) << "F row " << row;
  }
}

TEST(PictureValueParallel, CompareAndSortOnWorkers) {
  Workbook serial = Workbook::create();
  Workbook parallel = Workbook::create();
  ASSERT_NO_FATAL_FAILURE(SeedPictureWorkbook(serial));
  ASSERT_NO_FATAL_FAILURE(SeedPictureWorkbook(parallel));

  SchedulerConfig cfg;
  cfg.num_threads = 4U;
  SchedulerStats stats;
  ASSERT_TRUE(static_cast<bool>(serial.recalc(default_registry())));
  ASSERT_TRUE(static_cast<bool>(parallel.recalc_parallel(default_registry(), cfg, &stats)));
  EXPECT_GE(stats.parallel_steps, 1U);
  EXPECT_GE(stats.worker_threads_started, 2U);

  ASSERT_NO_FATAL_FAILURE(ExpectFixtureResolved(serial));
  ASSERT_NO_FATAL_FAILURE(ExpectFixtureResolved(parallel));
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    for (std::uint32_t col = kFirstCol; col <= kLastCol; ++col) {
      EXPECT_EQ(StoredValue(serial, row, col), StoredValue(parallel, row, col))
          << "value mismatch at (" << row << ", " << col << ")";
    }
  }
}

}  // namespace
}  // namespace formulon::eval
