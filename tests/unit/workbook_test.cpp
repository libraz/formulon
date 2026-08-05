//
// Unit tests for the Workbook skeleton. Verifies the factory shape, sheet
// accessor mutation, and the basic byte-level shape of the save() output.

#include "workbook.h"

#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/passthrough_part.h"
#include "sheet.h"
#include "value.h"

namespace formulon {
namespace {

TEST(WorkbookTest, CreateYieldsSingleSheetNamedSheet1) {
  Workbook wb = Workbook::create();
  ASSERT_EQ(wb.sheet_count(), 1u);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
}

TEST(WorkbookTest, SheetCountReturnsOne) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(wb.sheet_count(), static_cast<std::size_t>(1));
}

TEST(WorkbookTest, MutateSheetNamePropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_name("Daten");
  EXPECT_EQ(wb.sheet(0).name(), "Daten");

  // The const overload must observe the same state.
  const Workbook& const_wb = wb;
  EXPECT_EQ(const_wb.sheet(0).name(), "Daten");
}

TEST(WorkbookTest, SaveProducesNonEmptyBytes) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result)) << "save() failed: " << result.error().message;
  const std::vector<std::uint8_t>& bytes = result.value();
  EXPECT_GT(bytes.size(), 0u);
}

TEST(WorkbookTest, SaveIsZipMagicBytes) {
  Workbook wb = Workbook::create();
  auto result = wb.save();
  ASSERT_TRUE(static_cast<bool>(result));
  const std::vector<std::uint8_t>& bytes = result.value();
  ASSERT_GE(bytes.size(), 4u);
  EXPECT_EQ(bytes[0], 0x50u);  // 'P'
  EXPECT_EQ(bytes[1], 0x4Bu);  // 'K'
  EXPECT_EQ(bytes[2], 0x03u);
  EXPECT_EQ(bytes[3], 0x04u);
}

TEST(WorkbookTest, ApproximateMemoryGrowsWithTheCellStore) {
  // The figure exists so a host runtime can size a workbook it only sees
  // as a handle, which means the one property that must hold is that it
  // moves with the content rather than staying at the empty baseline.
  Workbook wb = Workbook::create();
  const std::size_t empty = wb.approximate_memory_bytes();
  EXPECT_GT(empty, sizeof(Workbook));

  for (std::uint32_t row = 0; row < 500U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 0U, "=1+2")));
  }
  const std::size_t filled = wb.approximate_memory_bytes();
  EXPECT_GT(filled, empty);
  // 500 rows of formula text and cell slots is several tens of KB; a
  // figure that ignored the cell store would land within a few hundred
  // bytes of the empty baseline instead.
  EXPECT_GT(filled - empty, 500U * sizeof(Cell));
}

TEST(WorkbookTest, ApproximateMemoryCountsPassthroughPayloads) {
  // Passthrough carries the package's unmodelled binaries — embedded
  // images above all — so a workbook that opens a media-heavy file must
  // not look small to the host.
  Workbook wb = Workbook::create();
  const std::size_t before = wb.approximate_memory_bytes();

  constexpr std::size_t kPayloadBytes = 256U * 1024U;
  std::vector<io::PassthroughPart> parts;
  parts.push_back(
      io::PassthroughPart{"xl/media/image1.png", "image/png", std::vector<std::uint8_t>(kPayloadBytes, 0x7FU)});
  wb.set_passthrough_parts(std::move(parts));

  EXPECT_GE(wb.approximate_memory_bytes() - before, kPayloadBytes);
}

TEST(WorkbookTest, ApproximateMemoryIsStableWithoutMutation) {
  // Hosts report the delta against the previous reading, so a repeated
  // call on an unchanged workbook has to return the same number or the
  // accounting drifts on every poll.
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  const std::size_t first = wb.approximate_memory_bytes();
  EXPECT_EQ(wb.approximate_memory_bytes(), first);
}

}  // namespace
}  // namespace formulon
