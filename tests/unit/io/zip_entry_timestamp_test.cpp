//
// Saved packages carry a fixed entry timestamp, so saving an unchanged model
// twice yields identical bytes however far apart the saves are.

#include <cstddef>
#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "value.h"
#include "workbook.h"
#include "workbook_format.h"

namespace formulon {
namespace {

// 1980-01-01 00:00 in MS-DOS date / time encoding.
constexpr std::uint16_t kDosDate = 0x0021U;
constexpr std::uint16_t kDosTime = 0x0000U;

std::uint16_t U16(const std::vector<std::uint8_t>& bytes, std::size_t at) {
  return static_cast<std::uint16_t>(bytes[at] | (bytes[at + 1U] << 8U));
}

std::uint32_t U32(const std::vector<std::uint8_t>& bytes, std::size_t at) {
  return static_cast<std::uint32_t>(U16(bytes, at)) | (static_cast<std::uint32_t>(U16(bytes, at + 2U)) << 16U);
}

// Checks the date and time of every central-directory entry and of the local
// header it points to; returns the number of entries seen.
std::size_t ExpectFixedTimestamps(const std::vector<std::uint8_t>& zip) {
  constexpr std::size_t kEocdSize = 22U;
  EXPECT_GE(zip.size(), kEocdSize);
  const std::size_t eocd = zip.size() - kEocdSize;
  EXPECT_EQ(U32(zip, eocd), 0x06054b50U);
  const std::size_t entries = U16(zip, eocd + 10U);
  std::size_t at = U32(zip, eocd + 16U);
  for (std::size_t i = 0; i < entries; ++i) {
    EXPECT_EQ(U32(zip, at), 0x02014b50U);
    EXPECT_EQ(U16(zip, at + 12U), kDosTime);
    EXPECT_EQ(U16(zip, at + 14U), kDosDate);
    const std::size_t local = U32(zip, at + 42U);
    EXPECT_EQ(U32(zip, local), 0x04034b50U);
    EXPECT_EQ(U16(zip, local + 10U), kDosTime);
    EXPECT_EQ(U16(zip, local + 12U), kDosDate);
    at += 46U + U16(zip, at + 28U) + U16(zip, at + 30U) + U16(zip, at + 32U);
  }
  return entries;
}

class ZipEntryTimestampTest : public ::testing::TestWithParam<WorkbookFormat> {};

TEST_P(ZipEntryTimestampTest, EveryEntryCarriesTheDosEpoch) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(1.5)));
  auto saved = wb.save_as(GetParam());
  ASSERT_TRUE(saved) << saved.error().message;
  EXPECT_GT(ExpectFixedTimestamps(saved.value()), 0U);
}

TEST_P(ZipEntryTimestampTest, UnchangedModelSavesIdenticalBytes) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(1.5)));
  auto first = wb.save_as(GetParam());
  auto second = wb.save_as(GetParam());
  ASSERT_TRUE(first && second);
  EXPECT_EQ(first.value(), second.value());

  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(2.5)));
  auto changed = wb.save_as(GetParam());
  ASSERT_TRUE(changed);
  EXPECT_NE(first.value(), changed.value());
}

INSTANTIATE_TEST_SUITE_P(Formats, ZipEntryTimestampTest,
                         ::testing::Values(WorkbookFormat::Ooxml, WorkbookFormat::Xlsb));

}  // namespace
}  // namespace formulon
