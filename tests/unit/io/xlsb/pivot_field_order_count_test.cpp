// Axis field-order records reject a count that disagrees with the payload size.
// The 0x40000001 case wraps `count * 4` to 4 on wasm32 only; on a 64-bit host it
// is already rejected, so it documents the contract and cannot be red here.

#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/pivot_reader.h"
#include "io/xlsb/record.h"
#include "utils/error.h"

namespace formulon {
namespace {

// BrtPivotRowFields (309): two-byte type varint, one-byte size.
std::vector<std::uint8_t> RowFieldsRecord(std::uint32_t count, std::uint32_t trailing_u32s) {
  constexpr std::uint16_t kType = 309;
  std::vector<std::uint8_t> out = {static_cast<std::uint8_t>((kType & 0x7F) | 0x80),
                                   static_cast<std::uint8_t>(kType >> 7),
                                   static_cast<std::uint8_t>(4U + 4U * trailing_u32s)};
  for (int shift = 0; shift < 32; shift += 8) {
    out.push_back(static_cast<std::uint8_t>((count >> shift) & 0xFFU));
  }
  out.insert(out.end(), 4U * trailing_u32s, 0U);
  return out;
}

TEST(XlsbPivotFieldOrder, CountWrappingTheByteSizeIsCorrupt) {
  const auto bytes = RowFieldsRecord(0x40000001U, 1U);
  auto table = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table));
  EXPECT_EQ(table.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbPivotFieldOrder, CountLargerThanPayloadIsCorrupt) {
  const auto bytes = RowFieldsRecord(3U, 1U);
  auto table = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table));
  EXPECT_EQ(table.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

}  // namespace
}  // namespace formulon
