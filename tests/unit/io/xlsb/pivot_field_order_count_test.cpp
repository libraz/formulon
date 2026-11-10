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

// A minimal pivot table part: BrtBeginPivotTableDef with the given flags byte at payload offset 5, one field and
// one data field.
std::vector<std::uint8_t> TableWithFlagsByte(std::uint8_t flags) {
  std::vector<std::uint8_t> out;
  const auto record = [&out](std::uint16_t type, const std::vector<std::uint8_t>& payload) {
    out.push_back(static_cast<std::uint8_t>((type & 0x7F) | 0x80));
    out.push_back(static_cast<std::uint8_t>(type >> 7));
    out.push_back(static_cast<std::uint8_t>(payload.size()));
    out.insert(out.end(), payload.begin(), payload.end());
  };
  std::vector<std::uint8_t> head(40U, 0U);
  head[5] = flags;
  record(280, head);
  record(287, {1, 0, 0, 0});
  record(285, {});
  record(293, std::vector<std::uint8_t>(29U, 0U));
  record(315, {});
  return out;
}

TEST(XlsbPivotTableDef, UseAutoFormattingIsBitZeroOfByteFive) {
  const auto on = TableWithFlagsByte(0x65U);
  auto on_table = io::xlsb::read_pivot_table_bin(io::ByteSpan{on.data(), on.size()});
  ASSERT_TRUE(static_cast<bool>(on_table)) << on_table.error().message;
  EXPECT_TRUE(on_table.value().auto_formatting());

  const auto off = TableWithFlagsByte(0x64U);
  auto off_table = io::xlsb::read_pivot_table_bin(io::ByteSpan{off.data(), off.size()});
  ASSERT_TRUE(static_cast<bool>(off_table)) << off_table.error().message;
  EXPECT_FALSE(off_table.value().auto_formatting());
}

}  // namespace
}  // namespace formulon
