// Axis field-order records reject a count that disagrees with the payload size.
// The 0x40000001 case wraps `count * 4` to 4 on wasm32 only; on a 64-bit host it
// is already rejected, so it documents the contract and cannot be red here.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/pivot_table_bridge.h"
#include "io/xlsb/pivot_table_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "pivot_twin_corpus.h"
#include "utils/error.h"

namespace formulon {
namespace {

/// B01_base's pivot table part.
std::vector<std::uint8_t> BasePart() {
  for (const std::string& path : io::xlsb::pivot_twins::CorpusPaths()) {
    if (path.find("/B01_base.xlsb") != std::string::npos) {
      for (auto& part : io::xlsb::pivot_twins::Parts(path, "xl/pivotTables/pivotTable")) {
        return std::move(part.second);
      }
    }
  }
  ADD_FAILURE() << "no B01_base pivot table part";
  return {};
}

/// The part with its BrtBeginISXVDRws payload replaced by `count` and `trailing_u32s` zero entries.
std::vector<std::uint8_t> WithRowFields(std::uint32_t count, std::uint32_t trailing_u32s) {
  const std::vector<std::uint8_t> part = BasePart();
  io::ByteSpan cursor{part.data(), part.size()};
  std::vector<std::uint8_t> out;
  while (cursor.size > 0U) {
    auto rec = io::xlsb::read_record(cursor);
    EXPECT_TRUE(static_cast<bool>(rec));
    if (!rec) {
      return {};
    }
    if (rec.value().type != io::xlsb::pivot_records::kBeginISXVDRws) {
      io::xlsb::emit_record(out, rec.value().type, rec.value().payload);
      continue;
    }
    std::vector<std::uint8_t> payload;
    io::xlsb::emit_u32(payload, count);
    payload.insert(payload.end(), 4U * trailing_u32s, 0U);
    io::xlsb::emit_record(out, rec.value().type, payload);
  }
  return out;
}

TEST(XlsbPivotFieldOrder, CountWrappingTheByteSizeIsCorrupt) {
  const auto bytes = WithRowFields(0x40000001U, 1U);
  auto table = io::xlsb::pivot_table_from_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table));
  EXPECT_EQ(table.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbPivotFieldOrder, CountLargerThanPayloadIsCorrupt) {
  const auto bytes = WithRowFields(3U, 1U);
  auto table = io::xlsb::pivot_table_from_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table));
  EXPECT_EQ(table.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbPivotTableDef, UseAutoFormattingIsBitZeroOfByteFive) {
  const std::vector<std::uint8_t> part = BasePart();
  auto doc = io::xlsb::decode_pivot_table(io::ByteSpan{part.data(), part.size()});
  ASSERT_TRUE(static_cast<bool>(doc)) << doc.error().message;
  for (const bool on : {true, false}) {
    io::xlsb::PivotTableDoc changed = doc.value();
    changed.view.flags5 = static_cast<std::uint8_t>(on ? changed.view.flags5 | 0x01U : changed.view.flags5 & ~0x01U);
    auto bytes = io::xlsb::encode_pivot_table(changed);
    ASSERT_TRUE(static_cast<bool>(bytes));
    auto table = io::xlsb::pivot_table_from_xlsb(io::ByteSpan{bytes.value().data(), bytes.value().size()});
    ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
    EXPECT_EQ(table.value().auto_formatting(), on);
  }
}

}  // namespace
}  // namespace formulon
