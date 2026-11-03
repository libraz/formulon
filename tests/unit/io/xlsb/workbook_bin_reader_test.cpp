//
// Unit tests for the `xl/workbook.bin` supporting-book decoder.

#include "io/xlsb/workbook_bin_reader.h"

#include <cstdint>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBeginExternals = static_cast<std::uint16_t>(XlsbRecordType::BrtBeginExternals);
constexpr std::uint16_t kSupSelf = static_cast<std::uint16_t>(XlsbRecordType::BrtSupSelf);
constexpr std::uint16_t kSupBookSrc = static_cast<std::uint16_t>(XlsbRecordType::BrtSupBookSrc);

TEST(XlsbWorkbookBinSupBooks, ListsSelfThenExternalBooks) {
  std::vector<std::uint8_t> body;
  emit_record(body, kBeginExternals, std::vector<std::uint8_t>{});
  emit_record(body, kSupSelf, std::vector<std::uint8_t>{});
  std::vector<std::uint8_t> payload;
  emit_xlwidestring(payload, "rId1");
  emit_record(body, kSupBookSrc, payload);

  auto books = DecodeSupBooks(body);
  ASSERT_TRUE(books.has_value());
  ASSERT_EQ(books.value().size(), 2U);
  EXPECT_EQ(books.value()[0].external_book, 0U);
  EXPECT_EQ(books.value()[1].external_book, 1U);
  EXPECT_EQ(books.value()[1].rel_id, "rId1");
}

TEST(XlsbWorkbookBinSupBooks, FramingErrorIsPropagated) {
  std::vector<std::uint8_t> body;
  emit_record(body, kBeginExternals, std::vector<std::uint8_t>{});
  emit_record(body, kSupSelf, std::vector<std::uint8_t>{});
  // A record header announcing 8 payload bytes with none following.
  body.push_back(static_cast<std::uint8_t>(kSupBookSrc & 0x7FU) | 0x80U);
  body.push_back(static_cast<std::uint8_t>(kSupBookSrc >> 7));
  body.push_back(8);

  auto books = DecodeSupBooks(body);
  ASSERT_FALSE(books.has_value());
  EXPECT_EQ(books.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
