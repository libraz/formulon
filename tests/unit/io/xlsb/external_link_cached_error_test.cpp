// A cached external-cell error is stored through `stored_cell_error`, like every
// other cell-value writer.

#include <algorithm>
#include <cstdint>
#include <limits>
#include <vector>

#include "external_link.h"
#include "gtest/gtest.h"
#include "io/stored_cell_error.h"
#include "io/xlsb/external_link_reader.h"
#include "io/xlsb/external_link_writer.h"
#include "io/xlsb/record.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

ErrorCode CachedErrorAfterWrite(ErrorCode written) {
  ExternalLinkRecord link;
  link.index = 1;
  link.kind = ExternalLinkRecord::Kind::kExternalBook;
  link.book.sheet_names = {"S"};
  link.book.sheet_data = {true};  // the writer emits only sheets flagged as cached
  link.book.cells[ExternalBook::cell_key(0, 0, 0)].value = Value::error(written);
  XlsbLinkTables tables;
  tables.index = 1;
  tables.sheet_names = {"S"};
  const std::vector<std::uint8_t> bin = build_external_link_bin(link, tables);
  auto book = read_external_link_bin(ByteSpan{bin.data(), bin.size()});
  EXPECT_TRUE(static_cast<bool>(book));
  if (!book) {
    return ErrorCode::Unknown;
  }
  const Value v = book.value().cached_cell(0, 0, 0);
  EXPECT_TRUE(v.is_error());
  return v.is_error() ? v.as_error() : ErrorCode::Unknown;
}

TEST(XlsbExternalLinkCachedError, SpillIsStoredAsStoredCellErrorMapsIt) {
  EXPECT_EQ(CachedErrorAfterWrite(ErrorCode::Spill), stored_cell_error(ErrorCode::Spill));
}

TEST(XlsbExternalLinkCachedError, CalcIsStoredAsStoredCellErrorMapsIt) {
  EXPECT_EQ(CachedErrorAfterWrite(ErrorCode::Calc), stored_cell_error(ErrorCode::Calc));
}

TEST(XlsbExternalLinkCachedError, LegacyErrorIsKept) {
  EXPECT_EQ(CachedErrorAfterWrite(ErrorCode::Div0), ErrorCode::Div0);
}

TEST(XlsbExternalLinkCachedError, NonFiniteNumbersAreStoredAsNumErrors) {
  ExternalLinkRecord link;
  link.index = 1;
  link.kind = ExternalLinkRecord::Kind::kExternalBook;
  link.book.sheet_names = {"S"};
  link.book.sheet_data = {true};
  const double inputs[] = {std::numeric_limits<double>::infinity(), -std::numeric_limits<double>::infinity(),
                           std::numeric_limits<double>::quiet_NaN(), 12.5};
  for (std::uint32_t col = 0; col < 4U; ++col) {
    link.book.cells[ExternalBook::cell_key(0, 0, col)].value = Value::number(inputs[col]);
  }
  XlsbLinkTables tables;
  tables.index = 1;
  tables.sheet_names = {"S"};
  const std::vector<std::uint8_t> bin = build_external_link_bin(link, tables);
  auto book = read_external_link_bin(ByteSpan{bin.data(), bin.size()});
  ASSERT_TRUE(static_cast<bool>(book));
  for (std::uint32_t col = 0; col < 3U; ++col) {
    const Value v = book.value().cached_cell(0, 0, col);
    ASSERT_TRUE(v.is_error()) << col;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << col;
  }
  const Value finite = book.value().cached_cell(0, 0, 3);
  ASSERT_TRUE(finite.is_number());
  EXPECT_EQ(finite.as_number(), 12.5);
}

TEST(XlsbExternalLinkCachedError, OffGridColumnDoesNotAliasTheNextRow) {
  ExternalLinkRecord link;
  link.index = 1;
  link.kind = ExternalLinkRecord::Kind::kExternalBook;
  link.book.sheet_names = {"S"};
  link.book.sheet_data = {true};
  link.book.cells[ExternalBook::cell_key(0, 0, 5)].value = Value::number(7.0);
  link.book.cells[ExternalBook::cell_key(0, 1, 0)].value = Value::number(2.0);
  XlsbLinkTables tables;
  tables.index = 1;
  tables.sheet_names = {"S"};
  std::vector<std::uint8_t> bin = build_external_link_bin(link, tables);
  // Rewrite the column of the (0, 5) cell to 16384, one past the grid.
  std::vector<std::uint8_t> needle = {5, 0, 0, 0};
  const double seven = 7.0;
  const auto* seven_bytes = reinterpret_cast<const std::uint8_t*>(&seven);
  needle.insert(needle.end(), seven_bytes, seven_bytes + sizeof(seven));
  const auto at = std::search(bin.begin(), bin.end(), needle.begin(), needle.end());
  ASSERT_NE(at, bin.end());
  at[0] = 0x00;
  at[1] = 0x40;
  auto book = read_external_link_bin(ByteSpan{bin.data(), bin.size()});
  ASSERT_TRUE(static_cast<bool>(book));
  const Value next_row = book.value().cached_cell(0, 1, 0);
  ASSERT_TRUE(next_row.is_number());
  EXPECT_EQ(next_row.as_number(), 2.0);
  EXPECT_EQ(book.value().cells.size(), 1U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
