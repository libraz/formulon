// A cached external-cell error is stored through `stored_cell_error`, like every
// other cell-value writer.

#include <cstdint>
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

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
