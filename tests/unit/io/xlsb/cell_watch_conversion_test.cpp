//
// Cell watches across formats, against Mac Excel 365's `children_0_base` pair under
// `tests/fixtures/excel/retained_children_edit/`: the `.xlsx` `<cellWatches>` element and the `.xlsb`
// BrtBeginCellWatches block hold the same two watches (C5, E8).

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/writer.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb {
namespace {

namespace tc = test;

constexpr const char* kSheetBin = "xl/worksheets/sheet1.bin";
constexpr const char* kSheetXml = "xl/worksheets/sheet1.xml";
constexpr const char* kWatches = "<cellWatches><cellWatch r=\"C5\"/><cellWatch r=\"E8\"/></cellWatches>";

std::string Fixture(const char* ext) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/retained_children_edit/children_0_base." + ext;
}

/// The watch block of `package`'s sheet, with the type of the record before it.
struct WatchBlock {
  std::uint16_t before = 0;
  std::vector<tc::IndexedRecord> records;
};

WatchBlock WatchesOf(const tc::Bytes& package) {
  WatchBlock out;
  const std::optional<tc::Bytes> part = tc::ReadPart(package, kSheetBin);
  EXPECT_TRUE(part.has_value());
  if (!part) {
    return out;
  }
  const std::vector<tc::IndexedRecord> records = tc::SplitPart(*part);
  for (std::size_t i = 0; i < records.size(); ++i) {
    const std::uint16_t type = records[i].type;
    if (type != kBrtBeginCellWatches && type != kBrtCellWatch && type != kBrtEndCellWatches) {
      continue;
    }
    if (out.records.empty() && i > 0U) {
      out.before = records[i - 1U].type;
    }
    out.records.push_back(records[i]);
  }
  return out;
}

void ExpectSameWatches(const WatchBlock& ours, const WatchBlock& excel) {
  ASSERT_EQ(ours.records.size(), excel.records.size());
  ASSERT_EQ(excel.records.size(), 4U);
  EXPECT_EQ(ours.before, excel.before);
  for (std::size_t i = 0; i < excel.records.size(); ++i) {
    EXPECT_EQ(ours.records[i].type, excel.records[i].type) << i;
    EXPECT_EQ(ours.records[i].payload, excel.records[i].payload) << i;
  }
}

TEST(XlsbCellWatchConversion, XlsxWatchesWriteAsExcelsRecords) {
  const tc::Converted converted =
      tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), "excel/retained_children_edit/children_0_base.xlsx");
  ASSERT_TRUE(converted.error.empty()) << converted.error;
  ExpectSameWatches(WatchesOf(converted.bytes), WatchesOf(tc::ReadBytes(Fixture("xlsb"))));
}

TEST(XlsbCellWatchConversion, XlsbWatchesWriteBackUnchanged) {
  const tc::Converted converted =
      tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), "excel/retained_children_edit/children_0_base.xlsb");
  ASSERT_TRUE(converted.error.empty()) << converted.error;
  ExpectSameWatches(WatchesOf(converted.bytes), WatchesOf(tc::ReadBytes(Fixture("xlsb"))));
}

TEST(XlsbCellWatchConversion, XlsbWatchesSaveAsExcelsXlsxElement) {
  const tc::Bytes xlsb = tc::ReadBytes(Fixture("xlsb"));
  auto read = read_xlsb(ByteSpan{xlsb.data(), xlsb.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  auto written = write_ooxml(read.value().workbook);
  ASSERT_TRUE(written.has_value()) << written.error().message;
  const std::optional<tc::Bytes> ours = tc::ReadPart(written.value(), kSheetXml);
  const std::optional<tc::Bytes> excel = tc::ReadPart(tc::ReadBytes(Fixture("xlsx")), kSheetXml);
  ASSERT_TRUE(ours.has_value());
  ASSERT_TRUE(excel.has_value());
  const std::string ours_text(ours->begin(), ours->end());
  const std::string excel_text(excel->begin(), excel->end());
  ASSERT_NE(excel_text.find(kWatches), std::string::npos);
  const std::size_t at = ours_text.find(kWatches);
  ASSERT_NE(at, std::string::npos);
  // The element keeps its schema slot: after the page margins, before the ignored errors Excel writes next.
  EXPECT_LT(ours_text.find("<pageMargins"), at);
}

}  // namespace
}  // namespace formulon::io::xlsb
