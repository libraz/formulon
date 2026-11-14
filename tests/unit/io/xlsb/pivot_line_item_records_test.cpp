#include "io/xlsb/pivot_line_item_records.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "pivot_twin_corpus.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using pivot_twins::CorpusPaths;
using pivot_twins::Parts;

constexpr const char* kTablePrefix = "xl/pivotTables/pivotTable";

/// `t` spellings in ST_ItemType order, the numbering the records use.
constexpr const char* kItemTypes[] = {"data",  "default", "sum",     "countA", "avg",  "max",   "min",  "product",
                                      "count", "stdDev",  "stdDevP", "var",    "varP", "grand", "blank"};

std::uint8_t ItemTypeOf(const char* name) {
  for (std::size_t k = 0; k < sizeof(kItemTypes) / sizeof(kItemTypes[0]); ++k) {
    if (std::string(name) == kItemTypes[k]) {
      return static_cast<std::uint8_t>(k);
    }
  }
  ADD_FAILURE() << "unknown item type " << name;
  return 0xFF;
}

/// A decoded block and the exact bytes it was read from.
struct Block {
  SxLineItems items;
  std::vector<std::uint8_t> bytes;
};

Block DecodeBlock(const std::vector<std::uint8_t>& part, SxLineItemAxis axis) {
  const std::vector<FramedRecord> framed = split_records(part);
  std::vector<XlsbRecord> records;
  for (const FramedRecord& rec : framed) {
    records.push_back(XlsbRecord{rec.type, rec.payload});
  }
  const std::uint16_t begin =
      axis == SxLineItemAxis::kRows ? pivot_records::kBeginSXLIRws : pivot_records::kBeginSXLICols;
  Block block;
  std::size_t pos = 0;
  while (pos < records.size() && records[pos].type != begin) {
    ++pos;
  }
  if (pos == records.size()) {
    ADD_FAILURE() << "no line-item block";
    return block;
  }
  const std::size_t first = pos;
  auto items = decode_line_items(records, &pos, axis);
  if (!items) {
    ADD_FAILURE() << items.error().message;
    return block;
  }
  block.items = std::move(items.value());
  block.bytes.assign(part.begin() + static_cast<std::ptrdiff_t>(framed[first].begin),
                     part.begin() + static_cast<std::ptrdiff_t>(framed[pos - 1U].end));
  return block;
}

/// `<rowItems>` / `<colItems>` of the twin `.xlsx` part, as typed rows.
SxLineItems XmlItems(const std::string& xml, const char* element) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(xml.data(), xml.size()));
  SxLineItems out;
  for (const pugi::xml_node i : doc.child("pivotTableDefinition").child(element).children("i")) {
    SxLineItem item;
    item.r = static_cast<std::uint16_t>(i.attribute("r").as_uint(0));
    item.t = ItemTypeOf(i.attribute("t").as_string("data"));
    item.i = i.attribute("i").as_uint(0);
    for (const pugi::xml_node x : i.children("x")) {
      item.x.push_back(x.attribute("v").as_uint(0));
    }
    out.items.push_back(std::move(item));
  }
  return out;
}

std::string TwinXml(const std::string& xlsb_path, const std::string& bin_part) {
  const std::vector<std::uint8_t> package = ReadFixtureFile(xlsb_path.substr(0, xlsb_path.size() - 1U) + "x");
  ZipReader zip;
  EXPECT_TRUE(zip.open(SpanOf(package)));
  auto bytes = zip.read_entry(bin_part.substr(0, bin_part.size() - 4U) + ".xml");
  EXPECT_TRUE(bytes.has_value());
  return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
}

void ExpectSame(const SxLineItems& actual, const SxLineItems& expected) {
  ASSERT_EQ(actual.items.size(), expected.items.size());
  for (std::size_t k = 0; k < actual.items.size(); ++k) {
    SCOPED_TRACE("item " + std::to_string(k));
    EXPECT_EQ(actual.items[k].r, expected.items[k].r);
    EXPECT_EQ(actual.items[k].t, expected.items[k].t);
    EXPECT_EQ(actual.items[k].i, expected.items[k].i);
    EXPECT_EQ(actual.items[k].x, expected.items[k].x);
  }
}

std::vector<std::uint8_t> TablePart(const std::string& stem, std::string* path_out) {
  for (const std::string& path : CorpusPaths()) {
    const std::string file = path.substr(path.find_last_of('/') + 1U);
    if (file.compare(0, stem.size(), stem) == 0) {
      for (auto& part : Parts(path, kTablePrefix)) {
        if (part.first == "xl/pivotTables/pivotTable1.bin") {
          *path_out = path;
          return std::move(part.second);
        }
      }
    }
  }
  ADD_FAILURE() << "no pivot table part for " << stem;
  return {};
}

TEST(XlsbPivotLineItemRecords, EveryBlockRoundTripsAndMatchesTheTwin) {
  std::size_t checked = 0;
  for (const std::string& path : CorpusPaths()) {
    for (const auto& part : Parts(path, kTablePrefix)) {
      SCOPED_TRACE(path + " " + part.first);
      const std::string xml = TwinXml(path, part.first);
      for (const SxLineItemAxis axis : {SxLineItemAxis::kRows, SxLineItemAxis::kCols}) {
        const Block block = DecodeBlock(part.second, axis);
        std::vector<std::uint8_t> encoded;
        encode_line_items(encoded, block.items, axis);
        EXPECT_EQ(encoded, block.bytes);
        ExpectSame(block.items, XmlItems(xml, axis == SxLineItemAxis::kRows ? "rowItems" : "colItems"));
        ++checked;
      }
    }
  }
  EXPECT_EQ(checked, 178U);
}

TEST(XlsbPivotLineItemRecords, DataFieldIndexAndKindsOnRows) {
  std::string path;
  const std::vector<std::uint8_t> part = TablePart("A3_ja_d8_two_rows_values_rows", &path);
  const SxLineItems rows = DecodeBlock(part, SxLineItemAxis::kRows).items;
  ASSERT_EQ(rows.items.size(), 26U);
  // <i r="2" i="1"><x v="1"/></i>
  EXPECT_EQ(rows.items[3].r, 2U);
  EXPECT_EQ(rows.items[3].i, 1U);
  EXPECT_EQ(rows.items[3].x, std::vector<std::uint32_t>({1U}));
  // <i t="default" i="1"><x/></i>
  EXPECT_EQ(rows.items[5].t, 1U);
  EXPECT_EQ(rows.items[5].i, 1U);
  // <i t="grand" i="1"><x/></i>
  EXPECT_EQ(rows.items[25].t, kSxLineItemTypeGrand);
  EXPECT_EQ(rows.items[25].i, 1U);
  const SxLineItems cols = DecodeBlock(part, SxLineItemAxis::kCols).items;
  ASSERT_EQ(cols.items.size(), 1U);
  EXPECT_TRUE(cols.items[0].x.empty());
}

TEST(XlsbPivotLineItemRecords, SeveralMembersOnColumns) {
  std::string path;
  const std::vector<std::uint8_t> part = TablePart("A4_ja_c3_cols_values_cols", &path);
  const SxLineItems cols = DecodeBlock(part, SxLineItemAxis::kCols).items;
  ASSERT_EQ(cols.items.size(), 6U);
  // <i><x/><x/></i>
  EXPECT_EQ(cols.items[0].x, std::vector<std::uint32_t>({0U, 0U}));
  // <i r="1" i="1"><x v="1"/></i>
  EXPECT_EQ(cols.items[1].r, 1U);
  EXPECT_EQ(cols.items[1].i, 1U);
  EXPECT_EQ(cols.items[1].x, std::vector<std::uint32_t>({1U}));
  // <i><x v="1"/><x/></i>
  EXPECT_EQ(cols.items[2].x, std::vector<std::uint32_t>({1U, 0U}));
  EXPECT_EQ(cols.items[5].t, kSxLineItemTypeGrand);
  EXPECT_EQ(cols.items[5].i, 1U);
}

TEST(XlsbPivotLineItemRecords, ValuesOnlyTable) {
  std::string path;
  const std::vector<std::uint8_t> part = TablePart("A9_ja_d13_values_only", &path);
  // <rowItems count="1"><i/></rowItems>
  const SxLineItems rows = DecodeBlock(part, SxLineItemAxis::kRows).items;
  ASSERT_EQ(rows.items.size(), 1U);
  EXPECT_EQ(rows.items[0].t, 0U);
  EXPECT_TRUE(rows.items[0].x.empty());
  // <colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>
  const SxLineItems cols = DecodeBlock(part, SxLineItemAxis::kCols).items;
  ASSERT_EQ(cols.items.size(), 2U);
  EXPECT_EQ(cols.items[1].i, 1U);
  EXPECT_EQ(cols.items[1].x, std::vector<std::uint32_t>({1U}));
}

TEST(XlsbPivotLineItemRecords, SubtotalAndBlankKinds) {
  std::string path;
  // <i t="sum" r="1">, <i t="var" r="1">, <i t="avg">, <i t="product">, <i t="count">
  const SxLineItems sums = DecodeBlock(TablePart("A5_C9__sum", &path), SxLineItemAxis::kRows).items;
  ASSERT_EQ(sums.items.size(), 29U);
  EXPECT_EQ(sums.items[4].t, 2U);
  EXPECT_EQ(sums.items[4].r, 1U);
  EXPECT_EQ(sums.items[5].t, 11U);
  EXPECT_EQ(sums.items[11].t, 4U);
  EXPECT_EQ(sums.items[12].t, 7U);
  EXPECT_EQ(sums.items[13].t, 8U);
  EXPECT_EQ(sums.items[13].x, std::vector<std::uint32_t>({3U}));
  // <i t="blank"><x/></i>
  const SxLineItems blank = DecodeBlock(TablePart("C02_insert_blank_row", &path), SxLineItemAxis::kRows).items;
  ASSERT_EQ(blank.items.size(), 13U);
  EXPECT_EQ(blank.items[3].t, kSxLineItemTypeMax);
}

std::vector<XlsbRecord> Frame(const std::vector<std::uint8_t>& bytes) {
  std::vector<XlsbRecord> records;
  for (const FramedRecord& rec : split_records(bytes)) {
    records.push_back(XlsbRecord{rec.type, rec.payload});
  }
  return records;
}

TEST(XlsbPivotLineItemRecords, RefusesUncharacterisedShapes) {
  SxLineItems items;
  items.items.push_back(SxLineItem{0, 0, 0, {2U}});
  std::vector<std::uint8_t> bytes;
  encode_line_items(bytes, items, SxLineItemAxis::kRows);

  // Reserved byte 3 of BrtBeginSXLI: the count record takes 7 bytes, the SXLI header 3.
  std::vector<std::uint8_t> flagged = bytes;
  flagged[7U + 3U + 3U] = 0x01;
  std::vector<XlsbRecord> records = Frame(flagged);
  std::size_t pos = 0;
  auto refused = decode_line_items(records, &pos, SxLineItemAxis::kRows);
  ASSERT_FALSE(refused.has_value());
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);

  // Declared count 2 against one row.
  std::vector<std::uint8_t> miscounted = bytes;
  miscounted[3] = 0x02;
  records = Frame(miscounted);
  pos = 0;
  refused = decode_line_items(records, &pos, SxLineItemAxis::kRows);
  ASSERT_FALSE(refused.has_value());
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);

  // A column block read as rows.
  records = Frame(bytes);
  pos = 0;
  refused = decode_line_items(records, &pos, SxLineItemAxis::kCols);
  ASSERT_FALSE(refused.has_value());
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoXlsbUnknownRecord);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
