// An Excel-saved .xlsb whose extension conditional format or validation referenced a sheet deleted before the
// save keeps that reference as a 3-D Ptg through a BrtExternSheet entry of this workbook with itabFirst =
// itabLast = -1 (the formula reads `#REF!`). Re-saving the book unedited keeps the entry in place and the record
// bytes as Excel wrote them.

#include <cstddef>
#include <cstdint>
#include <string>
#include <tuple>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Xti = std::tuple<std::uint32_t, std::int32_t, std::int32_t>;

std::vector<std::uint8_t> Fixture(const char* name) {
  return ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name);
}

std::vector<std::uint8_t> Part(const std::vector<std::uint8_t>& package, const std::string& name) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry)) << name;
  return entry ? entry.value() : std::vector<std::uint8_t>{};
}

std::vector<Xti> XtiTable(const std::vector<std::uint8_t>& package) {
  std::vector<Xti> out;
  const std::vector<std::uint8_t> workbook = Part(package, "xl/workbook.bin");
  for (const FramedRecord& r : split_records(workbook)) {
    if (r.type != static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      continue;
    }
    const std::uint32_t count = load_u32(r.payload.data);
    for (std::uint32_t i = 0; i < count; ++i) {
      const std::uint8_t* at = r.payload.data + 4U + 12U * i;
      out.emplace_back(load_u32(at), static_cast<std::int32_t>(load_u32(at + 4)),
                       static_cast<std::int32_t>(load_u32(at + 8)));
    }
  }
  return out;
}

// The future-record payloads (conditional format and validation extension records) of one sheet part, as
// "<type>:<hex>".
std::vector<std::string> FutureRecords(const std::vector<std::uint8_t>& sheet) {
  static constexpr char kHex[] = "0123456789abcdef";
  std::vector<std::string> out;
  for (const FramedRecord& r : split_records(sheet)) {
    if (r.type < 1046U || r.type > 1054U) {
      continue;
    }
    std::string text = std::to_string(r.type) + ":";
    for (std::size_t i = 0; i < r.payload.size; ++i) {
      text += kHex[r.payload.data[i] >> 4U];
      text += kHex[r.payload.data[i] & 0x0FU];
    }
    out.push_back(std::move(text));
  }
  return out;
}

std::string SheetXml(const Workbook& wb, int part) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved));
  if (!saved) {
    return {};
  }
  const std::vector<std::uint8_t> xml = Part(saved.value(), "xl/worksheets/sheet" + std::to_string(part) + ".xml");
  return std::string(xml.begin(), xml.end());
}

void ExpectUneditedResaveKeepsTheDeadSlot(const char* fixture, const std::string& sheet_part, const char* formula) {
  const std::vector<std::uint8_t> source = Fixture(fixture);
  auto loaded = read_xlsb(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  EXPECT_NE(SheetXml(wb, 1).find(formula), std::string::npos);

  auto written = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_EQ(XtiTable(written.value()), XtiTable(source));
  EXPECT_EQ(FutureRecords(Part(written.value(), sheet_part)), FutureRecords(Part(source, sheet_part)));
}

TEST(XlsbDeletedSheetSlot, ValidationListIntoADeletedSheetResaves) {
  // Sheet's list validation read Sheet2!$A$1:$A$3; Sheet2 was deleted. Table: [deleted, Sheet].
  ASSERT_EQ(XtiTable(Fixture("xlsb_deleted_sheet_slot_dv.xlsb")), (std::vector<Xti>{{0U, -1, -1}, {0U, 0, 0}}));
  auto loaded = read_xlsb(SpanOf(Fixture("xlsb_deleted_sheet_slot_dv.xlsb")));
  ASSERT_TRUE(static_cast<bool>(loaded));
  EXPECT_EQ(loaded.value().workbook.sheet(0).validations().size(), 1U);
  ExpectUneditedResaveKeepsTheDeadSlot("xlsb_deleted_sheet_slot_dv.xlsb", "xl/worksheets/sheet1.bin",
                                       "<xm:f>#REF!</xm:f>");
}

TEST(XlsbDeletedSheetSlot, ConditionalFormatIntoADeletedSheetResaves) {
  // Sheet2's rule read A1>Sheet!$A$1; Sheet was deleted. Table: [Sheet2, deleted].
  ASSERT_EQ(XtiTable(Fixture("xlsb_deleted_sheet_slot_cf.xlsb")), (std::vector<Xti>{{0U, 0, 0}, {0U, -1, -1}}));
  ExpectUneditedResaveKeepsTheDeadSlot("xlsb_deleted_sheet_slot_cf.xlsb", "xl/worksheets/sheet1.bin",
                                       "<xm:f>A1&gt;#REF!</xm:f>");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
