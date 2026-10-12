//
// Pins how both writers merge a row's stored cells with the spill phantoms
// covering it: a dynamic-array anchor at A1 spilling A1:C1, a style-only blank
// under the spill (B1), a stale literal under the spill (C1) and a value
// outside it (E1). The expected row text and record sequence are the writers'
// output for that row, so a change to how stored cells and phantoms combine
// shows up here.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kStyledXf = 1U;

/// A1 = SEQUENCE(1,3) spilling A1:C1, B1 a style-only blank, C1 a literal
/// written under the committed spill, E1 a value outside it.
Workbook BuildScenario() {
  Workbook wb = Workbook::create_empty();
  (void)wb.add_sheet("Sheet1");
  Sheet& s = wb.sheet(0U);
  s.set_cell_value(0U, 1U, Value::blank());
  s.set_cell_xf_index(0U, 1U, kStyledXf);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=SEQUENCE(1,3)")));
  s.set_cell_dynamic_array(0U, 0U, true);
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  s.set_cell_cached_value(0U, 2U, Value::number(99.0));
  s.set_cell_value(0U, 4U, Value::number(7.0));
  return wb;
}

/// "type@col/xf" for each cell record of row 0 and "type" for the other
/// records between its BrtRowHdr and the next one.
std::vector<std::string> RowZeroRecords(const std::vector<std::uint8_t>& part) {
  std::vector<std::string> out;
  io::ByteSpan cursor{part.data(), part.size()};
  bool in_row0 = false;
  auto u32 = [](const std::uint8_t* d) {
    return static_cast<std::uint32_t>(d[0]) | (static_cast<std::uint32_t>(d[1]) << 8) |
           (static_cast<std::uint32_t>(d[2]) << 16) | (static_cast<std::uint32_t>(d[3]) << 24);
  };
  while (cursor.size > 0U) {
    auto record = io::xlsb::read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const io::ByteSpan p = record.value().payload;
    const std::uint16_t type = record.value().type;
    if (type == static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtRowHdr)) {
      if (in_row0) {
        break;
      }
      in_row0 = p.size >= 4U && u32(p.data) == 0U;
      if (in_row0) {
        out.push_back("RowHdr");
      }
    } else if (type == static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtEndSheetData)) {
      break;
    } else if (in_row0) {
      const bool cell = type >= static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtCellBlank) &&
                        type <= static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtFmlaError);
      if (cell && p.size >= 8U) {
        out.push_back(std::to_string(type) + "@" + std::to_string(u32(p.data)) + "/" +
                      std::to_string(u32(p.data + 4U) & 0xFFFFFFU));
      } else {
        out.push_back(std::to_string(type));
      }
    }
  }
  return out;
}

TEST(RowMergeWriter, OoxmlRowMergesStoredCellsWithSpillPhantoms) {
  Workbook wb = BuildScenario();
  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  std::string xml;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.xml", &xml));
  const std::size_t begin = xml.find("<row ");
  ASSERT_NE(begin, std::string::npos) << xml;
  const std::size_t end = xml.find("</row>", begin);
  ASSERT_NE(end, std::string::npos) << xml;
  // The phantom wins over the stored cell on B1 (keeping its style) and on C1
  // (the stale literal 99 is not written).
  EXPECT_EQ(xml.substr(begin, end + 6U - begin),
            "<row r=\"1\"><c r=\"A1\" cm=\"1\"><f t=\"array\" ref=\"A1:C1\">_xlfn.SEQUENCE(1,3)</f><v>1</v></c>"
            "<c r=\"B1\" s=\"1\"><v>2</v></c><c r=\"C1\"><v>3</v></c><c r=\"E1\"><v>7</v></c></row>");
}

TEST(RowMergeWriter, XlsbRowMergesStoredCellsWithSpillPhantoms) {
  Workbook wb = BuildScenario();
  auto saved = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{saved.value().bytes.data(), saved.value().bytes.size()})));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet));
  std::string joined;
  for (const std::string& r : RowZeroRecords(sheet.value())) {
    joined += r + " ";
  }
  // BrtCellMeta and BrtArrFmla (426) follow the anchor; B1 and C1 are formula
  // records for the phantoms, B1 with the stored xf; E1 is a BrtCellRk (2).
  EXPECT_EQ(joined, "RowHdr 49 9@0/0 426 9@1/1 9@2/0 2@4/0 ");
}

}  // namespace
}  // namespace formulon
