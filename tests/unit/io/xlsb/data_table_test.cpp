// XLSB What-If data tables ([MS-XLSB] BrtTable, record 428): the record leads
// the anchor's plain value record (TABLECELL), survives a write -> read cycle
// from a model built here or loaded from xlsx, and the table evaluates again
// after the trip.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "support/ooxml_package_fixture.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Multiplication table: inputs F1 / F2, corner C3, row values D3:F3, column values C4:C6, results D4:F6.
Workbook BuildMultiplicationTable() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  EXPECT_TRUE(wb.set_cell_value(0, 0, 5, Value::number(1.0)));
  EXPECT_TRUE(wb.set_cell_value(0, 1, 5, Value::number(1.0)));
  EXPECT_TRUE(wb.set_cell_formula(0, 2, 2, "=F1*F2"));
  for (std::uint32_t k = 0; k < 3; ++k) {
    EXPECT_TRUE(wb.set_cell_value(0, 2, 3 + k, Value::number(static_cast<double>(k + 1))));
    EXPECT_TRUE(wb.set_cell_value(0, 3 + k, 2, Value::number(static_cast<double>(10 * (k + 1)))));
  }
  DataTableInput row_input;
  row_input.present = true;
  row_input.row = 0;
  row_input.col = 5;
  DataTableInput col_input = row_input;
  col_input.row = 1;
  EXPECT_TRUE(wb.add_data_table(0, MergeRange{3, 3, 5, 5}, row_input, col_input));
  auto stats = wb.recalc(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(stats));
  return wb;
}

std::vector<std::uint8_t> SheetBin(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(entry));
  return entry ? entry.value() : std::vector<std::uint8_t>();
}

std::uint32_t U32At(ByteSpan payload, std::size_t at) {
  return static_cast<std::uint32_t>(payload.data[at]) | (static_cast<std::uint32_t>(payload.data[at + 1]) << 8U) |
         (static_cast<std::uint32_t>(payload.data[at + 2]) << 16U) |
         (static_cast<std::uint32_t>(payload.data[at + 3]) << 24U);
}

TEST(XlsbDataTable, WritesTheTableRecordBeforeItsAnchorsValueRecord) {
  Workbook wb = BuildMultiplicationTable();
  auto bytes = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  const std::vector<std::uint8_t> sheet = SheetBin(bytes.value());
  ByteSpan records{sheet.data(), sheet.size()};
  std::size_t tables = 0;
  bool after_table = false;
  while (records.size != 0U) {
    auto rec = read_record(records);
    ASSERT_TRUE(static_cast<bool>(rec));
    const std::uint16_t type = rec.value().type;
    const ByteSpan payload = rec.value().payload;
    if (after_table) {
      // TABLECELL: the anchor (D4) follows as a plain number record.
      EXPECT_TRUE(type == static_cast<std::uint16_t>(XlsbRecordType::BrtCellRk) ||
                  type == static_cast<std::uint16_t>(XlsbRecordType::BrtCellReal))
          << type;
      EXPECT_EQ(U32At(payload, 0), 3U);
      after_table = false;
    }
    if (type != static_cast<std::uint16_t>(XlsbRecordType::BrtTable)) {
      continue;
    }
    ++tables;
    after_table = true;
    ASSERT_EQ(payload.size, 33U);
    EXPECT_EQ(U32At(payload, 0), 3U);   // rwFirst
    EXPECT_EQ(U32At(payload, 4), 5U);   // rwLast
    EXPECT_EQ(U32At(payload, 8), 3U);   // colFirst
    EXPECT_EQ(U32At(payload, 12), 5U);  // colLast
    EXPECT_EQ(U32At(payload, 16), 0U);  // rwInput1 / colInput1: F1
    EXPECT_EQ(U32At(payload, 20), 5U);
    EXPECT_EQ(U32At(payload, 24), 1U);  // rwInput2 / colInput2: F2
    EXPECT_EQ(U32At(payload, 28), 5U);
    EXPECT_EQ(payload.data[32], 0x13U);  // fRw, fTbl2, fAlwaysCalc
  }
  EXPECT_EQ(tables, 1U);
}

TEST(XlsbDataTable, SurvivesAWriteReadCycle) {
  Workbook wb = BuildMultiplicationTable();
  auto bytes = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  auto read = read_xlsb(SpanOf(bytes.value()));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  Workbook& back = read.value().workbook;

  ASSERT_EQ(back.sheet(0).data_tables().size(), 1U);
  const DataTable& table = back.sheet(0).data_tables()[0];
  const DataTable& original = wb.sheet(0).data_tables()[0];
  EXPECT_EQ(table.ref, original.ref);
  EXPECT_EQ(table.two_dimensional, original.two_dimensional);
  EXPECT_EQ(table.row_oriented, original.row_oriented);
  EXPECT_EQ(table.r1.row, 0U);
  EXPECT_EQ(table.r1.col, 5U);
  EXPECT_EQ(table.r2.row, 1U);
  EXPECT_EQ(table.r2.col, 5U);
  EXPECT_FALSE(table.r1.deleted);
  EXPECT_FALSE(table.r2.deleted);
  EXPECT_TRUE(table.always_calc);

  // The cached results came through as values, the anchor included, and are protected.
  for (std::uint32_t i = 0; i < 3; ++i) {
    for (std::uint32_t j = 0; j < 3; ++j) {
      const Cell* cell = back.sheet(0).cell_at(3 + i, 3 + j);
      ASSERT_NE(cell, nullptr);
      ASSERT_TRUE(cell->cached_value.is_number());
      EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), static_cast<double>(10 * (i + 1) * (j + 1)));
      EXPECT_TRUE(cell->formula_text.empty());
    }
  }
  EXPECT_FALSE(static_cast<bool>(back.set_cell_value(0, 4, 4, Value::number(1.0))));

  // And it computes again.
  ASSERT_TRUE(static_cast<bool>(back.set_cell_value(0, 2, 3, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(back.sheet(0).cell_at(3, 3)->cached_value.as_number(), 50.0);
}

TEST(XlsbDataTable, KeepsOneVariableOrientationAndDeletedInputs) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  // One-variable row table: input B6, values D1:E1, formula C2, results D2:E2.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 5, 1, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1, 2, "=B6+1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 3, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 4, Value::number(2.0))));
  DataTableInput row_input;
  row_input.present = true;
  row_input.row = 5;
  row_input.col = 1;
  ASSERT_TRUE(static_cast<bool>(wb.add_data_table(0, MergeRange{1, 3, 1, 4}, row_input, DataTableInput{})));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(1, 4)->cached_value.as_number(), 3.0);

  auto bytes = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  auto read = read_xlsb(SpanOf(bytes.value()));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  const DataTable& table = read.value().workbook.sheet(0).data_tables()[0];
  EXPECT_FALSE(table.two_dimensional);
  EXPECT_TRUE(table.row_oriented);
  EXPECT_TRUE(table.r1.present);
  EXPECT_FALSE(table.r2.present);
  EXPECT_EQ(table.r1.row, 5U);
  EXPECT_EQ(table.r1.col, 1U);
  EXPECT_EQ(data_table_formula_text(table), "=TABLE(B6,)");

  // A deleted input is written with fDeleted1 and its ignored cell zeroed.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 5, 1)));
  ASSERT_TRUE(wb.sheet(0).data_tables()[0].r1.deleted);
  auto deleted = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(deleted));
  auto reread = read_xlsb(SpanOf(deleted.value()));
  ASSERT_TRUE(static_cast<bool>(reread));
  const DataTable& gone = reread.value().workbook.sheet(0).data_tables()[0];
  EXPECT_TRUE(gone.r1.deleted);
  EXPECT_EQ(data_table_formula_text(gone), "=TABLE(#REF!,)");
}

// A sheet stream laid out per TABLECELL: BrtTable, then a BrtCellSt anchor.
// The second table's input cells are flagged deleted / unused and carry
// out-of-grid values the reader must ignore.
std::vector<std::uint8_t> HandBuiltTableSheet() {
  using reader_test_support::AppendRecord;
  using reader_test_support::AppendU32;
  using reader_test_support::AppendU8;
  using reader_test_support::AppendXLWideString;
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData
  const auto row_hdr = [&](std::uint32_t row) {
    std::vector<std::uint8_t> p;
    AppendU32(p, row);
    AppendRecord(body, 0, p);  // BrtRowHdr
  };
  const auto table = [&](std::uint32_t r0, std::uint32_t r1, std::uint32_t c0, std::uint32_t c1, std::uint32_t in_row,
                         std::uint32_t in_col, std::uint8_t flags) {
    std::vector<std::uint8_t> p;
    for (const std::uint32_t v : {r0, r1, c0, c1, in_row, in_col, 0xFFFFFFFFU, 0xFFFFFFFFU}) {
      AppendU32(p, v);
    }
    AppendU8(p, flags);
    AppendRecord(body, 428, p);  // BrtTable
  };
  const auto cell_st = [&](std::uint32_t col, std::string_view text) {
    std::vector<std::uint8_t> p;
    AppendU32(p, col);
    for (int i = 0; i < 4; ++i) {
      AppendU8(p, 0);  // style + fPhShow
    }
    AppendXLWideString(p, text);
    AppendRecord(body, 6, p);  // BrtCellSt
  };
  row_hdr(2);
  table(2, 4, 1, 1, 0, 1, 0x10);  // B3:B5, column input B1, fAlwaysCalc
  cell_st(1, "x");
  table(2, 3, 4, 4, 0xFFFFFFFFU, 0xFFFFFFFFU, 0x04);  // E3:E4, deleted column input
  cell_st(4, "y");
  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

TEST(XlsbDataTable, ReadsATableCellAndSavesItAgain) {
  namespace rts = reader_test_support;
  std::vector<rts::PartFile> parts;
  parts.push_back({"[Content_Types].xml", rts::StringToBytes(rts::ContentTypesXml())});
  parts.push_back({"_rels/.rels", rts::StringToBytes(rts::PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", rts::StringToBytes(rts::WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", rts::WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", HandBuiltTableSheet()});
  parts.push_back({"xl/worksheets/sheet2.bin", rts::SheetBinReal(0.0)});
  const std::vector<std::uint8_t> archive = rts::BuildZip(parts);
  auto read = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message << " | " << read.error().context;

  const auto check = [](const Workbook& wb) {
    const std::vector<DataTable>& tables = wb.sheet(0).data_tables();
    ASSERT_EQ(tables.size(), 2U);
    EXPECT_EQ(tables[0].ref, (MergeRange{2, 1, 4, 1}));
    EXPECT_FALSE(tables[0].two_dimensional);
    EXPECT_FALSE(tables[0].row_oriented);
    EXPECT_TRUE(tables[0].always_calc);
    EXPECT_EQ(tables[0].r1.row, 0U);
    EXPECT_EQ(tables[0].r1.col, 1U);
    EXPECT_FALSE(tables[0].r2.present);
    EXPECT_EQ(data_table_formula_text(tables[0]), "=TABLE(,B1)");
    EXPECT_TRUE(tables[1].r1.deleted);
    EXPECT_FALSE(tables[1].always_calc);
    EXPECT_EQ(data_table_formula_text(tables[1]), "=TABLE(,#REF!)");
    for (const auto& [col, text] : {std::pair<std::uint32_t, const char*>{1U, "x"}, {4U, "y"}}) {
      const Cell* anchor = wb.sheet(0).cell_at(2, col);
      ASSERT_NE(anchor, nullptr);
      ASSERT_TRUE(anchor->cached_value.is_text());
      EXPECT_EQ(anchor->cached_value.as_text(), text);
      EXPECT_TRUE(anchor->formula_text.empty());
    }
  };
  check(read.value().workbook);
  EXPECT_EQ(read.value().dropped_record_count, 0U);

  // xlsb -> xlsb: the text anchor goes out as BrtCellSt and the tables come back.
  auto bytes = write_xlsb(read.value().workbook);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  auto again = read_xlsb(SpanOf(bytes.value()));
  ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
  check(again.value().workbook);
}

TEST(XlsbDataTable, AnXlsxTableSurvivesAnXlsbSave) {
  // One-variable column table: input B1, formula B2, values A3:A5, results B3:B5.
  const std::string sheet_xml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
      "<row r=\"1\"><c r=\"B1\"><v>3</v></c></row>"
      "<row r=\"2\"><c r=\"B2\"><f>B1*B1</f><v>9</v></c></row>"
      "<row r=\"3\"><c r=\"A3\"><v>1</v></c><c r=\"B3\"><f t=\"dataTable\" ref=\"B3:B5\" dt2D=\"0\" "
      "dtr=\"0\" r1=\"B1\" ca=\"1\"/><v>1</v></c></row>"
      "<row r=\"4\"><c r=\"A4\"><v>2</v></c><c r=\"B4\"><v>4</v></c></row>"
      "<row r=\"5\"><c r=\"A5\"><v>3</v></c><c r=\"B5\"><v>9</v></c></row>"
      "</sheetData></worksheet>";
  const std::vector<std::uint8_t> package =
      test::BuildOoxmlPackage(test::OoxmlContentTypes(test::kXlsxWorkbookContentType), sheet_xml);
  auto loaded = read_ooxml(SpanOf(package));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  ASSERT_EQ(loaded.value().workbook.sheet(0).data_tables().size(), 1U);

  auto bytes = write_xlsb(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  auto read = read_xlsb(SpanOf(bytes.value()));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  Workbook& back = read.value().workbook;
  ASSERT_EQ(back.sheet(0).data_tables().size(), 1U);
  const DataTable& table = back.sheet(0).data_tables()[0];
  EXPECT_EQ(table.ref, (MergeRange{2, 1, 4, 1}));
  EXPECT_FALSE(table.row_oriented);
  EXPECT_TRUE(table.always_calc);
  EXPECT_EQ(table.r1.row, 0U);
  EXPECT_EQ(table.r1.col, 1U);
  EXPECT_DOUBLE_EQ(back.sheet(0).cell_at(4, 1)->cached_value.as_number(), 9.0);

  ASSERT_TRUE(static_cast<bool>(back.set_cell_formula(0, 1, 1, "=B1*10")));
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(back.sheet(0).cell_at(4, 1)->cached_value.as_number(), 30.0);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
