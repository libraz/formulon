// A workbook-scope built-in a same-name sheet-local name hides from its sheet keeps its binding through
// .xlsb and .xlsx round trips: stored `[0]!` in .xlsb, the bare name Excel stores in .xlsx.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Sheet1!E1 is `=ROWS(Print_Area)` bound to the workbook-scope Print_Area (3 rows); Sheet1 holds a local one (5 rows).
constexpr std::uint32_t kRow = 0U;
constexpr std::uint32_t kCol = 4U;

Workbook LoadFixture() {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/names/workbook_print_area_hidden_by_local.xlsb";
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(path);
  EXPECT_FALSE(bytes.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  if (!loaded) {
    return Workbook::create();
  }
  EXPECT_EQ(loaded.value().undecoded_formula_count, 0U);
  return std::move(loaded.value().workbook);
}

double RecalculatedNumber(Workbook& wb, std::uint32_t row, std::uint32_t col) {
  EXPECT_TRUE(static_cast<bool>(wb.recalc_full(eval::default_registry())));
  const Value value = wb.sheet(0).resolve_cell_value(row, col);
  EXPECT_TRUE(value.is_number());
  return value.is_number() ? value.as_number() : 0.0;
}

// The rgce of the BrtFmlaNum at `kCol` in sheet1.bin of the .xlsb package `bytes`.
std::vector<std::uint8_t> CellRgce(const std::vector<std::uint8_t>& bytes) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(test::span_of(bytes))));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(sheet));
  if (!sheet) {
    return {};
  }
  io::ByteSpan cursor = test::span_of(sheet.value());
  while (cursor.size > 0U) {
    auto record = io::xlsb::read_record(cursor);
    if (!record) {
      break;
    }
    if (record.value().type != static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtFmlaNum)) {
      continue;
    }
    io::ByteSpan p = record.value().payload;
    auto col = io::xlsb::read_u32(p);
    if (!col || col.value() != kCol) {
      continue;
    }
    // iStyleRef + fPhShow (4), the double (8), grbitFlags (2).
    p.data += 14;
    p.size -= 14;
    auto cce = io::xlsb::read_u32(p);
    if (!cce || cce.value() > p.size) {
      break;
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce.value());
  }
  return {};
}

// Excel's bytes for `=ROWS(Print_Area)` on the sheet whose local Print_Area hides the workbook one: PtgName (reference
// class) of the workbook-scope record, ilbl 2, then PtgFunc ROWS.
const std::vector<std::uint8_t> kExcelRgce = {0x23, 0x02, 0x00, 0x00, 0x00, 0x41, 0x4C, 0x00};

TEST(HiddenWorkbookName, XlsbFixtureStoresPtgNameOfTheWorkbookRecord) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/names/workbook_print_area_hidden_by_local.xlsb";
  EXPECT_EQ(CellRgce(test::read_file_bytes(path)), kExcelRgce);
}

TEST(HiddenWorkbookName, XlsbSaveWritesExcelsPtgName) {
  Workbook wb = LoadFixture();
  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(CellRgce(saved.value()), kExcelRgce);
}

TEST(HiddenWorkbookName, XlsbFormulaBindsTheWorkbookEntry) {
  Workbook wb = LoadFixture();
  const Cell* cell = wb.sheet(0).cell_at(kRow, kCol);
  ASSERT_NE(cell, nullptr);
  EXPECT_NE(cell->formula_text.find("[0]!"), std::string::npos) << cell->formula_text;
  EXPECT_EQ(RecalculatedNumber(wb, kRow, kCol), 3.0);
}

TEST(HiddenWorkbookName, XlsbRoundTripKeepsTheSelfBookForm) {
  Workbook wb = LoadFixture();
  const std::string before = wb.sheet(0).cell_at(kRow, kCol)->formula_text;

  auto saved = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  EXPECT_EQ(reloaded.value().undecoded_formula_count, 0U);
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(after.sheet(0).cell_at(kRow, kCol)->formula_text, before);
  EXPECT_EQ(RecalculatedNumber(after, kRow, kCol), 3.0);
}

TEST(HiddenWorkbookName, XlsxStoresTheBareNameAndReloadsBoundToTheLocalEntry) {
  Workbook wb = LoadFixture();

  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  std::string sheet_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.xml", &sheet_xml));
  EXPECT_NE(sheet_xml.find("ROWS(_xlnm.Print_Area)"), std::string::npos) << sheet_xml;
  EXPECT_EQ(sheet_xml.find("[0]!"), std::string::npos) << sheet_xml;

  auto reloaded = io::read_ooxml(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
  Workbook after = std::move(reloaded.value().workbook);
  EXPECT_EQ(RecalculatedNumber(after, kRow, kCol), 5.0);
}

}  // namespace
}  // namespace formulon
