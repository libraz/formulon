//
// A structured reference into another workbook's table, in xlsb, held to Excel's own save of
// `external_table_ref_closed`: a nonresident PtgList whose table and column names sit in the formula's rgcb and whose
// XTI names the link's file-stem sheet entry. The fixture's `Sheet1` holds
//   A1  =ROWS([1]!Tbl[#Data])        A2  =SUM([1]!Tbl[v])
//   A3  =ISREF([1]!Tbl[v])           B3  =ISREF([src.xlsx]Sheet1!Tbl[v]), stored as A3 is
//   A4  {=[1]!Tbl[v]}                (legacy array formula)

#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <tuple>
#include <vector>

#include "cell.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> FixtureBytes(const char* name) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(file);
  return out;
}

Workbook ReadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto read = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

Workbook ReadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

std::string ZipPart(const std::vector<std::uint8_t>& package, const char* name) {
  mz_zip_archive zip{};
  if (mz_zip_reader_init_mem(&zip, package.data(), package.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "not a zip package";
    return {};
  }
  std::size_t size = 0;
  void* data = mz_zip_reader_extract_file_to_heap(&zip, name, &size, 0);
  std::string out = data != nullptr ? std::string(static_cast<const char*>(data), size) : std::string();
  mz_free(data);
  mz_zip_reader_end(&zip);
  return out;
}

using CellKey = std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>;  // row, column, record type

/// Every formula cell record of a sheet part (BrtFmlaBool 10, BrtFmlaError 11, BrtArrFmla 426), keyed by its cell and
/// record type, holding its `CellParsedFormula` (cce, rgce, cb, rgcb).
std::map<CellKey, std::string> FormulaRecords(const std::string& part) {
  std::map<CellKey, std::string> out;
  std::size_t at = 0;
  const auto varint = [&](int max_bytes) {
    std::uint32_t value = 0;
    for (int i = 0; i < max_bytes && at < part.size(); ++i) {
      const auto byte = static_cast<std::uint8_t>(part[at++]);
      value |= static_cast<std::uint32_t>(byte & 0x7FU) << (7 * i);
      if ((byte & 0x80U) == 0U) {
        break;
      }
    }
    return value;
  };
  const auto u32 = [&part](std::size_t offset) {
    std::uint32_t value = 0;
    for (int i = 3; i >= 0; --i) {
      value = (value << 8U) | static_cast<std::uint8_t>(part[offset + static_cast<std::size_t>(i)]);
    }
    return value;
  };
  std::uint32_t row = 0;
  while (at < part.size()) {
    const std::uint32_t type = varint(2);
    const std::uint32_t size = varint(4);
    if (type == 0U) {  // BrtRowHdr
      row = u32(at);
    } else if (type == 10U || type == 11U) {  // Cell (8), a one-byte value, grbitFlags (2), formula
      out[{row, u32(at), type}] = part.substr(at + 11U, size - 11U);
    } else if (type == 426U) {  // RfX (16), a flag byte, formula
      out[{u32(at), u32(at + 8U), type}] = part.substr(at + 17U, size - 17U);
    }
    at += size;
  }
  return out;
}

std::string FormulaAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(0).cell_at(row, col);
  return cell == nullptr ? std::string() : cell->formula_text;
}

constexpr std::uint32_t kFormulaCells[][2] = {{0U, 0U}, {1U, 0U}, {2U, 0U}, {2U, 1U}, {3U, 0U}};

TEST(ExternalTablePtgList, ExcelsXlsbReadsAsItsXlsxTwin) {
  const Workbook from_xlsb = ReadXlsb(FixtureBytes("external_table_ref_closed.xlsb"));
  const Workbook from_xlsx = ReadXlsx(FixtureBytes("external_table_ref_closed.xlsx"));
  for (const auto& cell : kFormulaCells) {
    const std::string text = FormulaAt(from_xlsx, cell[0], cell[1]);
    EXPECT_NE(text.find("Tbl["), std::string::npos) << text;
    EXPECT_EQ(FormulaAt(from_xlsb, cell[0], cell[1]), text) << cell[0] << "," << cell[1];
  }
}

TEST(ExternalTablePtgList, SavingTheXlsxWritesExcelsTokens) {
  const std::vector<std::uint8_t> excel = FixtureBytes("external_table_ref_closed.xlsb");
  const std::map<CellKey, std::string> want = FormulaRecords(ZipPart(excel, "xl/worksheets/sheet1.bin"));
  ASSERT_EQ(want.size(), 6U);

  auto saved = io::xlsb::write_xlsb_with_result(ReadXlsx(FixtureBytes("external_table_ref_closed.xlsx")));
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 0U);
  EXPECT_EQ(FormulaRecords(ZipPart(saved.value().bytes, "xl/worksheets/sheet1.bin")), want);
}

TEST(ExternalTablePtgList, AnXlsbSaveReadsBackInBothFormats) {
  const Workbook from_excel = ReadXlsb(FixtureBytes("external_table_ref_closed.xlsb"));
  auto xlsb = io::xlsb::write_xlsb(from_excel);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  const Workbook reread = ReadXlsb(xlsb.value());
  auto xlsx = io::write_ooxml(reread);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const Workbook as_xlsx = ReadXlsx(xlsx.value());
  for (const auto& cell : kFormulaCells) {
    EXPECT_EQ(FormulaAt(reread, cell[0], cell[1]), FormulaAt(from_excel, cell[0], cell[1]));
    EXPECT_EQ(FormulaAt(as_xlsx, cell[0], cell[1]), FormulaAt(from_excel, cell[0], cell[1]));
  }
}

// Only every data row and one plain column have known bytes; any other form is saved as its cached value.
TEST(ExternalTablePtgList, AnUnmeasuredFormIsSavedAsItsValue) {
  for (const char* payload : {"#Headers", "#All", "[#Data],[v]", "[v]:[w]", "@v", "a'#b"}) {
    Workbook wb = ReadXlsx(FixtureBytes("external_table_ref_closed.xlsx"));
    // A2 is `=SUM(<table>[v])`.
    const std::string sum = FormulaAt(wb, 1U, 0U);
    const std::string formula = "=ROWS(" + sum.substr(5U, sum.find("Tbl[") + 3U - 5U) + "[" + payload + "])";
    wb.sheet(0).set_cell_formula(5U, 0U, formula);
    auto saved = io::xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 1U) << formula;
  }
}

// The XTI names the sheet entry spelled as the supporting file's stem; without one the reference is not encoded.
TEST(ExternalTablePtgList, ALinkWithoutItsFileStemSheetSavesTheValue) {
  Workbook wb = ReadXlsx(FixtureBytes("external_table_ref_closed.xlsx"));
  std::vector<ExternalLinkRecord> links = wb.external_links();
  ASSERT_EQ(links.size(), 1U);
  ASSERT_EQ(links[0].book.sheet_names, (std::vector<std::string>{"Sheet1", "src"}));
  links[0].book.sheet_names[1] = "source";
  wb.set_external_links(std::move(links));
  auto saved = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 5U);
}

}  // namespace
}  // namespace formulon
