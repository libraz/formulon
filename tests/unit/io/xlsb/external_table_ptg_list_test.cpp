//
// A structured reference into another workbook's table, in xlsb, held to Excel's own save of
// `external_table_ref_closed`: a nonresident PtgList whose table and column names sit in the formula's rgcb and whose
// XTI names the link's file-stem sheet entry. The fixture's `Sheet1` holds
//   A1  =ROWS([1]!Tbl[#Data])        A2  =SUM([1]!Tbl[v])
//   A3  =ISREF([1]!Tbl[v])           B3  =ISREF([src.xlsx]Sheet1!Tbl[v]), stored as A3 is
//   A4  {=[1]!Tbl[v]}                (legacy array formula)
// and `external_table_forms/` holds Excel's saves of every other measured form, each xlsb with its xlsx twin.

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

/// Every formula cell record of a sheet part (BrtFmlaBool 10, BrtFmlaError 11, BrtArrFmla 426, and BrtFmlaNum 9 when
/// `numbers`), keyed by its cell and record type, holding its `CellParsedFormula` (cce, rgce, cb, rgcb).
std::map<CellKey, std::string> FormulaRecords(const std::string& part, bool numbers = false) {
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
    } else if (type == 9U && numbers) {  // Cell (8), Xnum (8), grbitFlags (2), formula
      out[{row, u32(at), type}] = part.substr(at + 18U, size - 18U);
    } else if (type == 426U) {  // RfX (16), a flag byte, formula
      out[{u32(at), u32(at + 8U), type}] = part.substr(at + 17U, size - 17U);
    }
    at += size;
  }
  return out;
}

/// Every record of `wanted` type in a part, in order.
std::vector<std::string> RecordPayloads(const std::string& part, std::uint32_t wanted) {
  std::vector<std::string> out;
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
  while (at < part.size()) {
    const std::uint32_t type = varint(2);
    const std::uint32_t size = varint(4);
    if (type == wanted) {
      out.push_back(part.substr(at, size));
    }
    at += size;
  }
  return out;
}

constexpr std::uint32_t kBrtExternSheet = 362U;
constexpr std::uint32_t kBrtSupTabs = 359U;
constexpr std::uint32_t kBrtBeginExternTable = 363U;

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

// A totals-row form and a `#This Row` with other than one column have no known bytes; they are saved as their value.
TEST(ExternalTablePtgList, AnUnmeasuredFormIsSavedAsItsValue) {
  for (const char* payload : {"#Totals", "[#Data],[#Totals]", "[#Totals],[v]", "#This Row", "[#This Row],[v]:[w]"}) {
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

// A sheet entry spelled as the supporting file's stem that is not the last one has no measured save; every table
// reference through that link is saved as its value.
TEST(ExternalTablePtgList, AStemSheetThatIsNotLastSavesTheValue) {
  Workbook wb = ReadXlsx(FixtureBytes("external_table_ref_closed.xlsx"));
  std::vector<ExternalLinkRecord> links = wb.external_links();
  ASSERT_EQ(links.size(), 1U);
  ASSERT_EQ(links[0].book.sheet_names, (std::vector<std::string>{"Sheet1", "src"}));
  links[0].book.sheet_names = {"src", "Sheet1"};
  wb.set_external_links(std::move(links));
  auto saved = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 5U);
}

// Excel's saves of every measured form, the link's source closed: each xlsb next to its xlsx twin.
constexpr const char* kFormBooks[] = {"selectors", "column_ranges", "column_names", "two_tables",
                                      "table_on_second_sheet"};

std::vector<std::uint8_t> FormBook(const std::string& book, const char* extension) {
  return FixtureBytes(("external_table_forms/" + book + extension).c_str());
}

bool IsTotalsForm(const std::string& formula) {
  return formula.find("#Totals") != std::string::npos;
}

TEST(ExternalTablePtgList, EveryMeasuredFormReadsAsItsXlsxTwin) {
  for (const char* book : kFormBooks) {
    SCOPED_TRACE(book);
    const Workbook from_xlsb = ReadXlsb(FormBook(book, ".xlsb"));
    const Workbook from_xlsx = ReadXlsx(FormBook(book, ".xlsx"));
    std::size_t compared = 0;
    for (const CellAddress& at : from_xlsx.sheet(0).all_formula_cells()) {
      const std::string text = FormulaAt(from_xlsx, at.row, at.col);
      if (IsTotalsForm(text)) {
        continue;
      }
      EXPECT_EQ(FormulaAt(from_xlsb, at.row, at.col), text) << at.row << "," << at.col;
      ++compared;
    }
    EXPECT_GT(compared, 0U);
  }
}

TEST(ExternalTablePtgList, SavingEveryMeasuredFormWritesExcelsTokens) {
  for (const char* book : kFormBooks) {
    SCOPED_TRACE(book);
    const Workbook from_xlsx = ReadXlsx(FormBook(book, ".xlsx"));
    std::map<CellKey, std::string> want = FormulaRecords(ZipPart(FormBook(book, ".xlsb"), "xl/worksheets/sheet1.bin"));
    std::size_t totals = 0;
    for (const CellAddress& at : from_xlsx.sheet(0).all_formula_cells()) {
      if (IsTotalsForm(FormulaAt(from_xlsx, at.row, at.col))) {
        ++totals;
        for (auto it = want.begin(); it != want.end();) {
          it = std::get<0>(it->first) == at.row && std::get<1>(it->first) == at.col ? want.erase(it) : std::next(it);
        }
      }
    }
    auto saved = io::xlsb::write_xlsb_with_result(from_xlsx);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, totals);
    EXPECT_EQ(FormulaRecords(ZipPart(saved.value().bytes, "xl/worksheets/sheet1.bin")), want);
  }
}

// `[@v]` is stored as its long form, a plain formula: one row of one column does not spill.
TEST(ExternalTablePtgList, AnEnteredThisRowFormIsStoredLongAndPlain) {
  Workbook wb = ReadXlsx(FormBook("selectors", ".xlsx"));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 80U, 0U, "=src_a.xlsx!Tbl[@v]")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 81U, 0U, "=src_a.xlsx!Tbl[#Headers]")));
  auto xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const std::string sheet = ZipPart(xlsx.value(), "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet.find("<f>[1]!Tbl[[#This Row],[v]]</f>"), std::string::npos) << sheet;
  EXPECT_NE(sheet.find("<f t=\"array\" ref=\"A82\">[1]!Tbl[#Headers]</f>"), std::string::npos) << sheet;
}

// Entered with the source open, A1 `=SUM([1]!Tbl[v])` links a book whose sheet list is the source's own, with no
// entry spelled as the file's stem: `entered_open` (Sheet1), `entered_open_multi_a` (Sheet1, Data, Sheet3; the table
// on Data) and `entered_open_multi_b` (Sheet1, Sheet2, Data). Excel reopening each with the source absent saves an
// xlsb (`<book>_reopened`) whose link gains the stem as its last sheet entry, refresh failed, and whose XTI names it.
// The save leaves the workbook's link as loaded; only the saved part carries the entry.
TEST(ExternalTablePtgList, AReopenedLinkNamesItsTableThroughAnAppendedStemSheet) {
  for (const char* book : {"entered_open", "entered_open_multi_a", "entered_open_multi_b"}) {
    SCOPED_TRACE(book);
    const Workbook wb = ReadXlsx(FormBook(book, ".xlsx"));
    ASSERT_EQ(wb.external_links().size(), 1U);
    auto saved = io::xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 0U);
    EXPECT_EQ(wb.external_links()[0].book.sheet_names,
              ReadXlsx(FormBook(book, ".xlsx")).external_links()[0].book.sheet_names);
    const std::vector<std::uint8_t> excel = FormBook(std::string(book) + "_reopened", ".xlsb");
    const std::vector<std::uint8_t>& ours = saved.value().bytes;
    const char* const link = "xl/externalLinks/externalLink1.bin";
    EXPECT_EQ(RecordPayloads(ZipPart(ours, link), kBrtSupTabs), RecordPayloads(ZipPart(excel, link), kBrtSupTabs));
    EXPECT_EQ(RecordPayloads(ZipPart(ours, link), kBrtBeginExternTable),
              RecordPayloads(ZipPart(excel, link), kBrtBeginExternTable));
    EXPECT_EQ(RecordPayloads(ZipPart(ours, "xl/workbook.bin"), kBrtExternSheet),
              RecordPayloads(ZipPart(excel, "xl/workbook.bin"), kBrtExternSheet));
    EXPECT_EQ(FormulaRecords(ZipPart(ours, "xl/worksheets/sheet1.bin"), /*numbers=*/true),
              FormulaRecords(ZipPart(excel, "xl/worksheets/sheet1.bin"), /*numbers=*/true));
  }
}

TEST(ExternalTablePtgList, ALinkWithNoStemSheetEntryKeepsItsXlsxFormula) {
  const std::vector<std::uint8_t> excel = FormBook("entered_open", ".xlsx");
  auto saved = io::write_ooxml(ReadXlsx(excel));
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  const std::string sheet = ZipPart(saved.value(), "xl/worksheets/sheet1.xml");
  EXPECT_NE(sheet.find("<f>SUM([1]!Tbl[v])</f>"), std::string::npos) << sheet;
  const auto sheet_names = [](const std::string& part) {
    const std::size_t begin = part.find("<sheetNames>");
    const std::size_t end = part.find("</sheetNames>");
    return begin == std::string::npos || end == std::string::npos ? std::string() : part.substr(begin, end - begin);
  };
  const char* const link = "xl/externalLinks/externalLink1.xml";
  EXPECT_EQ(sheet_names(ZipPart(saved.value(), link)), sheet_names(ZipPart(excel, link)));
  EXPECT_FALSE(sheet_names(ZipPart(excel, link)).empty());
}

// A link with no sheet entry whose only use is A2 `=SUM([1]!Tbl[v])`, the source absent: `no_sheet_entry` (missing
// path, Target `My Sheet`) and `no_sheet_entry_file` (Target `Zero.xlsx`). Excel's saves (`<book>_saved`) keep the
// formula and give the link the target's stem as its one sheet, refresh failed; in xlsb the XTI names it. The save
// leaves the workbook's link as loaded. Excel resolved `no_sheet_entry_file`'s target against the folder the authored
// book records, so the link targets are not compared.
TEST(ExternalTablePtgList, ALinkWithNoSheetEntrySavesTheStemAsItsOneSheet) {
  const auto sheet_entries = [](const std::string& part) {
    const std::size_t begin = part.find("<sheetNames>");
    const std::size_t end = part.find("</sheetDataSet>");
    return begin == std::string::npos || end == std::string::npos ? std::string() : part.substr(begin, end - begin);
  };
  for (const char* book : {"no_sheet_entry", "no_sheet_entry_file"}) {
    SCOPED_TRACE(book);
    const Workbook wb = ReadXlsx(FormBook(book, ".xlsx"));
    ASSERT_EQ(wb.external_links().size(), 1U);
    const std::string excel_name = std::string(book) + "_saved";

    auto xlsb = io::xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
    EXPECT_EQ(xlsb.value().diagnostics.downgraded_formula_count, 0U);
    const std::vector<std::uint8_t> excel_xlsb = FormBook(excel_name, ".xlsb");
    const std::vector<std::uint8_t>& ours = xlsb.value().bytes;
    const char* const bin_link = "xl/externalLinks/externalLink1.bin";
    EXPECT_EQ(RecordPayloads(ZipPart(ours, bin_link), kBrtSupTabs),
              RecordPayloads(ZipPart(excel_xlsb, bin_link), kBrtSupTabs));
    EXPECT_EQ(RecordPayloads(ZipPart(ours, bin_link), kBrtBeginExternTable),
              RecordPayloads(ZipPart(excel_xlsb, bin_link), kBrtBeginExternTable));
    EXPECT_EQ(RecordPayloads(ZipPart(ours, "xl/workbook.bin"), kBrtExternSheet),
              RecordPayloads(ZipPart(excel_xlsb, "xl/workbook.bin"), kBrtExternSheet));
    EXPECT_EQ(FormulaRecords(ZipPart(ours, "xl/worksheets/sheet1.bin"), /*numbers=*/true),
              FormulaRecords(ZipPart(excel_xlsb, "xl/worksheets/sheet1.bin"), /*numbers=*/true));

    auto xlsx = io::write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
    const std::vector<std::uint8_t> excel_xlsx = FormBook(excel_name, ".xlsx");
    const char* const xml_link = "xl/externalLinks/externalLink1.xml";
    EXPECT_NE(ZipPart(xlsx.value(), "xl/worksheets/sheet1.xml").find("<f>SUM([1]!Tbl[v])</f>"), std::string::npos);
    EXPECT_EQ(sheet_entries(ZipPart(xlsx.value(), xml_link)), sheet_entries(ZipPart(excel_xlsx, xml_link)));
    EXPECT_FALSE(sheet_entries(ZipPart(excel_xlsx, xml_link)).empty());
    EXPECT_TRUE(wb.external_links()[0].book.sheet_names.empty());

    // A reload holds the stem as a sheet entry, and its save writes the same one.
    auto again = io::write_ooxml(ReadXlsx(xlsx.value()));
    ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
    EXPECT_EQ(sheet_entries(ZipPart(again.value(), xml_link)), sheet_entries(ZipPart(excel_xlsx, xml_link)));
  }
}

// A missing-path link whose target and only sheet are `My Sheet` (Target `My%20Sheet`): Excel's save holds
//   A1  ='[1]My Sheet'!A1        A2  =SUM([1]!Tbl[v])
// and both reach the link's one sheet entry through one XTI.
TEST(ExternalTablePtgList, AMissingPathLinkSavesExcelsTokens) {
  const std::vector<std::uint8_t> excel = FormBook("path_missing_space", ".xlsb");
  for (const char* extension : {".xlsx", ".xlsb"}) {
    SCOPED_TRACE(extension);
    const std::vector<std::uint8_t> source = FormBook("path_missing_space", extension);
    const Workbook wb = std::string(extension) == ".xlsb" ? ReadXlsb(source) : ReadXlsx(source);
    auto saved = io::xlsb::write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    EXPECT_EQ(saved.value().diagnostics.downgraded_formula_count, 0U);
    EXPECT_EQ(FormulaRecords(ZipPart(saved.value().bytes, "xl/worksheets/sheet1.bin"), /*numbers=*/true),
              FormulaRecords(ZipPart(excel, "xl/worksheets/sheet1.bin"), /*numbers=*/true));
    EXPECT_EQ(RecordPayloads(ZipPart(saved.value().bytes, "xl/workbook.bin"), kBrtExternSheet),
              RecordPayloads(ZipPart(excel, "xl/workbook.bin"), kBrtExternSheet));
  }
}

}  // namespace
}  // namespace formulon
