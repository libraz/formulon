//
// Checks cross-workbook reference support against real Mac Excel
// 365-produced packages, in both container formats.
//
// Each fixture carries both engines' answers at once: Excel evaluated
// every formula while the supporting workbooks were open and saved the
// results into the cached cell values, so the expectations below are
// direct comparisons against what Excel produced rather than numbers
// transcribed into this file. The cached value is exactly what Excel
// itself shows once the source is closed, which is the state the engine
// reproduces.
//
// The two fixture pairs are the same workbook saved twice, so the xlsx
// and xlsb readers are held to one answer:
//
//   `external_link_mixed.{xlsx,xlsb}` — `Use` sheet, over two supporting
//     workbooks, mixing internal and external references so the
//     supporting-book table has to be read rather than assumed:
//       A1  =Local!A1                  (internal, same book)
//       A2  =SUM(Local!A1:A3)          (internal, same book)
//       A3  =[1]Data!A1                (external sheet reference)
//       A4  =[1]!SrcTotal              (external book-scope name)
//       A5  =[2]!FarCell               (a second supporting workbook)
//
//   `external_link_cell_kinds.{xlsx,xlsb}` — `Use` sheet, over one
//     supporting workbook with two sheets, covering every cached cell
//     kind and both name shapes:
//       A1  =[1]!OnSecond              (name -> a cell on the 2nd sheet)
//       A2  =SUM([1]!SecondRange)      (name -> a rectangle)
//       A3  =[1]Second!B5              (cached boolean)
//       A4  =[1]Second!B6              (cached error)
//       A5  =[1]Second!B4              (never cached -> Excel reads 0)
//       A6  =[1]Second!B7              (cached text)
//
// A5 is the case worth stating twice: Excel caches only the cells the
// consuming workbook references, and reads an address it does not hold
// as numeric zero rather than as blank or `#REF!`.

#include <cstdint>
#include <cstdio>
#include <functional>
#include <map>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_book.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    const std::size_t read = std::fread(out.data(), 1, out.size(), file);
    if (read != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

/// Loads `<stem>.xlsx` or `<stem>.xlsb`, dispatching on the extension so
/// both readers run against the same expectations below.
Workbook LoadFixture(const std::string& stem, const std::string& extension) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + stem + extension;
  const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  const io::ByteSpan span{bytes.data(), bytes.size()};
  if (extension == ".xlsb") {
    auto result_or = io::xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(result_or)) << path << ": " << (result_or ? "" : result_or.error().message);
    if (!result_or) {
      return Workbook::create_empty();
    }
    return std::move(result_or.value().workbook);
  }
  auto result_or = io::read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(result_or)) << path << ": " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

/// The cached value Excel stored in `Use!A<row>` (1-based row).
Value ExcelAnswer(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row - 1U, 0U);
  if (cell == nullptr) {
    ADD_FAILURE() << "no cell at Use!A" << row;
    return Value::blank();
  }
  return cell->cached_value;
}

/// Recalculates and returns what the engine computed for `Use!A<row>`.
Value OurAnswer(Workbook& wb, std::uint32_t row) {
  auto recalc_or = wb.recalc(eval::default_registry());
  EXPECT_TRUE(static_cast<bool>(recalc_or)) << (recalc_or ? "" : recalc_or.error().message);
  const Cell* cell = wb.sheet(0).cell_at(row - 1U, 0U);
  if (cell == nullptr) {
    ADD_FAILURE() << "no cell at Use!A" << row;
    return Value::blank();
  }
  return cell->cached_value;
}

void ExpectSameAsExcel(const Value& ours, const Value& excel, const char* label) {
  ASSERT_EQ(ours.kind(), excel.kind()) << label;
  switch (excel.kind()) {
    case ValueKind::Number:
      EXPECT_DOUBLE_EQ(ours.as_number(), excel.as_number()) << label;
      break;
    case ValueKind::Bool:
      EXPECT_EQ(ours.as_boolean(), excel.as_boolean()) << label;
      break;
    case ValueKind::Error:
      EXPECT_EQ(ours.as_error(), excel.as_error()) << label;
      break;
    case ValueKind::Text:
      EXPECT_EQ(ours.as_text(), excel.as_text()) << label;
      break;
    default:
      ADD_FAILURE() << label << ": unexpected cached kind";
      break;
  }
}

// ---------------------------------------------------------------------------
// (a) The external link body reaches the model.
// ---------------------------------------------------------------------------

class ExternalLinkFixture : public ::testing::TestWithParam<const char*> {};

TEST_P(ExternalLinkFixture, SupportingWorkbookCachesReachTheModel) {
  Workbook wb = LoadFixture("external_link_mixed", GetParam());
  // Two supporting workbooks, in the order the `[N]` prefixes select
  // them. Getting this order wrong would bind `[2]` to the first book.
  ASSERT_EQ(wb.external_links().size(), 2U);

  const ExternalBook& first = wb.external_links()[0].book;
  ASSERT_EQ(first.sheet_names.size(), 1U);
  EXPECT_EQ(first.sheet_names[0], "Data");
  ASSERT_NE(first.find_name("SrcTotal"), nullptr);
  EXPECT_TRUE(first.find_name("SrcTotal")->resolvable);
  // Excel caches only the cells this workbook references: A1 for the
  // direct reference and A3 for the one the name resolves to.
  EXPECT_TRUE(first.cached_cell(0, 0, 0).is_number());
  EXPECT_DOUBLE_EQ(first.cached_cell(0, 2, 0).as_number(), 30.0);

  const ExternalBook& second = wb.external_links()[1].book;
  ASSERT_NE(second.find_name("FarCell"), nullptr);
  EXPECT_DOUBLE_EQ(second.cached_cell(0, 6, 3).as_number(), 77.0);
}

TEST_P(ExternalLinkFixture, AnUncachedAddressReadsAsZeroNotBlank) {
  Workbook wb = LoadFixture("external_link_cell_kinds", GetParam());
  ASSERT_EQ(wb.external_links().size(), 1U);
  const ExternalBook& book = wb.external_links()[0].book;
  const std::uint32_t second = book.sheet_index("Second");
  ASSERT_NE(second, ExternalBook::kNoSheet);
  // B4 (row index 3) was never cached; B5 was.
  const Value uncached = book.cached_cell(second, 3, 1);
  ASSERT_TRUE(uncached.is_number());
  EXPECT_DOUBLE_EQ(uncached.as_number(), 0.0);
  EXPECT_TRUE(book.cached_cell(second, 4, 1).is_boolean());
}

// ---------------------------------------------------------------------------
// (b) Recalc reproduces what Excel computed.
// ---------------------------------------------------------------------------

TEST_P(ExternalLinkFixture, RecalcReproducesTheExcelAnswersOverTwoBooks) {
  Workbook wb = LoadFixture("external_link_mixed", GetParam());
  const Value excel_a3 = ExcelAnswer(wb, 3);
  const Value excel_a4 = ExcelAnswer(wb, 4);
  const Value excel_a5 = ExcelAnswer(wb, 5);
  ExpectSameAsExcel(OurAnswer(wb, 3), excel_a3, "[1]Data!A1");
  ExpectSameAsExcel(OurAnswer(wb, 4), excel_a4, "[1]!SrcTotal");
  ExpectSameAsExcel(OurAnswer(wb, 5), excel_a5, "[2]!FarCell");
}

TEST_P(ExternalLinkFixture, RecalcReproducesEveryCachedCellKind) {
  Workbook wb = LoadFixture("external_link_cell_kinds", GetParam());
  const char* labels[] = {"[1]!OnSecond", "SUM([1]!SecondRange)", "[1]Second!B5",
                          "[1]Second!B6", "[1]Second!B4",         "[1]Second!B7"};
  Value expected[6] = {Value::blank(), Value::blank(), Value::blank(), Value::blank(), Value::blank(), Value::blank()};
  // Text payloads point into the workbook's own storage, which the
  // recalc below rewrites, so every cached value is captured first.
  std::string cached_text[6];
  for (std::uint32_t row = 1; row <= 6U; ++row) {
    const Value excel = ExcelAnswer(wb, row);
    if (excel.is_text()) {
      cached_text[row - 1U] = std::string(excel.as_text());
      expected[row - 1U] = Value::text(cached_text[row - 1U]);
    } else {
      expected[row - 1U] = excel;
    }
  }
  for (std::uint32_t row = 1; row <= 6U; ++row) {
    ExpectSameAsExcel(OurAnswer(wb, row), expected[row - 1U], labels[row - 1U]);
  }
}

// ---------------------------------------------------------------------------
// (c) The internal references in the same workbook are unaffected.
// ---------------------------------------------------------------------------

TEST_P(ExternalLinkFixture, InternalReferencesInTheSameWorkbookStillResolve) {
  // The supporting-book table is what keeps these apart from the
  // external ones. Binding an external sheet index to a local sheet
  // would show up here as a changed value, not as an error.
  Workbook wb = LoadFixture("external_link_mixed", GetParam());
  const Value excel_a1 = ExcelAnswer(wb, 1);
  const Value excel_a2 = ExcelAnswer(wb, 2);
  ExpectSameAsExcel(OurAnswer(wb, 1), excel_a1, "Local!A1");
  ExpectSameAsExcel(OurAnswer(wb, 2), excel_a2, "SUM(Local!A1:A3)");
}

// ---------------------------------------------------------------------------
// (d) Loaded `[N]` formulas read back in the formula bar's spelling.
// ---------------------------------------------------------------------------

std::string FormulaAt(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row - 1U, 0U);
  return cell == nullptr ? std::string() : cell->formula_text;
}

TEST_P(ExternalLinkFixture, LoadedFormulasSpellTheBookWithItsAbsolutePath) {
  Workbook wb = LoadFixture("external_link_mixed", GetParam());
  EXPECT_EQ(FormulaAt(wb, 1), "=Local!A1");
  EXPECT_EQ(FormulaAt(wb, 3), "='/Users/libraz/Documents/ext_link_probe/[ExtSource.xlsx]Data'!A1");
  // A legacy formula shows the `@` Excel 365 implies on a name.
  EXPECT_EQ(FormulaAt(wb, 4), "=@'/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx'!SrcTotal");
  EXPECT_EQ(FormulaAt(wb, 5), "=@'/Users/libraz/Documents/ext_link_probe/ExtSource2.xlsx'!FarCell");
  // No stored `[N]` link index survives into the model text.
  for (std::uint32_t row = 1; row <= 5U; ++row) {
    EXPECT_EQ(FormulaAt(wb, row).find("[1]"), std::string::npos) << row;
    EXPECT_EQ(FormulaAt(wb, row).find("[2]"), std::string::npos) << row;
  }
  // Reading the spelling back binds to the loaded links, not new ones.
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[0].target, "ExtSource.xlsx");
  EXPECT_EQ(wb.external_links()[0].absolute_target, "/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx");
  EXPECT_EQ(wb.external_links()[1].absolute_target, "/Users/libraz/Documents/ext_link_probe/ExtSource2.xlsx");
}

TEST_P(ExternalLinkFixture, LoadedFormulasOverTwoSheetsSpellEveryShape) {
  Workbook wb = LoadFixture("external_link_cell_kinds", GetParam());
  const std::string dir = "/Users/libraz/Documents/ext_link_probe/";
  EXPECT_EQ(FormulaAt(wb, 1), "=@'" + dir + "ExtSource3.xlsx'!OnSecond");
  EXPECT_EQ(FormulaAt(wb, 2), "=SUM('" + dir + "ExtSource3.xlsx'!SecondRange)");
  EXPECT_EQ(FormulaAt(wb, 3), "='" + dir + "[ExtSource3.xlsx]Second'!B5");
  EXPECT_EQ(FormulaAt(wb, 6), "='" + dir + "[ExtSource3.xlsx]Second'!B7");
  EXPECT_EQ(wb.external_links().size(), 1U);
}

// ---------------------------------------------------------------------------
// (e) Variants of the xlsx fixture, edited part by part.
// ---------------------------------------------------------------------------

using PartEdit = std::function<std::string(std::string)>;

/// The fixture `<stem>.xlsx` with each named entry rewritten by its edit.
std::vector<std::uint8_t> EditedFixture(const std::string& stem, const std::map<std::string, PartEdit>& edits) {
  const std::vector<std::uint8_t> src = ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + stem + ".xlsx");
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);
  for (mz_uint i = 0; i < count; ++i) {
    char name[256] = {};
    mz_zip_reader_get_filename(&reader, i, name, sizeof(name));
    std::size_t size = 0;
    void* data = mz_zip_reader_extract_to_heap(&reader, i, &size, 0);
    EXPECT_NE(data, nullptr) << name;
    std::string bytes(static_cast<const char*>(data), size);
    mz_free(data);
    if (const auto it = edits.find(name); it != edits.end()) {
      const std::string edited = it->second(bytes);
      EXPECT_NE(edited, bytes) << "edit did not apply to " << name;
      bytes = edited;
    }
    EXPECT_NE(
        mz_zip_writer_add_mem(&writer, name, bytes.data(), bytes.size(), static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
        MZ_FALSE);
  }
  mz_zip_reader_end(&reader);
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_FALSE);
  mz_zip_writer_end(&writer);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return out;
}

PartEdit Replace(std::string from, std::string to) {
  return [from = std::move(from), to = std::move(to)](std::string text) {
    const std::size_t at = text.find(from);
    if (at != std::string::npos) {
      text.replace(at, from.size(), to);
    }
    return text;
  };
}

Workbook LoadBytes(const std::vector<std::uint8_t>& bytes) {
  auto result_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

constexpr const char* kSheet1 = "xl/worksheets/sheet1.xml";
constexpr const char* kLink1 = "xl/externalLinks/externalLink1.xml";
constexpr const char* kLink2 = "xl/externalLinks/externalLink2.xml";
constexpr const char* kAlternateUrls = "<xxl21:alternateUrls><xxl21:absoluteUrl r:id=\"rId2\"/></xxl21:alternateUrls>";

/// Appends `cells` (a run of `<c>` elements) as a new row 6 of `Use`.
PartEdit AddRow6(const std::string& cells) {
  return Replace("</sheetData>", "<row r=\"6\">" + cells + "</row></sheetData>");
}

TEST(ExternalLinkFixtureEdited, WithoutAnAbsoluteUrlTheBookIsSpelledBare) {
  Workbook wb = LoadBytes(EditedFixture("external_link_mixed", {{kLink1, Replace(kAlternateUrls, "")}}));
  EXPECT_EQ(FormulaAt(wb, 3), "=[ExtSource.xlsx]Data!A1");
  EXPECT_EQ(FormulaAt(wb, 4), "=@ExtSource.xlsx!SrcTotal");
  EXPECT_TRUE(wb.external_links()[0].absolute_target.empty());
  ExpectSameAsExcel(OurAnswer(wb, 3), Value::number(10.0), "[ExtSource.xlsx]Data!A1");
  ExpectSameAsExcel(OurAnswer(wb, 4), Value::number(30.0), "ExtSource.xlsx!SrcTotal");
}

TEST(ExternalLinkFixtureEdited, AnIndexWithNoLinkCreatesARecordThatSurvivesTheLoad) {
  Workbook wb = LoadBytes(
      EditedFixture("external_link_mixed", {{kSheet1, AddRow6("<c r=\"A6\"><f>[3]Data!A1</f><v>0</v></c>")}}));
  ASSERT_EQ(wb.external_links().size(), 3U);
  EXPECT_EQ(wb.external_links()[2].index, 3U);
  EXPECT_EQ(wb.external_links()[2].target, "3");
  EXPECT_EQ(FormulaAt(wb, 6), "=[3]Data!A1");
  ExpectSameAsExcel(OurAnswer(wb, 6), Value::error(ErrorCode::Ref), "[3]Data!A1");
}

TEST(ExternalLinkFixtureEdited, ARefreshErrorSheetWithoutCellsReadsRef) {
  Workbook wb = LoadBytes(EditedFixture(
      "external_link_mixed",
      {{kLink2, Replace("<sheetData sheetId=\"0\"><row r=\"7\"><cell r=\"D7\"><v>77</v></cell></row></sheetData>",
                        "<sheetData sheetId=\"0\" refreshError=\"1\"/>")},
       {kSheet1, AddRow6("<c r=\"A6\"><f>[2]Data!A1</f></c>")}}));
  EXPECT_FALSE(wb.external_links()[1].book.sheet_has_data(0));
  EXPECT_TRUE(wb.external_links()[0].book.sheet_has_data(0));
  ExpectSameAsExcel(OurAnswer(wb, 6), Value::error(ErrorCode::Ref), "[2]Data!A1");
}

TEST(ExternalLinkFixtureEdited, ASheetIdScopesAnExternalName) {
  Workbook wb = LoadBytes(
      EditedFixture("external_link_mixed",
                    {{kLink1, Replace("</definedNames>",
                                      "<definedName name=\"OnData\" refersTo=\"='Data'!$A$1\" sheetId=\"0\"/>"
                                      "</definedNames>")},
                     {kSheet1, AddRow6("<c r=\"A6\"><f>[1]Data!OnData</f></c><c r=\"B6\"><f>[1]!OnData</f></c>")}}));
  const ExternalBookName* local = wb.external_links()[0].book.find_name("OnData", 0);
  ASSERT_NE(local, nullptr);
  EXPECT_EQ(wb.external_links()[0].book.find_name("OnData"), nullptr);
  ExpectSameAsExcel(OurAnswer(wb, 6), Value::number(10.0), "[1]Data!OnData");
  const Cell* b6 = wb.sheet(0).cell_at(5, 1);
  ASSERT_NE(b6, nullptr);
  ASSERT_TRUE(b6->cached_value.is_error());
  EXPECT_EQ(b6->cached_value.as_error(), ErrorCode::Name);
}

TEST(ExternalLinkFixtureEdited, NamesFeatureFormulasAndValidationsAreSpelledToo) {
  Workbook wb = LoadBytes(EditedFixture(
      "external_link_mixed",
      {{kSheet1, Replace("<phoneticPr fontId=\"1\"/>",
                         "<phoneticPr fontId=\"1\"/><conditionalFormatting sqref=\"B1\"><cfRule type=\"expression\" "
                         "priority=\"1\"><formula>[1]Data!A1&gt;5</formula></cfRule></conditionalFormatting>"
                         "<dataValidations count=\"1\"><dataValidation type=\"whole\" operator=\"lessThan\" "
                         "sqref=\"C1\"><formula1>[2]!FarCell</formula1></dataValidation></dataValidations>")},
       {"xl/workbook.xml",
        Replace("</externalReferences>",
                "</externalReferences><definedNames><definedName name=\"Ext\">[1]Data!$A$1</definedName>"
                "<definedName name=\"Fresh\">[4]Other!$B$2</definedName></definedNames>")}}));
  const std::string dir = "/Users/libraz/Documents/ext_link_probe/";
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  ASSERT_EQ(wb.sheet(0).conditional_formats()[0].rules.size(), 1U);
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].formula1.value_or(""),
            "'" + dir + "[ExtSource.xlsx]Data'!A1>5");
  ASSERT_EQ(wb.sheet(0).validations().size(), 1U);
  EXPECT_EQ(wb.sheet(0).validations()[0].formula1, "'" + dir + "ExtSource2.xlsx'!FarCell");
  ASSERT_EQ(wb.defined_names().size(), 2U);
  EXPECT_EQ(wb.defined_names()[0].formula, "'" + dir + "[ExtSource.xlsx]Data'!$A$1");
  EXPECT_EQ(wb.defined_names()[1].formula, "[4]Other!$B$2");
  // `[4]` had no link: the name created one.
  ASSERT_EQ(wb.external_links().size(), 3U);
  EXPECT_EQ(wb.external_links()[2].target, "4");
}

// Names each instantiation after its container extension without the dot,
// so the two cases read as `xlsx` and `xlsb`. It is a free function rather
// than a lambda spelled inside the macro argument because the macro expands
// around a parameter of its own named `info`, which a lambda parameter of
// that name shadows.
std::string ContainerSuffixName(const ::testing::TestParamInfo<const char*>& info) {
  return std::string(info.param).substr(1);
}

INSTANTIATE_TEST_SUITE_P(BothContainerFormats, ExternalLinkFixture, ::testing::Values(".xlsx", ".xlsb"),
                         ContainerSuffixName);

}  // namespace
}  // namespace formulon
