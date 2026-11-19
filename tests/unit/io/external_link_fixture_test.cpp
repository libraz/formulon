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
//       A5  =[1]Second!B4              (never cached -> blank, shown as 0)
//       A6  =[1]Second!B7              (cached text)
//
// A5 is the case worth stating twice: Excel caches only the cells the
// consuming workbook references, and reads an address it does not hold
// as numeric zero rather than as blank or `#REF!`.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <functional>
#include <map>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_book.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml/external_link_reader.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
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

TEST_P(ExternalLinkFixture, AnUncachedAddressReadsAsBlank) {
  Workbook wb = LoadFixture("external_link_cell_kinds", GetParam());
  ASSERT_EQ(wb.external_links().size(), 1U);
  const ExternalBook& book = wb.external_links()[0].book;
  const std::uint32_t second = book.sheet_index("Second");
  ASSERT_NE(second, ExternalBook::kNoSheet);
  // B4 (row index 3) was never cached; B5 was. In a closed book Excel reads
  // an uncached cell as blank (`&"x"` gives "x", ISBLANK is TRUE); only the
  // top-level display shows it as 0.
  EXPECT_TRUE(book.cached_cell(second, 3, 1).is_blank());
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

std::string ReplaceAllText(std::string text, std::string_view from, std::string_view to) {
  std::size_t at = 0U;
  while ((at = text.find(from, at)) != std::string::npos) {
    text.replace(at, from.size(), to);
    at += to.size();
  }
  return text;
}

PartEdit AddQuotedExternalNames() {
  return [](std::string xml) {
    xml = ReplaceAllText(std::move(xml), "<sheetName val=\"Data\"/>",
                         "<sheetName val=\"Data\"/><sheetName val=\"O'Brien! 東京\"/>");
    xml = ReplaceAllText(std::move(xml), "</definedNames>",
                         "<definedName name=\"QuotedCell\" refersTo=\"='O''Brien! 東京'!$A$1\"/>"
                         "<definedName name=\"QuotedRange\" refersTo=\"='O''Brien! 東京'!$A$1:$A$3\"/>"
                         "<definedName name=\"UnquotedCell\" refersTo=\"=Data!$A$1\"/></definedNames>");
    return ReplaceAllText(std::move(xml), "</sheetDataSet>",
                          "<sheetData sheetId=\"1\"><row r=\"1\"><cell r=\"A1\"><v>10</v></cell></row>"
                          "<row r=\"3\"><cell r=\"A3\"><v>30</v></cell></row></sheetData></sheetDataSet>");
  };
}

PartEdit PrefixExternalLinkRelationships() {
  return [](std::string xml) {
    constexpr std::string_view kRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    xml = ReplaceAllText(
        std::move(xml), "<Relationships xmlns=\"" + std::string(kRelationshipsNamespace) + "\">",
        "<pkg:Relationships xmlns:pkg=\"" + std::string(kRelationshipsNamespace) + "\" xmlns:foreign=\"urn:foreign\">");
    xml = ReplaceAllText(std::move(xml), "<Relationship ", "<pkg:Relationship ");
    xml = ReplaceAllText(std::move(xml), "</Relationships>", "</pkg:Relationships>");
    const std::size_t first_relationship = xml.find("<pkg:Relationship ");
    if (first_relationship != std::string::npos) {
      xml.insert(first_relationship,
                 "<foreign:Relationship Id=\"rId1\" "
                 "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" "
                 "Target=\"Wrong.xlsx\" TargetMode=\"Internal\"/>");
    }
    return xml;
  };
}

PartEdit NamespaceLessExternalLinkRelationships() {
  return [](std::string xml) {
    constexpr std::string_view kRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    xml = ReplaceAllText(std::move(xml), "<Relationships xmlns=\"" + std::string(kRelationshipsNamespace) + "\">",
                         "<Relationships xmlns:foreign=\"urn:foreign\">");
    const std::size_t first_relationship = xml.find("<Relationship ");
    if (first_relationship != std::string::npos) {
      xml.insert(first_relationship,
                 "<foreign:Relationship Id=\"rId1\" "
                 "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" "
                 "Target=\"Wrong.xlsx\" TargetMode=\"Internal\"/>");
    }
    return xml;
  };
}

PartEdit ForeignOnlyPrefixedExternalLinkRelationships() {
  return [](std::string) {
    return "<pkg:Relationships xmlns:pkg=\"http://schemas.openxmlformats.org/package/2006/relationships\" "
           "xmlns:foreign=\"urn:foreign\"><foreign:Relationship Id=\"rId1\" "
           "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath\" "
           "Target=\"Wrong.xlsx\" TargetMode=\"External\"/></pkg:Relationships>";
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
constexpr const char* kLink1Rels = "xl/externalLinks/_rels/externalLink1.xml.rels";
constexpr const char* kAlternateUrls = "<xxl21:alternateUrls><xxl21:absoluteUrl r:id=\"rId2\"/></xxl21:alternateUrls>";

/// Appends `cells` (a run of `<c>` elements) as a new row 6 of `Use`.
PartEdit AddRow6(const std::string& cells) {
  return Replace("</sheetData>", "<row r=\"6\">" + cells + "</row></sheetData>");
}

TEST(ExternalLinkFixtureEdited, QuotedExternalNamesDecodeDoubledApostrophesForLookup) {
  Workbook wb = LoadBytes(
      EditedFixture("external_link_mixed", {{kLink1, AddQuotedExternalNames()},
                                            {kSheet1, AddRow6("<c r=\"A6\"><f>[1]!QuotedCell</f><v>0</v></c>"
                                                              "<c r=\"B6\"><f>SUM([1]!QuotedRange)</f><v>0</v></c>"
                                                              "<c r=\"C6\"><f>[1]!UnquotedCell</f><v>0</v></c>")}}));
  ASSERT_EQ(wb.external_links().size(), 2U);
  const ExternalBook& book = wb.external_links()[0].book;
  ASSERT_EQ(book.sheet_index("O'Brien! 東京"), 1U);
  const ExternalBookName* single = book.find_name("QuotedCell");
  const ExternalBookName* range = book.find_name("QuotedRange");
  const ExternalBookName* unquoted = book.find_name("UnquotedCell");
  ASSERT_NE(single, nullptr);
  ASSERT_NE(range, nullptr);
  ASSERT_NE(unquoted, nullptr);
  EXPECT_TRUE(single->resolvable);
  EXPECT_EQ(single->sheet, 1U);
  EXPECT_EQ(single->row, 0U);
  EXPECT_EQ(single->col, 0U);
  EXPECT_TRUE(range->resolvable);
  EXPECT_TRUE(range->is_range);
  EXPECT_EQ(range->sheet, 1U);
  EXPECT_EQ(range->row_end, 2U);
  EXPECT_TRUE(unquoted->resolvable);
  EXPECT_EQ(unquoted->sheet, 0U);

  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << (recalc_or ? "" : recalc_or.error().message);
  const Cell* single_cell = wb.sheet(0).cell_at(5U, 0U);
  const Cell* range_cell = wb.sheet(0).cell_at(5U, 1U);
  const Cell* unquoted_cell = wb.sheet(0).cell_at(5U, 2U);
  ASSERT_NE(single_cell, nullptr);
  ASSERT_NE(range_cell, nullptr);
  ASSERT_NE(unquoted_cell, nullptr);
  ASSERT_TRUE(single_cell->cached_value.is_number());
  ASSERT_TRUE(range_cell->cached_value.is_number());
  ASSERT_TRUE(unquoted_cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(single_cell->cached_value.as_number(), 10.0);
  EXPECT_DOUBLE_EQ(range_cell->cached_value.as_number(), 40.0);
  EXPECT_DOUBLE_EQ(unquoted_cell->cached_value.as_number(), 10.0);

  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  Workbook roundtrip = LoadBytes(saved.value());
  ASSERT_FALSE(roundtrip.external_links().empty());
  const ExternalBookName* saved_single = roundtrip.external_links()[0].book.find_name("QuotedCell");
  ASSERT_NE(saved_single, nullptr);
  EXPECT_TRUE(saved_single->resolvable);
  ExpectSameAsExcel(OurAnswer(roundtrip, 6U), Value::number(10.0), "[1]!QuotedCell after save/reload");
}

TEST(ExternalLinkFixtureEdited, PrefixedPerLinkRelationshipsPreserveTargetsThroughSaveReload) {
  Workbook baseline = LoadFixture("external_link_mixed", ".xlsx");
  ASSERT_FALSE(baseline.external_links().empty());
  const ExternalLinkRecord& default_record = baseline.external_links()[0];
  ASSERT_EQ(default_record.target, "ExtSource.xlsx");
  ASSERT_TRUE(default_record.target_external);
  ASSERT_EQ(default_record.body_rel_id, "rId1");
  ASSERT_EQ(default_record.absolute_target, "/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx");

  Workbook wb = LoadBytes(EditedFixture("external_link_mixed", {{kLink1Rels, PrefixExternalLinkRelationships()}}));
  ASSERT_FALSE(wb.external_links().empty());
  const ExternalLinkRecord& prefixed_record = wb.external_links()[0];
  EXPECT_EQ(prefixed_record.target, default_record.target);
  EXPECT_EQ(prefixed_record.target_external, default_record.target_external);
  EXPECT_EQ(prefixed_record.body_rel_id, default_record.body_rel_id);
  EXPECT_EQ(prefixed_record.absolute_target, default_record.absolute_target);
  ExpectSameAsExcel(OurAnswer(wb, 3U), Value::number(10.0), "[1]Data!A1");

  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  Workbook roundtrip = LoadBytes(saved.value());
  ASSERT_FALSE(roundtrip.external_links().empty());
  const ExternalLinkRecord& reloaded = roundtrip.external_links()[0];
  EXPECT_EQ(reloaded.target, default_record.target);
  EXPECT_EQ(reloaded.target_external, default_record.target_external);
  EXPECT_EQ(reloaded.body_rel_id, default_record.body_rel_id);
  EXPECT_EQ(reloaded.absolute_target, default_record.absolute_target);
  ExpectSameAsExcel(OurAnswer(roundtrip, 3U), Value::number(10.0), "[1]Data!A1 after save/reload");
}

TEST(ExternalLinkFixtureEdited, NamespaceLessRelationshipPartsRemainReadable) {
  Workbook wb =
      LoadBytes(EditedFixture("external_link_mixed", {{kLink1Rels, NamespaceLessExternalLinkRelationships()}}));
  ASSERT_FALSE(wb.external_links().empty());
  const ExternalLinkRecord& record = wb.external_links()[0];
  EXPECT_EQ(record.target, "ExtSource.xlsx");
  EXPECT_TRUE(record.target_external);
  EXPECT_EQ(record.body_rel_id, "rId1");
  EXPECT_EQ(record.absolute_target, "/Users/libraz/Documents/ext_link_probe/ExtSource.xlsx");
}

TEST(ExternalLinkFixtureEdited, ForeignRelationshipChildrenAreIgnored) {
  const std::vector<std::uint8_t> bytes =
      EditedFixture("external_link_mixed", {{kLink1Rels, ForeignOnlyPrefixedExternalLinkRelationships()}});
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{bytes.data(), bytes.size()})));

  pugi::xml_document workbook_doc;
  ASSERT_TRUE(workbook_doc.load_string(
      "<workbook xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<externalReferences><externalReference r:id=\"rId1\"/></externalReferences></workbook>"));
  io::ooxml::WorkbookRels workbook_rels;
  workbook_rels.external_link_paths_by_rid.emplace("rId1", "xl/externalLinks/externalLink1.xml");
  auto loaded = io::ooxml::load_external_links(zip, workbook_doc.child("workbook"), workbook_rels);
  ASSERT_TRUE(static_cast<bool>(loaded));
  ASSERT_EQ(loaded.value().records.size(), 1U);
  const ExternalLinkRecord& record = loaded.value().records[0];
  EXPECT_EQ(record.body_rel_id, "rId1");
  EXPECT_TRUE(record.target.empty());
  EXPECT_TRUE(record.absolute_target.empty());
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

// A failed-refresh sheet that still caches some cells (Mac Excel 16.115): the mark is kept on load, a single uncached
// cell reads #REF!, a cached one its value, a range skips the uncached cells, and a save writes the mark back.
TEST(ExternalLinkFixtureEdited, ARefreshErrorSheetWithCellsKeepsTheMark) {
  Workbook wb = LoadBytes(EditedFixture(
      "external_link_mixed", {{kLink2, Replace("<sheetData sheetId=\"0\"><row r=\"7\">",
                                               "<sheetData sheetId=\"0\" refreshError=\"1\"><row r=\"7\">")},
                              {kSheet1, AddRow6("<c r=\"A6\"><f>[2]Data!D8</f></c><c r=\"B6\"><f>[2]Data!D7</f></c>"
                                                "<c r=\"C6\"><f>SUM([2]Data!D7:D8)</f></c>")}}));
  const ExternalBook& book = wb.external_links()[1].book;
  EXPECT_TRUE(book.sheet_has_data(0));
  EXPECT_TRUE(book.sheet_refresh_failed(0));
  ExpectSameAsExcel(OurAnswer(wb, 6), Value::error(ErrorCode::Ref), "[2]Data!D8");
  const Cell* b6 = wb.sheet(0).cell_at(5, 1);
  ASSERT_NE(b6, nullptr);
  ASSERT_TRUE(b6->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b6->cached_value.as_number(), 77.0);
  const Cell* c6 = wb.sheet(0).cell_at(5, 2);
  ASSERT_NE(c6, nullptr);
  ASSERT_TRUE(c6->cached_value.is_number());
  EXPECT_DOUBLE_EQ(c6->cached_value.as_number(), 77.0);

  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  Workbook reloaded = LoadBytes(saved.value());
  EXPECT_TRUE(reloaded.external_links()[1].book.sheet_refresh_failed(0));
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(io::ByteSpan{xlsb.value().data(), xlsb.value().size()});
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  EXPECT_TRUE(from_xlsb.value().workbook.external_links()[1].book.sheet_refresh_failed(0));
}

/// The named part of a zip package, empty when absent.
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

/// The payload of every `BrtBeginExternTable` (363) in an xlsb external link part, in order.
std::vector<std::string> ExternTablePayloads(const std::string& part) {
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
    if (type == 363U) {
      out.push_back(part.substr(at, size));
    }
    at += size;
  }
  return out;
}

// Excel's own saves of one book whose link marks only its second sheet as a failed refresh: the xlsx spells it
// `<sheetData sheetId="1" refreshError="1"/>`, the xlsb sets bit 0 of the byte after the table's sheetId.
TEST(ExternalLinkRefreshError, ExcelTwinMarksTheSameSheetInBothFormats) {
  for (const char* extension : {".xlsx", ".xlsb"}) {
    Workbook wb = LoadFixture("external_table_ref_closed", extension);
    ASSERT_EQ(wb.external_links().size(), 1U) << extension;
    const ExternalBook& book = wb.external_links()[0].book;
    EXPECT_FALSE(book.sheet_refresh_failed(0)) << extension;
    EXPECT_TRUE(book.sheet_refresh_failed(1)) << extension;
  }
}

TEST(ExternalLinkRefreshError, SavesWriteTheMarkAsExcelDoes) {
  const std::vector<std::uint8_t> excel_xlsb =
      ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/external_table_ref_closed.xlsb");
  const std::vector<std::string> excel_tables =
      ExternTablePayloads(ZipPart(excel_xlsb, "xl/externalLinks/externalLink1.bin"));
  ASSERT_EQ(excel_tables.size(), 2U);
  EXPECT_EQ(excel_tables[1], std::string("\x01\x00\x00\x00\x01", 5));

  Workbook wb = LoadFixture("external_table_ref_closed", ".xlsx");
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  EXPECT_EQ(ExternTablePayloads(ZipPart(xlsb.value(), "xl/externalLinks/externalLink1.bin")), excel_tables);

  auto xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  EXPECT_NE(ZipPart(xlsx.value(), "xl/externalLinks/externalLink1.xml")
                .find("<sheetData sheetId=\"0\"/><sheetData sheetId=\"1\" refreshError=\"1\"/>"),
            std::string::npos);
}

// A marked sheet with no cached cells survives xlsx -> xlsb -> xlsx and still reads #REF! for every address.
TEST(ExternalLinkRefreshError, AnEmptyMarkedSheetSurvivesBothFormats) {
  Workbook wb = LoadFixture("external_table_ref_closed", ".xlsx");
  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(io::ByteSpan{xlsb.value().data(), xlsb.value().size()});
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  auto xlsx = io::write_ooxml(from_xlsb.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  Workbook reloaded = LoadBytes(xlsx.value());
  const ExternalBook& book = reloaded.external_links()[0].book;
  EXPECT_FALSE(book.sheet_refresh_failed(0));
  EXPECT_TRUE(book.sheet_refresh_failed(1));
  EXPECT_TRUE(book.referenced_cell(1, 0, 0).is_error());
  EXPECT_EQ(book.referenced_cell(1, 0, 0).as_error(), ErrorCode::Ref);
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

TEST(ExternalLinkFixtureEdited, AbsentNameIsRecordedInPlaceAndReadsAsNoSuchName) {
  Workbook wb = LoadBytes(EditedFixture(
      "external_link_mixed", {{kLink1, Replace("</definedNames>",
                                               "<definedName name=\"NoSuch\"/>"
                                               "<definedName name=\"Dead\" refersTo=\"#REF!\"/>"
                                               "<definedName name=\"AfterGap\" refersTo=\"='Data'!$A$1\"/>"
                                               "</definedNames>")},
                              {kSheet1, AddRow6("<c r=\"A6\"><f>[1]!NoSuch</f></c><c r=\"B6\"><f>[1]!Dead</f></c>"
                                                "<c r=\"C6\"><f>[1]!AfterGap</f></c>")}}));
  const ExternalBook& book = wb.external_links()[0].book;
  const ExternalBookName* absent = book.find_name("NoSuch");
  const ExternalBookName* dead = book.find_name("Dead");
  const ExternalBookName* after = book.find_name("AfterGap");
  ASSERT_NE(absent, nullptr);
  ASSERT_NE(dead, nullptr);
  ASSERT_NE(after, nullptr);
  EXPECT_FALSE(absent->exists);
  EXPECT_TRUE(dead->exists);
  EXPECT_FALSE(dead->resolvable);
  EXPECT_TRUE(after->exists);
  EXPECT_TRUE(after->resolvable);
  ExpectSameAsExcel(OurAnswer(wb, 6), Value::error(ErrorCode::Name), "[1]!NoSuch");
  const Cell* b6 = wb.sheet(0).cell_at(5, 1);
  const Cell* c6 = wb.sheet(0).cell_at(5, 2);
  ASSERT_NE(b6, nullptr);
  ASSERT_NE(c6, nullptr);
  ASSERT_TRUE(b6->cached_value.is_error());
  EXPECT_EQ(b6->cached_value.as_error(), ErrorCode::Ref);
  ASSERT_TRUE(c6->cached_value.is_number());
  EXPECT_DOUBLE_EQ(c6->cached_value.as_number(), 10.0);
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

// ---------------------------------------------------------------------------
// (f) External links are numbered by first use at load.
// ---------------------------------------------------------------------------

std::string BaseName(const std::string& path) {
  const std::size_t slash = path.find_last_of('/');
  return slash == std::string::npos ? path : path.substr(slash + 1U);
}

/// The base names of the links' targets in index order.
std::vector<std::string> LinkTargets(const Workbook& wb) {
  std::vector<std::string> out;
  for (const ExternalLinkRecord* rec : wb.external_links_by_index(/*skip_ole_dde=*/false)) {
    out.push_back(BaseName(rec->target.empty() ? rec->absolute_target : rec->target));
  }
  return out;
}

/// The base names of the links' targets in the order the file stored them: a loaded link keeps its part path, and
/// a part's number is its stored position.
std::vector<std::string> StoredTargets(const Workbook& wb) {
  std::vector<std::pair<int, std::string>> numbered;
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    const std::size_t digits = rec.part_path.find_first_of("0123456789");
    EXPECT_NE(digits, std::string::npos) << rec.part_path;
    numbered.emplace_back(digits == std::string::npos ? 0 : std::atoi(rec.part_path.c_str() + digits),
                          BaseName(rec.target.empty() ? rec.absolute_target : rec.target));
  }
  std::sort(numbered.begin(), numbered.end());
  std::vector<std::string> out;
  for (const auto& entry : numbered) {
    out.push_back(entry.second);
  }
  return out;
}

/// Every cell formula as `sheet:row:col=text`, sorted.
std::vector<std::string> CellFormulas(const Workbook& wb) {
  std::vector<std::string> out;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    for (const auto& row : wb.sheet(i).rows()) {
      for (const RowCells::Entry entry : row.second) {
        if (!entry.cell.formula_text.empty()) {
          out.push_back(std::to_string(i) + ":" + std::to_string(row.first) + ":" + std::to_string(entry.col) + "=" +
                        entry.cell.formula_text);
        }
      }
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

Workbook LoadXlsbBytes(const std::vector<std::uint8_t>& bytes) {
  auto result_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

// In the external_link_renumber fixtures each link's absolute path is normalised to a neutral home directory;
// nothing else in the packages is edited. The xlsb_third_active pair is Excel's packages unedited.
struct RenumberRow {
  const char* row;
  const char* authored_extension;
  /// True when Excel's save drops a link nothing uses, so its book has fewer links than the authored one.
  bool drops_unused;
};

constexpr RenumberRow kRenumberRows[] = {
    {"in_use_order", ".xlsx", false},          {"out_of_use_order", ".xlsx", false},
    {"xlsb_in_use_order", ".xlsb", false},     {"unused_link", ".xlsx", true},
    {"sheets_in_tab_order", ".xlsx", false},   {"names_in_stored_order", ".xlsx", false},
    {"sheet_name_first", ".xlsx", false},      {"book_name_first", ".xlsx", false},
    {"external_defined_name", ".xlsx", false},
};

Workbook LoadRenumberAuthored(const RenumberRow& row) {
  return LoadFixture(std::string("external_link_renumber/") + row.row + "_authored", row.authored_extension);
}

Workbook LoadRenumberSaved(const RenumberRow& row) {
  return LoadFixture(std::string("external_link_renumber/") + row.row + "_saved", ".xlsx");
}

/// The links of `ours` in index order against Excel's `saved`: equal, or `saved` is the used prefix when Excel's
/// save drops unused links.
void ExpectLinkOrder(const RenumberRow& row, const std::vector<std::string>& ours,
                     const std::vector<std::string>& saved) {
  if (row.drops_unused) {
    ASSERT_LE(saved.size(), ours.size());
    EXPECT_TRUE(std::equal(saved.begin(), saved.end(), ours.begin()));
  } else {
    EXPECT_EQ(ours, saved);
  }
}

// Excel's save writes the order its load gave, and its own saves are renumbered again on the next load, so the
// comparison is against the order the saved file stores.
TEST(ExternalLinkRenumber, LoadNumbersLinksAsExcelDoes) {
  for (const RenumberRow& row : kRenumberRows) {
    SCOPED_TRACE(row.row);
    const Workbook authored = LoadRenumberAuthored(row);
    const Workbook saved = LoadRenumberSaved(row);
    ExpectLinkOrder(row, LinkTargets(authored), StoredTargets(saved));
  }
}

TEST(ExternalLinkRenumber, XlsxSaveMatchesExcelsSave) {
  for (const RenumberRow& row : kRenumberRows) {
    SCOPED_TRACE(row.row);
    const Workbook saved = LoadRenumberSaved(row);
    Workbook authored = LoadRenumberAuthored(row);
    auto bytes = io::write_ooxml(authored);
    ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
    const Workbook reloaded = LoadBytes(bytes.value());
    EXPECT_EQ(StoredTargets(reloaded), StoredTargets(saved));
    EXPECT_EQ(CellFormulas(reloaded), CellFormulas(saved));
  }
}

TEST(ExternalLinkRenumber, XlsbSaveMatchesExcelsSave) {
  for (const RenumberRow& row : kRenumberRows) {
    SCOPED_TRACE(row.row);
    const Workbook saved = LoadRenumberSaved(row);
    Workbook authored = LoadRenumberAuthored(row);
    auto bytes = io::xlsb::write_xlsb(authored);
    ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
    const Workbook reloaded = LoadXlsbBytes(bytes.value());
    EXPECT_EQ(StoredTargets(reloaded), StoredTargets(saved));
    EXPECT_EQ(CellFormulas(reloaded), CellFormulas(saved));
  }
}

struct ReopenRow {
  /// Fixture stem, under tests/fixtures/excel, of the book Excel opened.
  const char* input;
  /// Fixture stem of Excel's save of it, opened with no edit.
  const char* reopened;
};

constexpr ReopenRow kReopenRows[] = {
    {"extlink_gap_nameonly_excel", "external_link_renumber/name_link_reopened"},
    {"external_link_renumber/names_in_stored_order_saved", "external_link_renumber/names_in_stored_order_reopened"},
    {"external_link_renumber/sheet_name_first_saved", "external_link_renumber/sheet_name_first_reopened"},
    {"external_link_renumber/second_sheet_active_input", "external_link_renumber/second_sheet_active_reopened"},
    {"external_link_renumber/name_link_before_cells_input", "external_link_renumber/name_link_before_cells_reopened"},
    {"external_link_renumber/no_calc_chain_input", "external_link_renumber/no_calc_chain_reopened"},
    {"external_link_renumber/relative_cell_refs_input", "external_link_renumber/relative_cell_refs_reopened"},
    {"external_link_renumber/first_sheet_excel_save_input", "external_link_renumber/first_sheet_excel_save_reopened"},
    {"external_link_renumber/bare_link_targets_input", "external_link_renumber/bare_link_targets_reopened"},
    {"external_link_renumber/first_sheet_active_input", "external_link_renumber/first_sheet_active_reopened"},
    {"external_link_renumber/second_of_three_active_input", "external_link_renumber/second_of_three_active_reopened"},
    {"external_link_renumber/third_of_three_active_input", "external_link_renumber/third_of_three_active_reopened"},
    {"external_link_renumber/middle_active_rest_in_tab_order_input",
     "external_link_renumber/middle_active_rest_in_tab_order_reopened"},
    {"external_link_renumber/last_active_input", "external_link_renumber/last_active_reopened"},
    {"external_link_renumber/no_active_tab_input", "external_link_renumber/no_active_tab_reopened"},
    {"external_link_renumber/excel_session_book_input", "external_link_renumber/excel_session_book_reopened"},
};

// Names first, then cells from the active sheet (activeTab, else the first sheet; tabSelected is not read), then the
// other sheets in tab order. Several inputs are Excel's own saves, which are not fixed points of its load.
TEST(ExternalLinkRenumber, AReopenNumbersLinksAsExcelDoes) {
  for (const ReopenRow& row : kReopenRows) {
    SCOPED_TRACE(row.input);
    const Workbook input = LoadFixture(row.input, ".xlsx");
    const Workbook reopened = LoadFixture(row.reopened, ".xlsx");
    ASSERT_FALSE(input.external_links().empty());
    EXPECT_EQ(LinkTargets(input), StoredTargets(reopened));
  }
}

// An xlsb whose active tab is the third sheet, with its links stored S2, S1, S3 and the third sheet naming S3:
// Excel's reopen saves them in the stored order, which an xlsx load would renumber from the active sheet.
TEST(ExternalLinkRenumber, AnXlsbReopenKeepsTheStoredOrder) {
  const Workbook input = LoadFixture("external_link_renumber/xlsb_third_active_input", ".xlsb");
  const Workbook reopened = LoadFixture("external_link_renumber/xlsb_third_active_reopened", ".xlsx");
  ASSERT_EQ(input.active_sheet(), 2U);
  EXPECT_EQ(LinkTargets(input), StoredTargets(reopened));
}

/// `formulas` with each quoted book path cut to its file name: the fixtures' link directories differ.
std::vector<std::string> WithoutBookDirectories(std::vector<std::string> formulas) {
  for (std::string& text : formulas) {
    for (std::size_t open = text.find('\''); open != std::string::npos; open = text.find('\'', open + 1U)) {
      const std::size_t close = text.find('\'', open + 1U);
      if (close == std::string::npos) {
        break;
      }
      const std::size_t slash = text.rfind('/', close);
      if (slash != std::string::npos && slash > open) {
        text.erase(open + 1U, slash - open);
      }
      open = text.find('\'', open + 1U);
      if (open == std::string::npos) {
        break;
      }
    }
  }
  return formulas;
}

TEST(ExternalLinkRenumber, ASaveOfAReopenStoresExcelsOrder) {
  for (const ReopenRow& row : kReopenRows) {
    SCOPED_TRACE(row.input);
    Workbook input = LoadFixture(row.input, ".xlsx");
    const Workbook reopened = LoadFixture(row.reopened, ".xlsx");
    auto bytes = io::write_ooxml(input);
    ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
    const Workbook reloaded = LoadBytes(bytes.value());
    EXPECT_EQ(StoredTargets(reloaded), StoredTargets(reopened));
    EXPECT_EQ(WithoutBookDirectories(CellFormulas(reloaded)), WithoutBookDirectories(CellFormulas(reopened)));
  }
}

TEST(ExternalLinkRenumber, LinksAlreadyInFirstUseOrderAreUntouched) {
  Workbook wb = LoadFixture("external_link_renumber/in_use_order_authored", ".xlsx");
  ASSERT_EQ(wb.external_links().size(), 2U);
  for (std::uint32_t i = 0; i < 2U; ++i) {
    const ExternalLinkRecord& rec = wb.external_links()[i];
    EXPECT_EQ(rec.index, i + 1U);
    EXPECT_EQ(rec.part_path, "xl/externalLinks/externalLink" + std::to_string(i + 1U) + ".xml");
  }
  EXPECT_EQ(LinkTargets(wb), (std::vector<std::string>{"S1.xlsx", "S2.xlsx"}));
}

TEST(ExternalLinkRenumber, LoadedPartsKeepTheirPathsUntilTheNextSave) {
  Workbook wb = LoadFixture("external_link_renumber/out_of_use_order_authored", ".xlsx");
  ASSERT_EQ(LinkTargets(wb), (std::vector<std::string>{"S3.xlsx", "S1.xlsx", "S2.xlsx"}));
  // The first link is the part the file stored as the third.
  EXPECT_EQ(wb.external_links()[0].index, 1U);
  EXPECT_EQ(wb.external_links()[0].part_path, "xl/externalLinks/externalLink3.xml");
}

TEST(ExternalLinkRenumber, SaveWritesEachLinkAtTheDensePathOfItsNumber) {
  Workbook wb = LoadFixture("external_link_renumber/out_of_use_order_authored", ".xlsx");
  auto bytes = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  const char* rels[] = {"xl/externalLinks/_rels/externalLink1.xml.rels",
                        "xl/externalLinks/_rels/externalLink2.xml.rels",
                        "xl/externalLinks/_rels/externalLink3.xml.rels"};
  const char* targets[] = {"S3.xlsx", "S1.xlsx", "S2.xlsx"};
  for (std::size_t i = 0; i < 3U; ++i) {
    EXPECT_NE(ZipPart(bytes.value(), rels[i]).find(targets[i]), std::string::npos) << rels[i];
  }
  const Workbook reloaded = LoadBytes(bytes.value());
  ASSERT_EQ(reloaded.external_links().size(), 3U);
  for (const ExternalLinkRecord& rec : reloaded.external_links()) {
    EXPECT_EQ(rec.part_path, "xl/externalLinks/externalLink" + std::to_string(rec.index) + ".xml") << rec.index;
  }
}

TEST(ExternalLinkRenumber, ALinkNamedByAConditionalFormatKeepsItsNumber) {
  // Link 1 is named by a rule formula, so it stays while the later links take first-use order among themselves.
  Workbook wb = LoadBytes(EditedFixture(
      "external_link_renumber/out_of_use_order_authored",
      {{kSheet1, Replace("<phoneticPr fontId=\"1\"/>",
                         "<phoneticPr fontId=\"1\"/><conditionalFormatting sqref=\"B1\"><cfRule type=\"expression\" "
                         "priority=\"1\"><formula>[1]A!A1&gt;5</formula></cfRule></conditionalFormatting>")}}));
  EXPECT_EQ(LinkTargets(wb), (std::vector<std::string>{"S1.xlsx", "S3.xlsx", "S2.xlsx"}));
}

TEST(ExternalLinkRenumber, ALinkWithNoTargetStaysPinned) {
  // Link 1 has no target and no formula; links 2 and 3 are numbered by first use above it.
  Workbook wb = LoadBytes(EditedFixture(
      "external_link_renumber/out_of_use_order_authored",
      {{kLink1Rels,
        [](std::string) {
          return "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>";
        }},
       {kSheet1, Replace("<f>[1]A!A1</f>", "")}}));
  ASSERT_EQ(wb.external_links().size(), 3U);
  // A link with no target carries its number as its target.
  EXPECT_EQ(wb.external_links_by_index(false)[0]->target, "1");
  EXPECT_EQ(LinkTargets(wb), (std::vector<std::string>{"1", "S3.xlsx", "S2.xlsx"}));
}

// Excel's saves of books whose links name a source it could not locate (a sheet qualifier naming no sheet of the
// workbook): the link's target relationship is the `xlPathMissing` kind with the bare qualifier as its target.
//   one_link     Sheet1!A1 `[1]Nope!A1`
//   two_links    Sheet1!A1 `[1]Nope!A1`, A2 `[2]Other!B3`
//   book_name    Sheet1!A1 `[1]!LocalName`; the link lists the name with no target
constexpr const char* kPathMissingStems[] = {"one_link", "two_links", "book_name"};
constexpr std::string_view kPathMissingType =
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing";

std::string PathMissingStem(const char* stem) {
  return std::string("external_link_path_missing/") + stem;
}

std::vector<std::uint8_t> PathMissingBytes(const char* stem, const char* extension) {
  return ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + PathMissingStem(stem) + extension);
}

/// Every relationship of a rels part as `Type|Target|TargetMode`, sorted.
std::vector<std::string> Relationships(const std::string& rels) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(rels.c_str())) << rels;
  std::vector<std::string> out;
  for (const pugi::xml_node rel : doc.document_element().children("Relationship")) {
    out.push_back(std::string(rel.attribute("Type").value()) + "|" + rel.attribute("Target").value() + "|" +
                  rel.attribute("TargetMode").value());
  }
  std::sort(out.begin(), out.end());
  return out;
}

std::string LinkRelsPart(std::size_t number, const char* extension) {
  return "xl/externalLinks/_rels/externalLink" + std::to_string(number) +
         (std::string_view(extension) == ".xlsb" ? ".bin" : ".xml") + ".rels";
}

TEST(ExternalLinkPathMissing, LoadKeepsTheBareTarget) {
  for (const char* extension : {".xlsx", ".xlsb"}) {
    for (const char* stem : kPathMissingStems) {
      SCOPED_TRACE(std::string(stem) + extension);
      const Workbook wb = LoadFixture(PathMissingStem(stem), extension);
      const std::vector<std::string> expected = std::string_view(stem) == "two_links"
                                                    ? std::vector<std::string>{"Nope", "Other"}
                                                    : std::vector<std::string>{"Nope"};
      EXPECT_EQ(LinkTargets(wb), expected);
      for (const ExternalLinkRecord& rec : wb.external_links()) {
        EXPECT_TRUE(rec.target_path_missing);
        EXPECT_TRUE(rec.target_external);
        EXPECT_TRUE(rec.absolute_target.empty());
        ASSERT_EQ(rec.book.sheet_names.size(), 1U);
        EXPECT_EQ(rec.book.sheet_names[0], rec.target);
        EXPECT_TRUE(rec.book.sheet_refresh_failed(0));
      }
      if (std::string_view(stem) == "book_name") {
        EXPECT_NE(wb.external_links()[0].book.find_name("LocalName"), nullptr);
      }
    }
  }
}

TEST(ExternalLinkPathMissing, SavesWriteTheRelationshipAsExcelDoes) {
  for (const char* source : {".xlsx", ".xlsb"}) {
    for (const char* stem : kPathMissingStems) {
      Workbook wb = LoadFixture(PathMissingStem(stem), source);
      for (const char* extension : {".xlsx", ".xlsb"}) {
        SCOPED_TRACE(std::string(stem) + source + " -> " + extension);
        auto saved = std::string_view(extension) == ".xlsb" ? io::xlsb::write_xlsb(wb) : io::write_ooxml(wb);
        ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
        const std::vector<std::uint8_t> excel = PathMissingBytes(stem, extension);
        for (std::size_t number = 1; number <= wb.external_links().size(); ++number) {
          const std::vector<std::string> ours =
              Relationships(ZipPart(saved.value(), LinkRelsPart(number, extension).c_str()));
          EXPECT_EQ(ours, Relationships(ZipPart(excel, LinkRelsPart(number, extension).c_str())));
          ASSERT_EQ(ours.size(), 1U);
          EXPECT_EQ(ours[0].rfind(kPathMissingType, 0), 0U) << ours[0];
        }
        const Workbook reloaded =
            std::string_view(extension) == ".xlsb" ? LoadXlsbBytes(saved.value()) : LoadBytes(saved.value());
        EXPECT_EQ(LinkTargets(reloaded), LinkTargets(wb));
        EXPECT_EQ(CellFormulas(reloaded), CellFormulas(wb));
        for (const ExternalLinkRecord& rec : reloaded.external_links()) {
          EXPECT_TRUE(rec.target_path_missing);
        }
      }
    }
  }
}

TEST(ExternalLinkPathMissing, BothExcelSavesLoadTheSameFormulas) {
  for (const char* stem : kPathMissingStems) {
    SCOPED_TRACE(stem);
    EXPECT_EQ(CellFormulas(LoadFixture(PathMissingStem(stem), ".xlsx")),
              CellFormulas(LoadFixture(PathMissingStem(stem), ".xlsb")));
  }
}

// Mac Excel 365 opening an authored book whose missing-path link target is spelled `<stem>` below, then saving it
// as .xlsx and .xlsb: the target is percent-decoded (a malformed escape kept as written) and re-encoded with only
// a space and `%` escaped.
//   target_escaped_percent  a%25b     kept
//   target_raw_percent      a%b       a%25b
//   target_escaped_space    a%20b     kept
//   target_space            My Sheet  My%20Sheet
constexpr const char* kPathMissingTargetStems[] = {"target_escaped_percent", "target_raw_percent",
                                                   "target_escaped_space", "target_space"};

TEST(ExternalLinkPathMissing, SavesEncodeTheTargetAsExcelDoes) {
  for (const char* stem : kPathMissingTargetStems) {
    const std::string saved_stem = std::string(stem) + "_saved";
    for (const char* source : {"", "_saved"}) {
      Workbook wb = LoadFixture(PathMissingStem(stem) + source, ".xlsx");
      for (const char* extension : {".xlsx", ".xlsb"}) {
        SCOPED_TRACE(std::string(stem) + source + " -> " + extension);
        auto saved = std::string_view(extension) == ".xlsb" ? io::xlsb::write_xlsb(wb) : io::write_ooxml(wb);
        ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
        const std::vector<std::uint8_t> excel = PathMissingBytes(saved_stem.c_str(), extension);
        EXPECT_EQ(Relationships(ZipPart(saved.value(), LinkRelsPart(1, extension).c_str())),
                  Relationships(ZipPart(excel, LinkRelsPart(1, extension).c_str())));
      }
    }
  }
}

}  // namespace
}  // namespace formulon
