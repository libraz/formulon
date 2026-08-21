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
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/external_book.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
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

  const io::ExternalBook& first = wb.external_links()[0].book;
  ASSERT_EQ(first.sheet_names.size(), 1U);
  EXPECT_EQ(first.sheet_names[0], "Data");
  ASSERT_NE(first.find_name("SrcTotal"), nullptr);
  EXPECT_TRUE(first.find_name("SrcTotal")->resolvable);
  // Excel caches only the cells this workbook references: A1 for the
  // direct reference and A3 for the one the name resolves to.
  EXPECT_TRUE(first.cached_cell(0, 0, 0).is_number());
  EXPECT_DOUBLE_EQ(first.cached_cell(0, 2, 0).as_number(), 30.0);

  const io::ExternalBook& second = wb.external_links()[1].book;
  ASSERT_NE(second.find_name("FarCell"), nullptr);
  EXPECT_DOUBLE_EQ(second.cached_cell(0, 6, 3).as_number(), 77.0);
}

TEST_P(ExternalLinkFixture, AnUncachedAddressReadsAsZeroNotBlank) {
  Workbook wb = LoadFixture("external_link_cell_kinds", GetParam());
  ASSERT_EQ(wb.external_links().size(), 1U);
  const io::ExternalBook& book = wb.external_links()[0].book;
  const std::uint32_t second = book.sheet_index("Second");
  ASSERT_NE(second, io::ExternalBook::kNoSheet);
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

INSTANTIATE_TEST_SUITE_P(BothContainerFormats, ExternalLinkFixture, ::testing::Values(".xlsx", ".xlsb"),
                         [](const ::testing::TestParamInfo<const char*>& info) {
                           return std::string(info.param).substr(1);
                         });

}  // namespace
}  // namespace formulon
