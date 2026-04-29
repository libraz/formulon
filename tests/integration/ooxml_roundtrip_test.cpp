// Copyright 2026 libraz. Licensed under the MIT License.
//
// Round-trip integration tests: writer -> reader. The earlier slice only
// covered sheet names; this slice exercises cell-level round-tripping
// for literals and formulas. Cells written via `Workbook::set_cell_*`
// must come back through `read_ooxml` with the same shape, and the
// reader must register formula cells with the recalc engine so a
// post-load `recalc()` reproduces the original cached values.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) { return io::ByteSpan{bytes.data(), bytes.size()}; }

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed: " << save_or.error().message;
  return save_or.value();
}

TEST(OoxmlRoundTrip, SingleSheet) {
  Workbook src = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;

  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Sheet1");
}

TEST(OoxmlRoundTrip, MultipleSheetsPreserveOrder) {
  Workbook src = Workbook::create_empty();
  src.add_sheet("Alpha");
  src.add_sheet("Beta");
  src.add_sheet("Gamma");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.sheet_count(), 3U);
  EXPECT_EQ(dst.sheet(0).name(), "Alpha");
  EXPECT_EQ(dst.sheet(1).name(), "Beta");
  EXPECT_EQ(dst.sheet(2).name(), "Gamma");
}

TEST(OoxmlRoundTrip, JapaneseSheetName) {
  Workbook src = Workbook::create_empty();
  // "売上" in UTF-8: E5 A3 B2 E4 B8 8A
  src.add_sheet("\xE5\xA3\xB2\xE4\xB8\x8A");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "\xE5\xA3\xB2\xE4\xB8\x8A");
}

TEST(OoxmlRoundTrip, SheetNameWithSpaceAndAmpersand) {
  Workbook src = Workbook::create_empty();
  src.add_sheet("Q1 & Q2 Summary");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  // The writer emits `&amp;` and pugixml decodes it; round-trip yields
  // the original ampersand.
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Q1 & Q2 Summary");
}

TEST(OoxmlRoundTrip, SheetNameWithSingleQuoteRequiresQuoting) {
  // Excel itself wraps such names in single quotes when emitting
  // formulas, but the workbook.xml `<sheet name=...>` attribute is just
  // an XML attribute and accepts the apostrophe directly. We verify the
  // name comes back unchanged.
  Workbook src = Workbook::create_empty();
  src.add_sheet("Joe's Notes");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Joe's Notes");
}

TEST(OoxmlRoundTrip, RenamedDefaultSheet) {
  Workbook src = Workbook::create();
  src.sheet(0).set_name("Renamed");
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  ASSERT_EQ(result_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "Renamed");
}

TEST(OoxmlRoundTrip, NumericLiteralAndFormulaRecalcMatches) {
  // Build a workbook with A1=42, A2==A1*2, recalc, save, read back,
  // recalc, assert the cached value at A2 is 84 again. This is the
  // canonical "true round-trip" check for the cell-aware reader.
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0U, 1U, 0U, "=A1*2")));
  ASSERT_TRUE(static_cast<bool>(src.recalc(eval::default_registry())));

  // Sanity: the source workbook recalc'd correctly.
  {
    const Cell* a2 = src.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(a2, nullptr);
    ASSERT_TRUE(a2->cached_value.is_number());
    EXPECT_DOUBLE_EQ(a2->cached_value.as_number(), 84.0);
  }

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;

  Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.sheet_count(), 1U);

  // After read but BEFORE recalc, A1 should already hold 42 (a literal),
  // and A2 should carry the formula text (the cached <v> emitted by the
  // writer is dropped by the reader on purpose; recalc populates it).
  {
    const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
    ASSERT_NE(a1, nullptr);
    ASSERT_TRUE(a1->cached_value.is_number());
    EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 42.0);
    EXPECT_TRUE(a1->formula_text.empty());

    const Cell* a2 = dst.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(a2, nullptr);
    EXPECT_EQ(a2->formula_text, "=A1*2");
  }

  // Recalc and verify A2 == 84.
  ASSERT_TRUE(static_cast<bool>(dst.recalc(eval::default_registry())));
  {
    const Cell* a2 = dst.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(a2, nullptr);
    ASSERT_TRUE(a2->cached_value.is_number());
    EXPECT_DOUBLE_EQ(a2->cached_value.as_number(), 84.0);
  }
}

TEST(OoxmlRoundTrip, InlineStringCellRoundTrips) {
  // The empty-workbook writer emits text values via t="inlineStr" (SST
  // is a Bundle 2.3 concern), so this is the round-trip path that
  // currently exists end-to-end without SST resolution.
  Workbook src = Workbook::create();
  std::string greeting = "Hello, world!";
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::text(greeting))));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  // No SST cells were used.
  EXPECT_EQ(result_or.value().pending_sst_count, 0U);

  const Workbook& dst = result_or.value().workbook;
  const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "Hello, world!");
}

TEST(OoxmlRoundTrip, ErrorCellRoundTrips) {
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::error(ErrorCode::Div0))));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;
  const Cell* a1 = dst.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_error());
  EXPECT_EQ(a1->cached_value.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace formulon
