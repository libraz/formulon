// Copyright 2026 libraz. Licensed under the MIT License.
//
// Round-trip integration tests: writer -> reader (Bundle 2.1 slice).
// Ensures `Workbook::save()` output is consumable by `read_ooxml`,
// preserving sheet names and order. Cell parsing is out of scope until
// later bundles, so the assertions are limited to the workbook-level
// metadata the reader actually populates.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
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

}  // namespace
}  // namespace formulon
