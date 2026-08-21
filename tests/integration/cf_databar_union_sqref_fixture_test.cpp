//
// Checks the conditional-formatting reader against a real Excel 365
// (ja-JP) workbook (`tests/fixtures/excel/cf_databar_union_sqref.xlsx`).
//
// The CF reader and writer are otherwise driven entirely by hand-written
// XML strings, so nothing in the suite says what Excel actually emits.
// This fixture supplies that: Excel opened the package and saved it back,
// and the bytes under test are Excel's own serialisation.
//
// It pins one property in particular. A DataBar's richer settings live in
// a worksheet-level `<extLst>` overlay whose `<x14:conditionalFormatting>`
// carries its own `<xm:sqref>`, separate from the legacy
// `<conditionalFormatting sqref>` it extends. Whether those two are
// always equal — or whether the overlay may name a subset — decides
// whether the overlay's range can be derived from the model rather than
// tracked independently. Excel writes them equal, including when the
// range is a union and when Excel itself recomputed it across an
// inserted row and column.
//
// The union is the point. Two single rectangles look identical whatever
// the rule is, so only a union can expose a subset relationship. Both
// rules here therefore span two disjoint rectangles:
//
//   Bars!A1:A5 C1:C5   negative fill FFFF0000
//   Bars!E2:E6 G2:G6   negative fill FF00B050
//
// The legacy rule links to its overlay through a nested
// `<extLst><ext><x14:id>{GUID}</x14:id>` element rather than an `id`
// attribute; `CT_CfRule` has no such attribute and Excel discards the
// attribute spelling on re-save, taking the orphaned overlay with it.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

// Both rules as Excel spelled them, in the order the sheet lists them.
constexpr const char* kFirstSqref = "A1:A5 C1:C5";
constexpr const char* kSecondSqref = "E2:E6 G2:G6";

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

Workbook LoadFixture() {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_databar_union_sqref.xlsx";
  const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

// Round-trips `wb` through the writer and reader so the assertions can be
// re-run against our own serialisation.
Workbook RoundTrip(const Workbook& wb) {
  auto bytes_or = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml failed";
  if (!bytes_or) {
    return Workbook::create_empty();
  }
  const std::vector<std::uint8_t>& bytes = bytes_or.value();
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << "re-read failed";
  if (!read_or) {
    return Workbook::create_empty();
  }
  return std::move(read_or.value().workbook);
}

}  // namespace

TEST(CfDataBarUnionSqrefFixture, ExcelWritesEachRuleOverAUnionOfTwoRectangles) {
  Workbook wb = LoadFixture();
  ASSERT_GT(wb.sheet_count(), 0U);
  const std::vector<cf::ConditionalFormat>& blocks = wb.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 2U) << "expected one block per authored rule";

  // A1:A5 C1:C5 — two disjoint rectangles, not one merged span.
  ASSERT_EQ(blocks[0].sqref.size(), 2U);
  EXPECT_EQ(blocks[0].sqref[0].first.row, 0U);
  EXPECT_EQ(blocks[0].sqref[0].first.col, 0U);
  EXPECT_EQ(blocks[0].sqref[0].last.row, 4U);
  EXPECT_EQ(blocks[0].sqref[0].last.col, 0U);
  EXPECT_EQ(blocks[0].sqref[1].first.col, 2U);
  EXPECT_EQ(blocks[0].sqref[1].last.col, 2U);

  // E2:E6 G2:G6 — offset by a row so the two blocks cannot be confused.
  ASSERT_EQ(blocks[1].sqref.size(), 2U);
  EXPECT_EQ(blocks[1].sqref[0].first.row, 1U);
  EXPECT_EQ(blocks[1].sqref[0].first.col, 4U);
  EXPECT_EQ(blocks[1].sqref[0].last.row, 5U);
  EXPECT_EQ(blocks[1].sqref[1].first.col, 6U);
}

TEST(CfDataBarUnionSqrefFixture, EachRuleIsADataBarCarryingItsOverlayLink) {
  Workbook wb = LoadFixture();
  ASSERT_GT(wb.sheet_count(), 0U);
  const std::vector<cf::ConditionalFormat>& blocks = wb.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 2U);
  for (const cf::ConditionalFormat& block : blocks) {
    ASSERT_EQ(block.rules.size(), 1U);
    EXPECT_EQ(block.rules.front().type, cf::RuleType::DataBar);
    // The `<x14:id>` link Excel kept. Without it Excel drops the overlay,
    // so a non-empty id here is what makes the next test meaningful.
    EXPECT_FALSE(block.rules.front().id.empty()) << "Excel kept no x14:id link";
  }
  EXPECT_NE(blocks[0].rules.front().id, blocks[1].rules.front().id);
}

TEST(CfDataBarUnionSqrefFixture, TheOverlaySqrefMirrorsTheLegacySqref) {
  // The property the fixture exists for: Excel writes `<xm:sqref>` equal
  // to the legacy sqref rather than a subset of it, unions included.
  Workbook wb = LoadFixture();
  ASSERT_GT(wb.sheet_count(), 0U);
  const std::string& overlay = wb.sheet(0).ext_lst_xml();
  ASSERT_FALSE(overlay.empty()) << "Excel wrote no worksheet-level extLst";
  EXPECT_NE(overlay.find(std::string("<xm:sqref>") + kFirstSqref + "</xm:sqref>"), std::string::npos)
      << "first rule's overlay sqref is not the legacy union";
  EXPECT_NE(overlay.find(std::string("<xm:sqref>") + kSecondSqref + "</xm:sqref>"), std::string::npos)
      << "second rule's overlay sqref is not the legacy union";
}

TEST(CfDataBarUnionSqrefFixture, TheOverlayAndItsLinksSurviveOurOwnRoundTrip) {
  // The overlay is carried verbatim rather than re-derived, so a write /
  // read cycle must leave both the sqrefs and the id links intact.
  Workbook wb = LoadFixture();
  ASSERT_GT(wb.sheet_count(), 0U);
  Workbook again = RoundTrip(wb);
  ASSERT_GT(again.sheet_count(), 0U);

  const std::vector<cf::ConditionalFormat>& blocks = again.sheet(0).conditional_formats();
  ASSERT_EQ(blocks.size(), 2U);
  EXPECT_EQ(blocks[0].sqref.size(), 2U);
  EXPECT_EQ(blocks[1].sqref.size(), 2U);
  EXPECT_EQ(blocks[0].rules.front().id, wb.sheet(0).conditional_formats()[0].rules.front().id);

  const std::string& overlay = again.sheet(0).ext_lst_xml();
  EXPECT_NE(overlay.find(std::string("<xm:sqref>") + kFirstSqref + "</xm:sqref>"), std::string::npos);
  EXPECT_NE(overlay.find(std::string("<xm:sqref>") + kSecondSqref + "</xm:sqref>"), std::string::npos);
}

}  // namespace formulon
