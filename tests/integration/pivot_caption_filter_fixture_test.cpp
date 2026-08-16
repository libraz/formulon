//
// Checks the authored `<filters>` decode against a real Mac Excel 365
// workbook (`tests/fixtures/excel/pivot_caption_filter.xlsx`).
//
// The fixture is the reference file the decode was blocked on: a pivot
// Excel built over A1:C7 (Region on rows, Qtr on columns, sum of Amt),
// then filtered from the row-label dropdown with
// "label filters -> equals -> North". Excel's own bytes are what makes
// the expectations below measurements rather than a reading of the
// schema.
//
// Two properties of those bytes drive the whole design:
//
//   * The row field's `<items>` list still holds all three regions with
//     no `h="1"` on any of them. A caption filter is therefore not
//     expressible through `PivotItem::visible`, which is the persisted
//     filter surface the reader and writer already model — the rule
//     lives only in `<filters>`, and `<rowItems>` records the result.
//   * The criterion appears twice: `stringValue1="North"` on the
//     `<filter>` element, and the same text repeated as an `<autoFilter>`
//     pattern underneath.
//
// Fixture layout (`Sheet1`), with Excel's filtered render cached in the
// grid:
//   A1:C7  source table, headers Region / Qtr / Amt
//   E1:H4  the pivot — 行ラベル / North / 総計 rows, Q1 / Q2 / 総計 cols
//   Excel's values: North = 10, 5, 15; grand total = 10, 5, 15

#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

// Excel's cached render of the filtered pivot.
constexpr double kNorthQ1 = 10.0;
constexpr double kNorthQ2 = 5.0;
constexpr double kNorthTotal = 15.0;

std::string FixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_caption_filter.xlsx";
}

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
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath());
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

const pivot::PivotTable* FirstPivot(const Workbook& wb) {
  if (wb.sheet_count() == 0 || wb.sheet(0).pivot_tables().empty()) {
    return nullptr;
  }
  return wb.sheet(0).pivot_tables().front().get();
}

// ---------------------------------------------------------------------------
// (a) The decode, against Excel's own attribute spelling.
// ---------------------------------------------------------------------------

TEST(PivotCaptionFilterFixture, ExcelAuthoredFilterDecodesFromStringValue) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_EQ(table->authored_caption_filters().size(), 1U);
  const pivot::AuthoredCaptionFilter& filter = table->authored_caption_filters().front();
  EXPECT_EQ(filter.field_index, 0U);  // Region
  EXPECT_EQ(filter.predicate, pivot::CaptionPredicate::Equal);
  EXPECT_EQ(filter.value, "North");
  // A slicer selection is a different list and must stay empty.
  EXPECT_TRUE(table->active_filters().empty());
}

// The rule is not mirrored into item visibility, so an engine that only
// consults `PivotItem::visible` sees an unfiltered table. Pinning this
// keeps the decode from being "optimised away" later on the theory that
// the items list already covers it.
TEST(PivotCaptionFilterFixture, ExcelLeavesEveryRowItemVisible) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_FALSE(table->fields().empty());
  const pivot::PivotField& region = table->fields().front();
  ASSERT_EQ(region.items.size(), 3U);
  for (const pivot::PivotItem& item : region.items) {
    EXPECT_TRUE(item.visible) << "item '" << item.name << "' unexpectedly carries h=\"1\"";
  }
}

// ---------------------------------------------------------------------------
// (b) Evaluation reproduces Excel's filtered render.
// ---------------------------------------------------------------------------

TEST(PivotCaptionFilterFixture, EvaluationMatchesTheExcelFilteredRender) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  const pivot::PivotCache* cache = wb.find_pivot_cache(table->pivot_cache_id());
  ASSERT_NE(cache, nullptr);

  auto result_or = pivot::evaluate(*table, *cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const pivot::PivotResult& result = result_or.value();

  // South and East are filtered away; only North survives.
  ASSERT_EQ(result.rows.size(), 1U);
  EXPECT_EQ(result.rows.front().label, "North");
  ASSERT_EQ(result.values.size(), 1U);
  ASSERT_EQ(result.values.front().size(), 2U);
  EXPECT_DOUBLE_EQ(result.values[0][0][0].as_number(), kNorthQ1);
  EXPECT_DOUBLE_EQ(result.values[0][1][0].as_number(), kNorthQ2);
  ASSERT_TRUE(result.grand_total.is_number());
  EXPECT_DOUBLE_EQ(result.grand_total.as_number(), kNorthTotal);
}

// ---------------------------------------------------------------------------
// (c) Excel's own `<location ref>` survives a read -> write cycle.
// ---------------------------------------------------------------------------

// The writer projects the rendered extent into `ref` for a pivot built in
// memory, because such a table carries a placeholder span that understates
// the report and makes Excel terminate on refresh. That projection must not
// reach a `ref` Excel wrote: here the authored range is E1:H4 while the
// pivot renders a filtered single-region report, so a writer that reprojected
// unconditionally would silently shrink Excel's own bytes.
TEST(PivotCaptionFilterFixture, ExcelAuthoredLocationRefIsReEmittedUnchanged) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = FirstPivot(wb);
  ASSERT_NE(table, nullptr);
  ASSERT_TRUE(table->has_authored_span());

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{bytes_or.value().data(), bytes_or.value().size()})));
  auto part_or = zip.read_entry("xl/pivotTables/pivotTable1.xml");
  ASSERT_TRUE(static_cast<bool>(part_or)) << "missing pivotTable1.xml";
  const std::string part(part_or.value().begin(), part_or.value().end());
  EXPECT_NE(part.find("<location ref=\"E1:H4\""), std::string::npos) << part;
}

}  // namespace
}  // namespace formulon
