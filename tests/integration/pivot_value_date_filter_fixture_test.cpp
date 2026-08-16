//
// Checks the authored `<filters>` value and date decode against a real
// Excel 365 (ja-JP) workbook
// (`tests/fixtures/excel/pivot_value_date_filters.xlsx`).
//
// The fixture carries four pivots over one source table, each filtered
// from the row-label dropdown so that every decoded family appears
// exactly once. Excel's own bytes make the expectations below
// measurements rather than a reading of the schema, and its cached
// render of each report is what the evaluation is compared against.
//
// Source table (`Sheet5`, A1:D9) — Region / Product / Date / Amt:
//   North Alpha 2024-01-15 100   East Alpha 2024-01-25  80
//   North Beta  2024-02-20  50   East Beta  2024-06-18  20
//   South Alpha 2024-03-10 200   West Alpha 2024-02-08 150
//   South Beta  2024-05-05 300   West Beta  2024-07-22  25
// Region totals: North 150, South 500, East 100, West 175.
//
// Three properties of Excel's bytes drive the decode:
//
//   * The "top ten" dialog does not write `top10` as the filter type. An
//     item count is `type="count"` carrying a nested `<top10 val="N"/>`;
//     `percent` and `sum` are its other two flavours and rank on a
//     different quantity.
//   * These families carry no `stringValue*` shorthand, unlike a caption
//     filter. The criteria live only in the nested `<autoFilter>`, as one
//     `<customFilter>` for a threshold and two for a range, and a date
//     bound is written as a serial rather than as a formatted label.
//   * `<items>` again holds every value with no `h="1"` on any of them,
//     so the rule is not expressible through item visibility.
//
// East totals exactly 100, which is what makes the two threshold pivots
// decide the boundary: Excel drops East under `> 100` and keeps it under
// `between 100 and 175`.

#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
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

// Sheet index of each pivot inside the fixture, in the order Excel
// created them.
constexpr std::size_t kTopCountSheet = 0;
constexpr std::size_t kGreaterThanSheet = 1;
constexpr std::size_t kBetweenSheet = 2;
constexpr std::size_t kDateBetweenSheet = 3;

// Region aggregates, as Excel rendered them.
constexpr double kNorth = 150.0;
constexpr double kSouth = 500.0;
constexpr double kEast = 100.0;
constexpr double kWest = 175.0;

// Serials of 2024-01-01 and 2024-03-31, the bounds Excel wrote.
constexpr double kQ1Start = 45292.0;
constexpr double kQ1End = 45382.0;

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
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_value_date_filters.xlsx";
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

const pivot::PivotTable* PivotOn(const Workbook& wb, std::size_t sheet_index) {
  if (sheet_index >= wb.sheet_count() || wb.sheet(sheet_index).pivot_tables().empty()) {
    return nullptr;
  }
  return wb.sheet(sheet_index).pivot_tables().front().get();
}

// Evaluates the pivot on `sheet_index` and returns its row labels paired
// with the single data-field aggregate on each row.
std::vector<std::pair<std::string, double>> EvaluateRows(const Workbook& wb, std::size_t sheet_index) {
  std::vector<std::pair<std::string, double>> rows;
  const pivot::PivotTable* table = PivotOn(wb, sheet_index);
  EXPECT_NE(table, nullptr) << "no pivot on sheet " << sheet_index;
  if (table == nullptr) {
    return rows;
  }
  const pivot::PivotCache* cache = wb.find_pivot_cache(table->pivot_cache_id());
  EXPECT_NE(cache, nullptr);
  if (cache == nullptr) {
    return rows;
  }
  auto result_or = pivot::evaluate(*table, *cache);
  EXPECT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  if (!result_or) {
    return rows;
  }
  const pivot::PivotResult& result = result_or.value();
  for (std::size_t i = 0; i < result.rows.size() && i < result.values.size(); ++i) {
    double total = 0.0;
    for (const std::vector<Value>& slot : result.values[i]) {
      if (!slot.empty() && slot.front().is_number()) {
        total += slot.front().as_number();
      }
    }
    rows.emplace_back(result.rows[i].label, total);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// (a) The decode, against Excel's own attribute spelling.
// ---------------------------------------------------------------------------

TEST(PivotValueDateFilterFixture, TopCountDecodesFromTheNestedTop10Element) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = PivotOn(wb, kTopCountSheet);
  ASSERT_NE(table, nullptr);
  ASSERT_EQ(table->authored_value_filters().size(), 1U);
  const pivot::AuthoredValueFilter& filter = table->authored_value_filters().front();
  EXPECT_EQ(filter.field_index, 0U);  // Region
  EXPECT_EQ(filter.type, pivot::FilterType::ValueTop10);
  EXPECT_DOUBLE_EQ(filter.value, 2.0);
  EXPECT_FALSE(filter.value_high.has_value());
  EXPECT_EQ(filter.data_field_index, 0U);  // iMeasureFld
  // A slicer selection is a different list and must stay empty.
  EXPECT_TRUE(table->active_filters().empty());
  // The caption decoder must not also claim these entries.
  EXPECT_TRUE(table->authored_caption_filters().empty());
}

TEST(PivotValueDateFilterFixture, ThresholdAndRangeDecodeFromTheNestedCustomFilters) {
  Workbook wb = LoadFixture();

  const pivot::PivotTable* greater = PivotOn(wb, kGreaterThanSheet);
  ASSERT_NE(greater, nullptr);
  ASSERT_EQ(greater->authored_value_filters().size(), 1U);
  EXPECT_EQ(greater->authored_value_filters().front().type, pivot::FilterType::ValueGreaterThan);
  EXPECT_DOUBLE_EQ(greater->authored_value_filters().front().value, kEast);
  EXPECT_FALSE(greater->authored_value_filters().front().value_high.has_value());

  const pivot::PivotTable* between = PivotOn(wb, kBetweenSheet);
  ASSERT_NE(between, nullptr);
  ASSERT_EQ(between->authored_value_filters().size(), 1U);
  const pivot::AuthoredValueFilter& range = between->authored_value_filters().front();
  EXPECT_EQ(range.type, pivot::FilterType::ValueBetween);
  EXPECT_DOUBLE_EQ(range.value, kEast);
  ASSERT_TRUE(range.value_high.has_value());
  EXPECT_DOUBLE_EQ(*range.value_high, kWest);
}

TEST(PivotValueDateFilterFixture, DateBoundsDecodeAsSerialsOnTheDateField) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = PivotOn(wb, kDateBetweenSheet);
  ASSERT_NE(table, nullptr);
  ASSERT_EQ(table->authored_value_filters().size(), 1U);
  const pivot::AuthoredValueFilter& filter = table->authored_value_filters().front();
  EXPECT_EQ(filter.field_index, 2U);  // Date
  EXPECT_EQ(filter.type, pivot::FilterType::LabelDate);
  EXPECT_DOUBLE_EQ(filter.value, kQ1Start);
  ASSERT_TRUE(filter.value_high.has_value());
  EXPECT_DOUBLE_EQ(*filter.value_high, kQ1End);
}

// The rule is not mirrored into item visibility, so an engine that only
// consults `PivotItem::visible` sees an unfiltered table. Pinning this
// keeps the decode from being "optimised away" later on the theory that
// the items list already covers it.
TEST(PivotValueDateFilterFixture, ExcelLeavesEveryItemVisible) {
  Workbook wb = LoadFixture();
  for (const std::size_t sheet : {kTopCountSheet, kGreaterThanSheet, kBetweenSheet, kDateBetweenSheet}) {
    const pivot::PivotTable* table = PivotOn(wb, sheet);
    ASSERT_NE(table, nullptr) << "sheet " << sheet;
    for (const pivot::PivotField& field : table->fields()) {
      for (const pivot::PivotItem& item : field.items) {
        EXPECT_TRUE(item.visible) << "sheet " << sheet << " item '" << item.name << "' carries h=\"1\"";
      }
    }
  }
}

// ---------------------------------------------------------------------------
// (b) Evaluation reproduces Excel's filtered renders.
// ---------------------------------------------------------------------------

TEST(PivotValueDateFilterFixture, TopCountKeepsTheTwoLargestRegions) {
  const Workbook wb = LoadFixture();
  const std::vector<std::pair<std::string, double>> rows = EvaluateRows(wb, kTopCountSheet);
  ASSERT_EQ(rows.size(), 2U);
  EXPECT_EQ(rows[0].first, "South");
  EXPECT_DOUBLE_EQ(rows[0].second, kSouth);
  EXPECT_EQ(rows[1].first, "West");
  EXPECT_DOUBLE_EQ(rows[1].second, kWest);
}

// East totals exactly the threshold. Excel drops it, so the comparison is
// strict -- the one thing about this filter a schema reading cannot settle.
TEST(PivotValueDateFilterFixture, GreaterThanExcludesTheRegionOnTheThreshold) {
  const Workbook wb = LoadFixture();
  const std::vector<std::pair<std::string, double>> rows = EvaluateRows(wb, kGreaterThanSheet);
  ASSERT_EQ(rows.size(), 3U);
  EXPECT_EQ(rows[0].first, "North");
  EXPECT_DOUBLE_EQ(rows[0].second, kNorth);
  EXPECT_EQ(rows[1].first, "South");
  EXPECT_DOUBLE_EQ(rows[1].second, kSouth);
  EXPECT_EQ(rows[2].first, "West");
  EXPECT_DOUBLE_EQ(rows[2].second, kWest);
  for (const auto& row : rows) {
    EXPECT_NE(row.first, "East") << "a region equal to the threshold survived a strict comparison";
  }
}

// The mirror of the case above: both bounds of a range are inclusive, and
// this fixture puts a region on each of them.
TEST(PivotValueDateFilterFixture, BetweenKeepsBothBoundaryRegions) {
  const Workbook wb = LoadFixture();
  const std::vector<std::pair<std::string, double>> rows = EvaluateRows(wb, kBetweenSheet);
  ASSERT_EQ(rows.size(), 3U);
  EXPECT_EQ(rows[0].first, "East");
  EXPECT_DOUBLE_EQ(rows[0].second, kEast);
  EXPECT_EQ(rows[1].first, "North");
  EXPECT_DOUBLE_EQ(rows[1].second, kNorth);
  EXPECT_EQ(rows[2].first, "West");
  EXPECT_DOUBLE_EQ(rows[2].second, kWest);
}

TEST(PivotValueDateFilterFixture, DateBetweenKeepsOnlyTheFirstQuarter) {
  const Workbook wb = LoadFixture();
  const std::vector<std::pair<std::string, double>> rows = EvaluateRows(wb, kDateBetweenSheet);
  // Excel renders one row per surviving date: 1/15, 1/25, 2/8, 2/20, 3/10.
  ASSERT_EQ(rows.size(), 5U);
  double total = 0.0;
  for (const auto& row : rows) {
    total += row.second;
  }
  EXPECT_DOUBLE_EQ(total, 580.0);
}

}  // namespace
}  // namespace formulon
