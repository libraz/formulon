//
// Checks the authored `<filters>` recurring-period decode against a real
// Excel 365 (ja-JP) workbook
// (`tests/fixtures/excel/pivot_recurring_period_filter.xlsx`).
//
// The recurring selectors — `M1`..`M12` and `Q1`..`Q4` — name a position
// in the calendar rather than a span of it. `Q2` keeps every April, May
// and June of every year, so unlike the relative-period family it needs
// no clock reading to resolve and unlike the absolute `dateBetween`
// family it carries no serial bounds. That is why it is modelled apart
// from both, as `PivotTable::authored_recurring_filters()`.
//
// Source table (`Sheet1`, A1:D9) — Region / Product / Date / Amt:
//   North Alpha 2024-01-15 100   East Alpha 2024-01-25  80
//   North Beta  2024-02-20  50   East Beta  2024-06-18  20
//   South Alpha 2024-03-10 200   West Alpha 2024-02-08 150
//   South Beta  2024-05-05 300   West Beta  2024-07-22  25
//
// The pivot under test lays Date down the rows with sum of Amt, so the
// surviving row labels name the selected dates outright. Of the eight
// source rows only 2024-05-05 and 2024-06-18 fall in the second quarter,
// which is what makes this selector's result unambiguous: two rows at 300
// and 20, and Excel's cached `<rowItems count="3">` (two dates plus the
// grand total) agrees.
//
// The year-independence is the property that separates a recurring
// selector from a window, and the fixture pins it: the dates are from
// 2024 and the file selects them whatever year it is read in.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "eval/date_time.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pivot/filter_engine.h"
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

// The pivot Excel filtered is the one that places Date on the row axis;
// it is the fourth the workbook created.
constexpr std::size_t kRecurringSheet = 3;

// The two Amt values inside Q2, and their total as Excel rendered it.
constexpr double kMay = 300.0;
constexpr double kJune = 20.0;
constexpr double kQ2Total = 320.0;

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
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_recurring_period_filter.xlsx";
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

// Evaluates the pivot on `sheet_index` against `env` and returns its row
// labels paired with the single data-field aggregate on each row.
std::vector<std::pair<std::string, double>> EvaluateRows(const Workbook& wb, std::size_t sheet_index,
                                                         const pivot::PivotFilterEnv& env) {
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
  auto result_or = pivot::evaluate(*table, *cache, pivot::PivotLayoutOptions{}, env);
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

}  // namespace

TEST(PivotRecurringPeriodFilterFixture, QuarterSelectorDecodesToItsThreeCalendarMonths) {
  Workbook wb = LoadFixture();
  const pivot::PivotTable* table = PivotOn(wb, kRecurringSheet);
  ASSERT_NE(table, nullptr);
  ASSERT_EQ(table->authored_recurring_filters().size(), 1U);
  const pivot::AuthoredRecurringFilter& filter = table->authored_recurring_filters().front();
  EXPECT_EQ(filter.field_index, 2U);  // Date
  EXPECT_EQ(filter.month_low, 4U);
  EXPECT_EQ(filter.month_high, 6U);
  // A recurring entry is neither a window nor a bounded criterion, so no
  // sibling decoder may also claim it.
  EXPECT_TRUE(table->authored_period_filters().empty());
  EXPECT_TRUE(table->authored_value_filters().empty());
  EXPECT_TRUE(table->authored_caption_filters().empty());
  EXPECT_TRUE(table->active_filters().empty());
}

TEST(PivotRecurringPeriodFilterFixture, EvaluationKeepsTheSameRowsExcelCached) {
  Workbook wb = LoadFixture();
  const auto rows = EvaluateRows(wb, kRecurringSheet, pivot::PivotFilterEnv{});
  ASSERT_EQ(rows.size(), 2U) << "expected only the two second-quarter dates";
  EXPECT_DOUBLE_EQ(rows[0].second, kMay);
  EXPECT_DOUBLE_EQ(rows[1].second, kJune);
  EXPECT_DOUBLE_EQ(rows[0].second + rows[1].second, kQ2Total);
}

TEST(PivotRecurringPeriodFilterFixture, TheResultDoesNotMoveWithTheClock) {
  // The property that makes this family recurring rather than relative.
  // A window filter over 2024 dates would be empty at both readings.
  Workbook wb = LoadFixture();
  pivot::PivotFilterEnv early;
  early.pinned_now = eval::date_time::CivilTime{{2024, 5U, 20U}, {0U, 0U, 0U}};
  pivot::PivotFilterEnv late;
  late.pinned_now = eval::date_time::CivilTime{{2099, 11U, 3U}, {0U, 0U, 0U}};

  const auto early_rows = EvaluateRows(wb, kRecurringSheet, early);
  const auto late_rows = EvaluateRows(wb, kRecurringSheet, late);
  ASSERT_EQ(early_rows.size(), 2U);
  ASSERT_EQ(late_rows.size(), 2U);
  EXPECT_EQ(early_rows[0].first, late_rows[0].first);
  EXPECT_EQ(early_rows[1].first, late_rows[1].first);
  EXPECT_DOUBLE_EQ(early_rows[0].second, kMay);
  EXPECT_DOUBLE_EQ(late_rows[0].second, kMay);
}

}  // namespace formulon
