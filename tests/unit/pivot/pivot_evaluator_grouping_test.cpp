// Axis grouping, placeholders, date buckets, and malformed-axis handling.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot_evaluator_fixtures.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

PivotCache build_date_cache() {
  // Excel serials (1900 leap-bug aware):
  //   2018-12-31 -> 43465  (Heisei)
  //   2019-04-30 -> 43585  (Heisei, last day of era)
  //   2019-05-01 -> 43586  (Reiwa, first day of era)
  //   2019-12-31 -> 43830  (Reiwa)
  //   2024-03-15 -> 45366  (Reiwa)
  //   2024-07-04 -> 45477  (Reiwa)
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](double serial, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(Value::number(serial));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };
  add(43465.0, 10.0);
  add(43585.0, 20.0);
  add(43586.0, 40.0);
  add(43830.0, 80.0);
  add(45366.0, 160.0);
  add(45477.0, 320.0);
  return cache;
}

PivotTable build_date_grouped_table(DateGrouping g, CalendarSystem cal) {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField date_f;
  date_f.source_name = "Date";
  date_f.axis = PivotAxis::Row;
  PivotDateGroup dg;
  dg.granularity = g;
  dg.calendar = cal;
  date_f.date_group = dg;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(date_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

PivotCache build_two_field_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  return cache;
}

void push_record(PivotCache& cache, double serial, double amount) {
  PivotCacheRecord rec;
  rec.cells.push_back(Value::number(serial));
  rec.cells.push_back(Value::number(amount));
  cache.mutable_records().push_back(std::move(rec));
}
PivotCache build_days_interval_cache() {
  PivotCache cache = build_two_field_cache();
  for (double serial = 0.0; serial <= 70.0; serial += 1.0) {
    push_record(cache, serial, 1.0);
  }
  return cache;
}

PivotTable build_days_grouped_table(std::optional<double> start_serial, std::uint32_t interval_days = 7,
                                    std::optional<double> end_serial = std::nullopt) {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField date_f;
  date_f.source_name = "Date";
  date_f.axis = PivotAxis::Row;
  PivotDateGroup dg;
  dg.granularity = DateGrouping::Days;
  dg.interval_days = interval_days;
  dg.start_serial = start_serial;
  dg.end_serial = end_serial;
  date_f.date_group = dg;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(date_f));
  table.mutable_fields().push_back(std::move(amount_f));
  table.mutable_row_field_order() = {0};
  PivotDataField sum;
  sum.name = "Sum of Amount";
  sum.field_index = 1;
  sum.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  return table;
}

TEST(PivotEvaluator, SparseRowColumnIntersectionIsBlank) {
  PivotCache cache = build_basic_cache();
  auto& records = cache.mutable_records();
  records.erase(std::remove_if(records.begin(), records.end(),
                               [](const PivotCacheRecord& record) {
                                 return record.cells[0].as_text() == "North" && record.cells[1].as_text() == "Gadget";
                               }),
                records.end());
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  const std::size_t north = row_index(result, "North");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_EQ(result.cols.size(), 2U);
  ASSERT_EQ(result.cols[0].label, "Gadget");
  EXPECT_TRUE(result.values[north][0][0].is_blank());
}

TEST(PivotEvaluator, UnresolvedHiddenItemDoesNotFilterBlankRecords) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_region;
  blank_region.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_region));
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  PivotItem malformed_hidden;
  malformed_hidden.visible = false;
  malformed_hidden.has_cache_index = true;
  malformed_hidden.cache_index = 999U;
  table.mutable_fields()[0].items.push_back(std::move(malformed_hidden));
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  const std::size_t blank_leaf = row_index(result, PivotLayoutOptions{}.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(result.values[blank_leaf][0][0].as_number(), 75.0);
}

TEST(PivotEvaluator, BlankAxisItemTakesThePlaceholderLabel) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_region;
  blank_region.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_region));
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();

  for (const AxisHierarchyNode& row : result.rows) {
    EXPECT_FALSE(row.label.empty());
  }
  const std::size_t blank_leaf = row_index(result, PivotLayoutOptions{}.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(result.values[blank_leaf][0][0].as_number(), 75.0);
  // The named groups are untouched.
  EXPECT_NE(row_index(result, "North"), static_cast<std::size_t>(-1));
  EXPECT_NE(row_index(result, "South"), static_cast<std::size_t>(-1));
}

TEST(PivotEvaluator, BlankAxisItemPlaceholderComesFromTheSuppliedVocabulary) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_region;
  blank_region.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_region));
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});

  PivotLayoutOptions options;
  options.blank_item_label = "<none>";
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  EXPECT_NE(row_index(result_or.value(), "<none>"), static_cast<std::size_t>(-1));
}

TEST(PivotEvaluator, BlankGroupSubtotalCarriesThePlaceholderLabel) {
  PivotCache cache = build_basic_cache();
  PivotCacheRecord blank_widget;
  blank_widget.cells = {Value::blank(), owned_text(cache, "Widget"), Value::number(75.0)};
  cache.mutable_records().push_back(std::move(blank_widget));
  PivotCacheRecord blank_gadget;
  blank_gadget.cells = {Value::blank(), owned_text(cache, "Gadget"), Value::number(25.0)};
  cache.mutable_records().push_back(std::move(blank_gadget));
  PivotTable table = build_sum_amount_table(/*row=*/{0, 1}, /*col=*/{});

  PivotLayoutOptions options;
  options.blank_item_label = "<no-value>";
  auto result_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();

  bool saw_blank_subtotal = false;
  for (const RowSubtotal& subtotal : result.row_subtotals) {
    ASSERT_FALSE(subtotal.labels.empty());
    if (subtotal.labels[0] == options.blank_item_label) {
      saw_blank_subtotal = true;
      EXPECT_DOUBLE_EQ(subtotal.values[0].as_number(), 100.0);  // 75 + 25.
    }
    EXPECT_FALSE(subtotal.labels[0].empty());
  }
  EXPECT_TRUE(saw_blank_subtotal);
}

TEST(PivotEvaluator, DateGroupingByYearGregorian) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Three Gregorian years: 2018, 2019, 2024 (chronological).
  ASSERT_EQ(r.rows.size(), 3U);
  EXPECT_EQ(r.rows[0].label, "2018");
  EXPECT_EQ(r.rows[1].label, "2019");
  EXPECT_EQ(r.rows[2].label, "2024");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 20.0 + 40.0 + 80.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 160.0 + 320.0);
}

TEST(PivotEvaluator, DateGroupingByQuarterGregorian) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Quarter, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 2018-Q4 (Dec), 2019-Q2 (Apr+May), 2019-Q4 (Dec), 2024-Q1 (Mar), 2024-Q3 (Jul).
  ASSERT_EQ(r.rows.size(), 5U);
  EXPECT_EQ(r.rows[0].label, "2018-Q4");
  EXPECT_EQ(r.rows[1].label, "2019-Q2");
  EXPECT_EQ(r.rows[2].label, "2019-Q4");
  EXPECT_EQ(r.rows[3].label, "2024-Q1");
  EXPECT_EQ(r.rows[4].label, "2024-Q3");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 20.0 + 40.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 80.0);
  EXPECT_DOUBLE_EQ(r.values[3][0][0].as_number(), 160.0);
  EXPECT_DOUBLE_EQ(r.values[4][0][0].as_number(), 320.0);
}

TEST(PivotEvaluator, DateGroupingByMonthGregorian) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Month, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Six distinct months across the dataset.
  ASSERT_EQ(r.rows.size(), 6U);
  EXPECT_EQ(r.rows[0].label, "2018-12");
  EXPECT_EQ(r.rows[1].label, "2019-04");
  EXPECT_EQ(r.rows[2].label, "2019-05");
  EXPECT_EQ(r.rows[3].label, "2019-12");
  EXPECT_EQ(r.rows[4].label, "2024-03");
  EXPECT_EQ(r.rows[5].label, "2024-07");
}

TEST(PivotEvaluator, DateGroupingByDay) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Day, CalendarSystem::Gregorian);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Each record is a distinct day -> 6 leaves in chronological order.
  ASSERT_EQ(r.rows.size(), 6U);
  EXPECT_EQ(r.rows[0].label, "2018-12-31");
  EXPECT_EQ(r.rows[1].label, "2019-04-30");
  EXPECT_EQ(r.rows[2].label, "2019-05-01");
  EXPECT_EQ(r.rows[3].label, "2019-12-31");
  EXPECT_EQ(r.rows[4].label, "2024-03-15");
  EXPECT_EQ(r.rows[5].label, "2024-07-04");
}

TEST(PivotEvaluator, DateGroupingByYearJapaneseCalendar) {
  PivotCache cache = build_date_cache();
  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Japanese);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // 2018 -> Heisei 30, 2019-04-30 -> Heisei 31, 2019-05-01 -> Reiwa 1,
  // 2019-12-31 -> Reiwa 1, 2024 -> Reiwa 6.
  // Buckets in chronological order: 平成30 / 平成31 / 令和1 / 令和6.
  ASSERT_EQ(r.rows.size(), 4U);
  // 平成 = E5 B9 B3 E6 88 90, 令和 = E4 BB A4 E5 92 8C, 年 = E5 B9 B4
  EXPECT_EQ(r.rows[0].label, std::string("\xE5\xB9\xB3\xE6\x88\x90") + "30" + "\xE5\xB9\xB4");
  EXPECT_EQ(r.rows[1].label, std::string("\xE5\xB9\xB3\xE6\x88\x90") + "31" + "\xE5\xB9\xB4");
  EXPECT_EQ(r.rows[2].label, std::string("\xE4\xBB\xA4\xE5\x92\x8C") + "1" + "\xE5\xB9\xB4");
  EXPECT_EQ(r.rows[3].label, std::string("\xE4\xBB\xA4\xE5\x92\x8C") + "6" + "\xE5\xB9\xB4");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 40.0 + 80.0);
  EXPECT_DOUBLE_EQ(r.values[3][0][0].as_number(), 160.0 + 320.0);
}

TEST(PivotEvaluator, DateGroupingNonNumericPassesThrough) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  // A blank (not a serial) should not crash; bucket_date is only invoked
  // for numeric values.
  PivotCacheRecord rec;
  rec.cells.push_back(Value::blank());
  rec.cells.push_back(Value::number(42.0));
  cache.mutable_records().push_back(std::move(rec));

  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  // The single blank-keyed bucket survives without exploding; label is the
  // empty string (per `display_string` for Blank).
  ASSERT_EQ(r_or.value().rows.size(), 1U);
}

TEST(PivotEvaluator, DateGroupingByHour) {
  // Two records on 2024-03-15 within different hours, plus a third in the
  // same hour as the second so the second bucket sums two rows.
  const double base = 45366.0;  // 2024-03-15
  PivotCache cache = build_two_field_cache();
  push_record(cache, base + 9.0 / 24.0 + 30.0 / 1440.0, 1.0);   // 09:30
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0, 2.0);   // 10:05
  push_record(cache, base + 10.0 / 24.0 + 45.0 / 1440.0, 4.0);  // 10:45

  PivotTable table = build_date_grouped_table(DateGrouping::Hour, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-15 09");
  EXPECT_EQ(r.rows[1].label, "2024-03-15 10");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 2.0 + 4.0);
}

TEST(PivotEvaluator, DateGroupingByMinute) {
  // Two minutes within the same hour; chronological ordering is preserved.
  const double base = 45366.0;  // 2024-03-15
  PivotCache cache = build_two_field_cache();
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0, 1.0);   // 10:05
  push_record(cache, base + 10.0 / 24.0 + 45.0 / 1440.0, 2.0);  // 10:45
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0, 4.0);   // 10:05 again

  PivotTable table = build_date_grouped_table(DateGrouping::Minute, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-15 10:05");
  EXPECT_EQ(r.rows[1].label, "2024-03-15 10:45");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0 + 4.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, DateGroupingBySecond) {
  // Two seconds within the same minute.
  const double base = 45366.0;  // 2024-03-15
  PivotCache cache = build_two_field_cache();
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0 + 12.0 / 86400.0, 1.0);  // 10:05:12
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0 + 47.0 / 86400.0, 2.0);  // 10:05:47
  push_record(cache, base + 10.0 / 24.0 + 5.0 / 1440.0 + 12.0 / 86400.0, 4.0);  // 10:05:12 again

  PivotTable table = build_date_grouped_table(DateGrouping::Second, CalendarSystem::Gregorian);
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024-03-15 10:05:12");
  EXPECT_EQ(r.rows[1].label, "2024-03-15 10:05:47");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0 + 4.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, DateGroupingByYearHonorsDate1904Epoch) {
  PivotCache cache = build_two_field_cache();
  push_record(cache, date_time::serial_from_ymd(2024, 3, 15, /*date1904=*/true), 10.0);
  push_record(cache, date_time::serial_from_ymd(2023, 6, 1, /*date1904=*/true), 20.0);

  PivotTable table = build_date_grouped_table(DateGrouping::Year, CalendarSystem::Gregorian);
  PivotFilterEnv env;
  env.date1904 = true;
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2023");
  EXPECT_EQ(r.rows[1].label, "2024");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 10.0);
}

TEST(PivotEvaluator, DateGroupingByDaysIntervalAutoStart) {
  PivotCache cache = build_days_interval_cache();
  PivotTable table = build_days_grouped_table(/*start_serial=*/std::nullopt);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Ten full 7-day buckets from the data minimum (serial 0) plus one
  // trailing bucket clipped to the resolved auto End (data max + 1),
  // which is one day past the last record (serial 70 = "1900/3/10").
  ASSERT_EQ(r.rows.size(), 11U);
  EXPECT_EQ(r.rows[0].label, "1900/1/0 - 1900/1/6");
  EXPECT_EQ(r.rows[1].label, "1900/1/7 - 1900/1/13");
  EXPECT_EQ(r.rows[2].label, "1900/1/14 - 1900/1/20");
  EXPECT_EQ(r.rows[3].label, "1900/1/21 - 1900/1/27");
  EXPECT_EQ(r.rows[4].label, "1900/1/28 - 1900/2/3");
  EXPECT_EQ(r.rows[5].label, "1900/2/4 - 1900/2/10");
  EXPECT_EQ(r.rows[6].label, "1900/2/11 - 1900/2/17");
  EXPECT_EQ(r.rows[7].label, "1900/2/18 - 1900/2/24");
  EXPECT_EQ(r.rows[8].label, "1900/2/25 - 1900/3/2");
  EXPECT_EQ(r.rows[9].label, "1900/3/3 - 1900/3/9");
  EXPECT_EQ(r.rows[10].label, "1900/3/10 - 1900/3/11");
  for (std::size_t i = 0; i < 10; ++i) {
    EXPECT_DOUBLE_EQ(r.values[i][0][0].as_number(), 7.0) << "bucket " << i;
  }
  EXPECT_DOUBLE_EQ(r.values[10][0][0].as_number(), 1.0);
}

TEST(PivotEvaluator, DateGroupingByDaysIntervalExplicitStartEqualsMinMatchesAuto) {
  // Start=0 explicit coincides with the data minimum, so Excel renders the
  // identical bucket set as the auto case above.
  PivotCache cache = build_days_interval_cache();
  PivotTable table = build_days_grouped_table(/*start_serial=*/0.0);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 11U);
  EXPECT_EQ(r.rows[0].label, "1900/1/0 - 1900/1/6");
  EXPECT_EQ(r.rows[10].label, "1900/3/10 - 1900/3/11");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(r.values[10][0][0].as_number(), 1.0);
}

TEST(PivotEvaluator, DateGroupingByDaysIntervalExplicitStartAboveMinAddsCatchAll) {
  // Start=1 is one past the data minimum (serial 0): a catch-all bucket
  // absorbs everything below it, and since (71-1) records divide evenly
  // by 7 the trailing regular bucket is NOT clipped here, unlike the
  // auto/Start=0 case above.
  PivotCache cache = build_days_interval_cache();
  PivotTable table = build_days_grouped_table(/*start_serial=*/1.0);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 11U);
  EXPECT_EQ(r.rows[0].label, "<1900/1/1");
  EXPECT_EQ(r.rows[1].label, "1900/1/1 - 1900/1/7");
  EXPECT_EQ(r.rows[2].label, "1900/1/8 - 1900/1/14");
  EXPECT_EQ(r.rows[3].label, "1900/1/15 - 1900/1/21");
  EXPECT_EQ(r.rows[4].label, "1900/1/22 - 1900/1/28");
  EXPECT_EQ(r.rows[5].label, "1900/1/29 - 1900/2/4");
  EXPECT_EQ(r.rows[6].label, "1900/2/5 - 1900/2/11");
  EXPECT_EQ(r.rows[7].label, "1900/2/12 - 1900/2/18");
  EXPECT_EQ(r.rows[8].label, "1900/2/19 - 1900/2/25");
  EXPECT_EQ(r.rows[9].label, "1900/2/26 - 1900/3/3");
  EXPECT_EQ(r.rows[10].label, "1900/3/4 - 1900/3/10");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);
  for (std::size_t i = 1; i < 11; ++i) {
    EXPECT_DOUBLE_EQ(r.values[i][0][0].as_number(), 7.0) << "bucket " << i;
  }
}

TEST(PivotEvaluator, DateGroupingByDaysIntervalExplicitStartDeepInDataClipsLastBucket) {
  // Start=61 puts 61 of the 71 records into the catch-all; the trailing
  // regular bucket clips to the resolved auto End (data max + 1), not to
  // the true data max, so its label spans one day past the last record.
  PivotCache cache = build_days_interval_cache();
  PivotTable table = build_days_grouped_table(/*start_serial=*/61.0);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 3U);
  EXPECT_EQ(r.rows[0].label, "<1900/3/1");
  EXPECT_EQ(r.rows[1].label, "1900/3/1 - 1900/3/7");
  EXPECT_EQ(r.rows[2].label, "1900/3/8 - 1900/3/11");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 61.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 3.0);
}

TEST(PivotEvaluator, DateGroupingByDaysIntervalExplicitEndAboveDataMaxAddsCatchAll) {
  // Mirror image of ExplicitStartAboveMinAddsCatchAll: an explicit End
  // (60) short of the data maximum (70) collapses everything above it
  // into one ">1900/2/29" catch-all, and the last regular bucket clips
  // to End the same way the auto-End case clips to the data maximum.
  // Serial 60 is "1900/2/29" (Excel's 1900 system has a phantom leap
  // day real Gregorian 1900 does not), not "1900/3/1" -- confirmed by
  // pivot_week_1900's own auto-End bucket "1900/2/25 - 1900/3/2",
  // which spans serials 56-62 including the ghost day. Not in the
  // fixture itself (which only varies Start, never End); label shape
  // mirrors the measured `<start` bucket, but `>end` itself is
  // unmeasured.
  PivotCache cache = build_days_interval_cache();
  PivotTable table = build_days_grouped_table(/*start_serial=*/std::nullopt, /*interval_days=*/7,
                                              /*end_serial=*/60.0);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 10U);
  EXPECT_EQ(r.rows[0].label, "1900/1/0 - 1900/1/6");
  EXPECT_EQ(r.rows[7].label, "1900/2/18 - 1900/2/24");
  EXPECT_EQ(r.rows[8].label, "1900/2/25 - 1900/2/29");
  EXPECT_EQ(r.rows[9].label, ">1900/2/29");
  for (std::size_t i = 0; i < 8; ++i) {
    EXPECT_DOUBLE_EQ(r.values[i][0][0].as_number(), 7.0) << "bucket " << i;
  }
  EXPECT_DOUBLE_EQ(r.values[8][0][0].as_number(), 5.0);   // 1900/2/25..1900/3/1: 5 records.
  EXPECT_DOUBLE_EQ(r.values[9][0][0].as_number(), 10.0);  // 1900/3/2..1900/3/11: 10 records.
}

TEST(PivotEvaluator, CacheIdMismatchYieldsMissing) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_pivot_cache_id(99);  // Cache has id 1.

  auto r_or = evaluate(table, cache);
  ASSERT_FALSE(static_cast<bool>(r_or));
  EXPECT_EQ(r_or.error().code, FormulonErrorCode::kEvalPivotMissing);
}

TEST(PivotEvaluator, OutOfRangeFieldIndexYieldsInvalid) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.mutable_data_fields()[0].field_index = 99;

  auto r_or = evaluate(table, cache);
  ASSERT_FALSE(static_cast<bool>(r_or));
  EXPECT_EQ(r_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
}

TEST(PivotEvaluator, GroupingDerivedAxisFieldWithNoDateGroupIsRefused) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields().push_back(PivotCacheField{"Region Years", {}});
  cache.mutable_fields().back().is_database_field = false;
  cache.mutable_fields().back().field_group_xml = "<fieldGroup base=\"0\"><rangePr groupBy=\"years\"/></fieldGroup>";

  PivotTable table = build_sum_amount_table(/*row=*/{3}, /*col=*/{});
  PivotField grouped_f;
  grouped_f.source_name = "Region Years";
  grouped_f.axis = PivotAxis::Row;
  table.mutable_fields().push_back(std::move(grouped_f));
  table.mutable_row_field_order() = {3};

  auto r_or = evaluate(table, cache);
  ASSERT_FALSE(static_cast<bool>(r_or));
  EXPECT_EQ(r_or.error().code, FormulonErrorCode::kEvalPivotInvalid);
}

TEST(PivotEvaluator, GroupingDerivedAxisFieldWithDateGroupSetEvaluatesNormally) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields().push_back(PivotCacheField{"Region Years", {}});
  cache.mutable_fields().back().is_database_field = false;
  cache.mutable_fields().back().field_group_xml = "<fieldGroup base=\"0\"><rangePr groupBy=\"years\"/></fieldGroup>";

  PivotTable table = build_sum_amount_table(/*row=*/{3}, /*col=*/{});
  PivotField grouped_f;
  grouped_f.source_name = "Region Years";
  grouped_f.axis = PivotAxis::Row;
  grouped_f.date_group = PivotDateGroup{};
  table.mutable_fields().push_back(std::move(grouped_f));
  table.mutable_row_field_order() = {3};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
}

TEST(PivotRecordAccess, OutOfDomainSharedItemIndexCollapsesToBlank) {
  // One shared item, so index 0 is the only resolvable reference. A cache
  // built through the mutation API leaves `cell_is_index` empty, which is
  // what makes a numeric cell in a shared field be read as an index.
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {Value::number(42.0)}});

  const struct {
    double stored;
    bool resolves;
  } cases[] = {
      {std::numeric_limits<double>::quiet_NaN(), false},
      {std::numeric_limits<double>::infinity(), false},
      {-std::numeric_limits<double>::infinity(), false},
      {-1.0, false},
      {1e30, false},
      {4.3e9, false},  // Past the wasm32 `size_t` range as well as past the container.
      {1.0, false},    // One past the only shared item.
      {0.0, true},
      {0.5, true},
  };

  for (const auto& c : cases) {
    PivotCacheRecord record;
    record.cells.push_back(Value::number(c.stored));
    ASSERT_TRUE(record.cell_is_index.empty());
    const Value v = cell_value(cache, record, 0);
    if (c.resolves) {
      ASSERT_TRUE(v.is_number()) << "stored=" << c.stored;
      EXPECT_DOUBLE_EQ(v.as_number(), 42.0) << "stored=" << c.stored;
    } else {
      EXPECT_TRUE(v.is_blank()) << "stored=" << c.stored;
    }
  }
}

}  // namespace
}  // namespace formulon::pivot
