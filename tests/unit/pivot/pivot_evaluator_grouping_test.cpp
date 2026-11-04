// Axis grouping, placeholders, date buckets, and malformed-axis handling.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <map>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "pivot_evaluator_fixtures.h"
#include "styles.h"
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

// Date on rows, Amount summed.
PivotTable build_date_table() {
  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField date_f;
  date_f.source_name = "Date";
  date_f.axis = PivotAxis::Row;
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

// Groups Date (field 0) of `cache` by `spec` with ja-JP labels, as
// `set_date_group` does, and returns `build_date_table` showing the group
// fields on rows in its place, coarsest outermost.
Expected<PivotTable, Error> group_date_rows(PivotCache& cache, const DateGroupSpec& spec, bool date1904 = false) {
  Expected<std::size_t, Error> first = group_cache_by_date(cache, 0U, spec, date1904, mac_365_ja_jp_profile());
  if (!first) {
    return first.error();
  }
  PivotTable table = build_date_table();
  table.mutable_fields()[0].axis = PivotAxis::None;
  table.mutable_row_field_order().clear();
  for (std::size_t fi = first.value(); fi < cache.fields().size(); ++fi) {
    PivotField group_f;
    group_f.source_name = cache.fields()[fi].name;
    group_f.axis = PivotAxis::Row;
    group_f.sort.manual = true;
    for (std::size_t i = 0; i < cache.fields()[fi].field_group->group_items.size(); ++i) {
      PivotItem item;
      item.has_cache_index = true;
      item.cache_index = static_cast<std::uint32_t>(i);
      group_f.items.push_back(std::move(item));
    }
    table.mutable_fields().push_back(std::move(group_f));
    table.mutable_row_field_order().insert(table.mutable_row_field_order().begin(), static_cast<std::uint32_t>(fi));
  }
  resolve_pivot_names(table, cache);
  return Expected<PivotTable, Error>(std::move(table));
}

DateGroupSpec days_spec(std::optional<double> start, std::uint32_t interval_days = 7,
                        std::optional<double> end = std::nullopt) {
  DateGroupSpec spec;
  spec.periods = date_period_bit(GroupBy::Days);
  spec.interval_days = interval_days;
  spec.start = start;
  spec.end = end;
  return spec;
}

PivotResult evaluate_days_group(std::optional<double> start, std::uint32_t interval_days = 7,
                                std::optional<double> end = std::nullopt) {
  PivotCache cache = build_days_interval_cache();
  Expected<PivotTable, Error> table = group_date_rows(cache, days_spec(start, interval_days, end));
  EXPECT_TRUE(static_cast<bool>(table)) << (table ? std::string() : table.error().message);
  if (!table) {
    return PivotResult{};
  }
  auto r_or = evaluate(table.value(), cache);
  EXPECT_TRUE(static_cast<bool>(r_or)) << (r_or ? std::string() : r_or.error().message);
  return r_or ? r_or.value() : PivotResult{};
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

// Windows Excel 365 ja-JP, tests/fixtures/excel/win/pivot_week_1900: one
// record a day over serials 0..70, grouped by runs of seven days.
TEST(PivotEvaluator, DateGroupingByDaysIntervalAutoStart) {
  const PivotResult r = evaluate_days_group(/*start=*/std::nullopt);

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
  const PivotResult r = evaluate_days_group(/*start=*/0.0);
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
  const PivotResult r = evaluate_days_group(/*start=*/1.0);
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
  const PivotResult r = evaluate_days_group(/*start=*/61.0);
  ASSERT_EQ(r.rows.size(), 3U);
  EXPECT_EQ(r.rows[0].label, "<1900/3/1");
  EXPECT_EQ(r.rows[1].label, "1900/3/1 - 1900/3/7");
  EXPECT_EQ(r.rows[2].label, "1900/3/8 - 1900/3/11");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 61.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(r.values[2][0][0].as_number(), 3.0);
}

TEST(PivotEvaluator, DateGroupingByDaysIntervalExplicitEndAboveDataMaxAddsCatchAll) {
  // An explicit End (60) short of the data maximum (70) collapses
  // everything above it into one ">1900/2/29" catch-all, and the last
  // regular bucket clips to End. Serial 60 is Excel's phantom 1900/2/29.
  // Windows Excel 365: tests/fixtures/excel/win/pivot_days_end (end_60).
  const PivotResult r = evaluate_days_group(/*start=*/std::nullopt, /*interval_days=*/7, /*end=*/60.0);
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

TEST(PivotEvaluator, DateGroupingByDaysIntervalEndInsideTheDataClipsTheLastBucket) {
  // Windows Excel 365: tests/fixtures/excel/win/pivot_days_end (end_30).
  const PivotResult r = evaluate_days_group(/*start=*/std::nullopt, /*interval_days=*/7, /*end=*/30.0);
  ASSERT_EQ(r.rows.size(), 6U);
  EXPECT_EQ(r.rows[0].label, "1900/1/0 - 1900/1/6");
  EXPECT_EQ(r.rows[3].label, "1900/1/21 - 1900/1/27");
  EXPECT_EQ(r.rows[4].label, "1900/1/28 - 1900/1/30");
  EXPECT_EQ(r.rows[5].label, ">1900/1/30");
  for (std::size_t i = 0; i < 4; ++i) {
    EXPECT_DOUBLE_EQ(r.values[i][0][0].as_number(), 7.0) << "bucket " << i;
  }
  EXPECT_DOUBLE_EQ(r.values[4][0][0].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(r.values[5][0][0].as_number(), 40.0);
}

TEST(PivotEvaluator, DateGroupingRefusesValuesOutsideTheDateDomain) {
  for (const double raw : {-std::numeric_limits<double>::infinity(), std::numeric_limits<double>::infinity(),
                           std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::max(), -1.0}) {
    PivotCache cache = build_two_field_cache();
    push_record(cache, 0.0, 5.0);
    push_record(cache, raw, 1.0);
    const Expected<PivotTable, Error> table = group_date_rows(cache, days_spec(/*start=*/5.0, 1U, /*end=*/5.0));
    ASSERT_FALSE(static_cast<bool>(table)) << raw;
    EXPECT_EQ(table.error().code, FormulonErrorCode::kInvalidArgument) << raw;
    EXPECT_EQ(cache.fields().size(), 2U) << raw;
  }
}

TEST(PivotEvaluator, DateGroupingByDaysKeepsFractionalBoundsAndComparesDays) {
  const PivotResult r = evaluate_days_group(/*start=*/1.9, /*interval_days=*/7, /*end=*/60.9);
  ASSERT_EQ(r.rows.size(), 11U);
  EXPECT_EQ(r.rows[0].label, "<1900/1/1");
  EXPECT_EQ(r.rows[1].label, "1900/1/1 - 1900/1/7");
  EXPECT_EQ(r.rows[9].label, "1900/2/26 - 1900/2/29");
  EXPECT_EQ(r.rows[10].label, ">1900/2/29");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(r.values[9][0][0].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(r.values[10][0][0].as_number(), 10.0);
}

TEST(PivotEvaluator, DateGroupingLastDayWithATimeEndsTheRangeInEitherEpoch) {
  for (const bool date1904 : {false, true}) {
    const double last_valid = date_time::serial_from_ymd(9999, 12U, 31U, date1904);
    PivotCache cache = build_two_field_cache();
    push_record(cache, last_valid + 0.75, 1.0);
    Expected<PivotTable, Error> table = group_date_rows(cache, days_spec(std::nullopt), date1904);
    ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
    // A maximum with a time of day is the end itself.
    EXPECT_EQ(cache.fields()[2].field_group->range.end_date, "9999-12-31T18:00:00") << date1904;
    PivotFilterEnv env;
    env.date1904 = date1904;
    auto r_or = evaluate(table.value(), cache, PivotLayoutOptions{}, env);
    ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
    ASSERT_EQ(r_or.value().rows.size(), 1U) << date1904;
    EXPECT_EQ(r_or.value().rows[0].label, "9999/12/31 - 9999/12/31") << date1904;

    // The day after the last date is no date.
    PivotCache past = build_two_field_cache();
    push_record(past, last_valid + 1.0, 1.0);
    const Expected<PivotTable, Error> refused = group_date_rows(past, days_spec(std::nullopt), date1904);
    ASSERT_FALSE(static_cast<bool>(refused)) << date1904;
    EXPECT_EQ(refused.error().code, FormulonErrorCode::kInvalidArgument);
  }
}

TEST(PivotEvaluator, DateGroupingByDaysRejectsInvalidExplicitBounds) {
  const double last_valid = date_time::serial_from_ymd(9999, 12U, 31U);
  const double invalid_bounds[] = {
      -0.1,
      std::numeric_limits<double>::quiet_NaN(),
      std::numeric_limits<double>::infinity(),
      last_valid + 1.0,
  };
  for (const double invalid : invalid_bounds) {
    for (const bool invalid_start : {true, false}) {
      PivotCache cache = build_days_interval_cache();
      const Expected<PivotTable, Error> table = group_date_rows(
          cache,
          days_spec(invalid_start ? std::optional<double>(invalid) : std::optional<double>(0.0),
                    /*interval_days=*/7, invalid_start ? std::optional<double>(70.0) : std::optional<double>(invalid)));
      ASSERT_FALSE(static_cast<bool>(table)) << "invalid=" << invalid << " start=" << invalid_start;
      EXPECT_EQ(table.error().code, FormulonErrorCode::kInvalidArgument)
          << "invalid=" << invalid << " start=" << invalid_start;
    }
  }
}

TEST(PivotEvaluator, DateGroupingByDaysRejectsRawInvertedFractionalBounds) {
  PivotCache cache = build_days_interval_cache();
  const Expected<PivotTable, Error> table =
      group_date_rows(cache, days_spec(/*start=*/0.9, /*interval_days=*/7, /*end=*/0.1));
  ASSERT_FALSE(static_cast<bool>(table));
  EXPECT_EQ(table.error().code, FormulonErrorCode::kInvalidArgument);
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

FieldGroup YearsGroupOverFieldZero() {
  FieldGroup group;
  group.base = 0U;
  group.kind = GroupKind::Range;
  group.range.by = GroupBy::Years;
  return group;
}

// Date / Amount with a months group field over Date, materialised as a load
// leaves it.
PivotCache build_month_grouped_cache() {
  PivotCache cache = build_two_field_cache();
  push_record(cache, date_time::serial_from_ymd(2024, 1U, 15U), 100.0);
  push_record(cache, date_time::serial_from_ymd(2024, 3U, 10U), 200.0);
  push_record(cache, date_time::serial_from_ymd(2024, 1U, 25U), 80.0);
  FieldGroup months;
  months.base = 0U;
  months.kind = GroupKind::Range;
  months.range.by = GroupBy::Months;
  months.range.is_date = true;
  months.range.start_date = "2024-01-15T00:00:00";
  months.range.end_date = "2024-07-23T00:00:00";
  months.group_items.push_back(owned_text(cache, "<2024/1/15"));
  for (int m = 1; m <= 12; ++m) {
    months.group_items.push_back(owned_text(cache, std::to_string(m) + "\xE6\x9C\x88"));
  }
  months.group_items.push_back(owned_text(cache, ">2024/7/23"));
  cache.mutable_fields().push_back(PivotCacheField{"Months", {}});
  cache.mutable_fields().back().is_database_field = false;
  cache.mutable_fields().back().field_group = std::move(months);
  materialize_group_fields(cache, /*date1904=*/false);
  return cache;
}

PivotTable build_month_grouped_table() {
  PivotTable table = build_date_table();
  table.mutable_fields()[0].axis = PivotAxis::None;
  PivotField months_f;
  months_f.source_name = "Months";
  months_f.axis = PivotAxis::Row;
  table.mutable_fields().push_back(std::move(months_f));
  table.mutable_row_field_order() = {2};
  return table;
}

TEST(PivotEvaluator, GroupFieldLabelsItsNodesWithGroupItems) {
  const PivotCache cache = build_month_grouped_cache();
  const PivotTable table = build_month_grouped_table();

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "1\xE6\x9C\x88");
  EXPECT_EQ(r.rows[1].label, "3\xE6\x9C\x88");
  ASSERT_TRUE(r.rows[0].identity.has_value());
  EXPECT_EQ(*r.rows[0].identity, Value::number(1.0));
  EXPECT_EQ(*r.rows[1].identity, Value::number(3.0));
  EXPECT_TRUE(r.rows[0].number_format.empty());
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 180.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 200.0);
}

TEST(PivotEvaluator, GroupFieldManualOrderFollowsItsItems) {
  const PivotCache cache = build_month_grouped_cache();
  PivotTable table = build_month_grouped_table();
  PivotField& months_f = table.mutable_fields()[2];
  months_f.sort.manual = true;
  for (const std::uint32_t x : {3U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    months_f.items.push_back(item);
  }

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "3\xE6\x9C\x88");
  EXPECT_EQ(r.rows[1].label, "1\xE6\x9C\x88");
}

TEST(PivotEvaluator, HiddenGroupItemHidesItsRecords) {
  const PivotCache cache = build_month_grouped_cache();
  PivotTable table = build_month_grouped_table();
  for (std::uint32_t x = 0; x < 14U; ++x) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    item.visible = x != 1U;  // January
    table.mutable_fields()[2].items.push_back(item);
  }
  resolve_pivot_names(table, cache);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "3\xE6\x9C\x88");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);
}

TEST(PivotEvaluator, UngroupedDateItemsAreLabelledInTheFieldNumberFormat) {
  PivotCache cache = build_two_field_cache();
  push_record(cache, date_time::serial_from_ymd(2024, 2U, 8U), 150.0);
  push_record(cache, date_time::serial_from_ymd(2024, 1U, 15U), 100.0);
  cache.mutable_fields()[0].num_fmt_id = 14U;
  PivotTable table = build_date_table();
  PivotLayoutOptions options;
  options.profile = mac_365_ja_jp_profile();

  auto r_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "2024/1/15");
  EXPECT_EQ(r.rows[1].label, "2024/2/8");
  EXPECT_EQ(r.rows[0].number_format, "yyyy/m/d");
  ASSERT_TRUE(r.rows[0].identity.has_value());
  EXPECT_EQ(*r.rows[0].identity, Value::number(date_time::serial_from_ymd(2024, 1U, 15U)));

  // The pivot field's own format wins over the cache field's.
  table.mutable_fields()[0].number_format = "22";
  auto r2_or = evaluate(table, cache, options);
  ASSERT_TRUE(static_cast<bool>(r2_or)) << r2_or.error().message;
  EXPECT_EQ(r2_or.value().rows[0].label, "2024/1/15 0:00");
}

TEST(PivotEvaluator, CustomItemNumberFormatResolvesThroughTheWorkbookStyles) {
  PivotCache cache = build_two_field_cache();
  push_record(cache, date_time::serial_from_ymd(2024, 1U, 15U), 100.0);
  cache.mutable_fields()[0].num_fmt_id = 164U;
  PivotTable table = build_date_table();
  StylesTable styles;
  styles.num_fmt_strings.push_back("yyyy-mm-dd");
  styles.num_fmts.push_back(NumFmtRecord{164U, 0U});

  PivotFilterEnv env;
  env.styles = &styles;
  auto r_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "2024-01-15");
  EXPECT_EQ(r_or.value().rows[0].number_format, "yyyy-mm-dd");

  // Without the styles the custom id has no code and the item stays General.
  auto bare_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(bare_or)) << bare_or.error().message;
  EXPECT_EQ(bare_or.value().rows[0].label, "45306");
}

TEST(PivotEvaluator, UnformattedNumericItemsKeepTheirGeneralLabel) {
  PivotCache cache = build_two_field_cache();
  push_record(cache, 45306.0, 1.0);
  PivotTable table = build_date_table();
  cache.mutable_fields()[0].num_fmt_id = 0U;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "45306");
  EXPECT_TRUE(r_or.value().rows[0].number_format.empty());
}

TEST(PivotEvaluator, ItemCaptionRenamesTheLabelButNotTheOrderOrIdentity) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields()[0].shared_items = {owned_text(cache, "North"), owned_text(cache, "South")};
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  // South is renamed to sort first by label, which must not move it.
  for (const std::uint32_t x : {0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    item.caption = x == 1U ? "AAA" : "";
    table.mutable_fields()[0].items.push_back(item);
  }

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_EQ(r.rows[1].label, "AAA");
  ASSERT_TRUE(r.rows[1].identity.has_value());
  ASSERT_TRUE(r.rows[1].identity->is_text());
  EXPECT_EQ(r.rows[1].identity->as_text(), "South");
}

TEST(PivotEvaluator, PercentOfBaseItemFindsARenamedItem) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields()[0].shared_items = {owned_text(cache, "North"), owned_text(cache, "South")};
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  for (const std::uint32_t x : {0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    item.caption = x == 1U ? "AAA" : "";
    table.mutable_fields()[0].items.push_back(item);
  }
  resolve_pivot_names(table, cache);
  PivotDataField& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOf;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 1U;  // South, drawn as "AAA".

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.values.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0 / 500.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 1.0);
}

TEST(PivotEvaluator, PercentOfBaseItemFindsAFormattedDateAndAGroupItem) {
  // A formatted date item: its label is "2024/1/15", its name the serial's General text.
  PivotCache dates = build_two_field_cache();
  const double jan15 = date_time::serial_from_ymd(2024, 1U, 15U);
  const double feb8 = date_time::serial_from_ymd(2024, 2U, 8U);
  dates.mutable_fields()[0].shared_items = {Value::number(jan15), Value::number(feb8)};
  dates.mutable_fields()[0].num_fmt_id = 14U;
  push_record(dates, 0.0, 100.0);  // Shared-item indices.
  push_record(dates, 1.0, 150.0);
  PivotTable table = build_date_table();
  for (const std::uint32_t x : {0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    table.mutable_fields()[0].items.push_back(item);
  }
  resolve_pivot_names(table, dates);
  PivotDataField& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOf;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 0U;
  PivotLayoutOptions options;
  options.profile = mac_365_ja_jp_profile();
  auto r_or = evaluate(table, dates, options);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().values.size(), 2U);
  EXPECT_EQ(r_or.value().rows[0].label, "2024/1/15");
  EXPECT_DOUBLE_EQ(r_or.value().values[1][0][0].as_number(), 1.5);

  // A group item: base item x=3 is March.
  const PivotCache months = build_month_grouped_cache();
  PivotTable grouped = build_month_grouped_table();
  for (std::uint32_t x = 0; x < 14U; ++x) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    grouped.mutable_fields()[2].items.push_back(item);
  }
  resolve_pivot_names(grouped, months);
  PivotDataField& gdf = grouped.mutable_data_fields()[0];
  gdf.show_as = ShowValuesAs::PercentOf;
  gdf.show_as_base_field = 2U;
  gdf.show_as_base_item = 3U;
  auto g_or = evaluate(grouped, months);
  ASSERT_TRUE(static_cast<bool>(g_or)) << g_or.error().message;
  ASSERT_EQ(g_or.value().values.size(), 2U);
  EXPECT_DOUBLE_EQ(g_or.value().values[0][0][0].as_number(), 180.0 / 200.0);
  EXPECT_DOUBLE_EQ(g_or.value().values[1][0][0].as_number(), 1.0);
}

TEST(PivotEvaluator, LabelFilterOnAGroupFieldReadsTheGroupItemText) {
  const PivotCache cache = build_month_grouped_cache();
  PivotTable table = build_month_grouped_table();
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Months";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("1\xE6\x9C\x88");
  table.mutable_active_filters().push_back(f);
  AuthoredCaptionFilter caption;
  caption.field_index = 2U;
  caption.predicate = CaptionPredicate::Equal;
  caption.value = "1\xE6\x9C\x88";
  table.mutable_authored_caption_filters().push_back(caption);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "1\xE6\x9C\x88");
  EXPECT_TRUE(r_or.value().rows[0].identity_is_group_item);
}

TEST(PivotEvaluator, CaptionedNumericItemCarriesNoNumberFormat) {
  PivotCache cache = build_two_field_cache();
  const double jan15 = date_time::serial_from_ymd(2024, 1U, 15U);
  cache.mutable_fields()[0].shared_items = {Value::number(jan15)};
  cache.mutable_fields()[0].num_fmt_id = 14U;
  push_record(cache, 0.0, 100.0);
  PivotTable table = build_date_table();
  PivotItem item;
  item.has_cache_index = true;
  item.cache_index = 0U;
  item.caption = "Start";
  table.mutable_fields()[0].items.push_back(item);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "Start");
  EXPECT_TRUE(r_or.value().rows[0].number_format.empty());
}

TEST(PivotEvaluator, ExplicitGeneralPivotFieldFormatWinsOverTheCacheField) {
  PivotCache cache = build_two_field_cache();
  push_record(cache, 45306.0, 1.0);
  cache.mutable_fields()[0].num_fmt_id = 14U;
  PivotTable table = build_date_table();
  table.mutable_fields()[0].number_format = "0";

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "45306");
  EXPECT_TRUE(r_or.value().rows[0].number_format.empty());
}

TEST(PivotEvaluator, PercentOfBaseItemMatchesByCollationAndFallsBackToItsLabel) {
  // The bound shared item differs from the records' text only in case.
  PivotCache cache = build_basic_cache();
  cache.mutable_fields()[0].shared_items = {owned_text(cache, "NORTH"), owned_text(cache, "SOUTH")};
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  for (const std::uint32_t x : {0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    table.mutable_fields()[0].items.push_back(item);
  }
  PivotDataField& df = table.mutable_data_fields()[0];
  df.show_as = ShowValuesAs::PercentOf;
  df.show_as_base_field = 0U;
  df.show_as_base_item = 1U;
  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().values.size(), 2U);
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 175.0 / 500.0);

  // A years group: item 0 collects the values before the start, item 1 is 2018.
  PivotCache dates = build_date_cache();
  DateGroupSpec years;
  years.periods = date_period_bit(GroupBy::Years);
  Expected<PivotTable, Error> grouped_or = group_date_rows(dates, years);
  ASSERT_TRUE(static_cast<bool>(grouped_or)) << grouped_or.error().message;
  PivotTable& grouped = grouped_or.value();
  PivotDataField& gdf = grouped.mutable_data_fields()[0];
  gdf.show_as = ShowValuesAs::PercentOf;
  gdf.show_as_base_field = 2U;
  gdf.show_as_base_item = 1U;  // "2018年" = 10.
  auto g_or = evaluate(grouped, dates);
  ASSERT_TRUE(static_cast<bool>(g_or)) << g_or.error().message;
  ASSERT_EQ(g_or.value().rows.size(), 3U);
  EXPECT_EQ(g_or.value().rows[1].label, "2019\xE5\xB9\xB4");
  EXPECT_DOUBLE_EQ(g_or.value().values[1][0][0].as_number(), 140.0 / 10.0);
}

TEST(PivotEvaluator, HiddenRenamedItemIsMatchedByItsSourceName) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields()[0].shared_items = {owned_text(cache, "North"), owned_text(cache, "South")};
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  for (const std::uint32_t x : {0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    item.caption = x == 1U ? "AAA" : "";
    item.visible = x != 1U;
    table.mutable_fields()[0].items.push_back(item);
  }
  resolve_pivot_names(table, cache);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, GroupFieldWithoutStatedBoundsEvaluates) {
  PivotCache cache = build_basic_cache();
  cache.mutable_fields().push_back(PivotCacheField{"Region Years", {}});
  cache.mutable_fields().back().is_database_field = false;
  cache.mutable_fields().back().field_group = YearsGroupOverFieldZero();

  PivotTable table = build_sum_amount_table(/*row=*/{3}, /*col=*/{});
  PivotField grouped_f;
  grouped_f.source_name = "Region Years";
  grouped_f.axis = PivotAxis::Row;
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
