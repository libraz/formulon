// Numeric pivot grouping: the group field it appends, its range items, the
// pivot fields and axes it rewires, ungrouping, and reading a date or numeric
// grouping back.

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon::pivot {
namespace {

constexpr std::size_t kRegion = 0U;
constexpr std::size_t kAmt = 1U;
constexpr std::size_t kQty = 2U;
constexpr std::size_t kGroup = 3U;

// Region (shared) / Amt (inline, `amounts`) / Qty (inline); each table puts
// Amt then Region on rows and sums Qty.
Workbook AmountWorkbook(const std::vector<Value>& amounts, std::size_t tables = 1U,
                        ExcelProfile profile = mac_365_ja_jp_profile()) {
  Workbook wb = Workbook::create_for_profile(profile);
  auto cache = std::make_unique<PivotCache>();
  cache->set_cache_id(1U);
  cache->mutable_fields().push_back(PivotCacheField{"Region", {Value::text("North"), Value::text("South")}});
  cache->mutable_fields().push_back(PivotCacheField{"Amt", {}});
  cache->mutable_fields().push_back(PivotCacheField{"Qty", {}});
  for (std::size_t i = 0; i < amounts.size(); ++i) {
    PivotCacheRecord record;
    record.cells = {Value::number(static_cast<double>(i % 2U)), amounts[i], Value::number(1.0)};
    record.cell_is_index = {true, false, false};
    cache->mutable_records().push_back(std::move(record));
  }
  wb.add_pivot_cache(std::move(cache));
  for (std::size_t t = 0; t < tables; ++t) {
    auto table = std::make_unique<PivotTable>();
    table->set_name("PT" + std::to_string(t + 1U));
    table->set_pivot_cache_id(1U);
    for (const char* name : {"Region", "Amt", "Qty"}) {
      PivotField field;
      field.source_name = name;
      table->mutable_fields().push_back(std::move(field));
    }
    table->mutable_fields()[kAmt].axis = PivotAxis::Row;
    table->mutable_fields()[kRegion].axis = PivotAxis::Row;
    table->mutable_fields()[kQty].axis = PivotAxis::Value;
    PivotDataField sum;
    sum.name = "合計 / Qty";
    sum.field_index = static_cast<std::uint32_t>(kQty);
    table->mutable_data_fields().push_back(std::move(sum));
    table->mutable_row_field_order() = {static_cast<std::uint32_t>(kAmt), static_cast<std::uint32_t>(kRegion)};
    table->set_anchor(2U, 6U * static_cast<std::uint32_t>(t), 10U, 2U);
    wb.sheet(0).add_pivot_table(std::move(table));
  }
  return wb;
}

std::vector<Value> Numbers(std::initializer_list<double> values) {
  std::vector<Value> out;
  for (const double v : values) {
    out.push_back(Value::number(v));
  }
  return out;
}

const PivotCache& CacheOf(const Workbook& wb) {
  return *wb.pivot_caches().front();
}

std::vector<std::string> GroupTexts(const PivotCacheField& field) {
  std::vector<std::string> texts;
  for (const Value& item : field.field_group->group_items) {
    texts.emplace_back(item.is_text() ? item.as_text() : std::string_view("?"));
  }
  return texts;
}

std::vector<double> GroupNumbers(const PivotCache& cache) {
  std::vector<double> numbers;
  for (const PivotCacheRecord& record : cache.records()) {
    numbers.push_back(
        kGroup < record.cells.size() && record.cells[kGroup].is_number() ? record.cells[kGroup].as_number() : -1.0);
  }
  return numbers;
}

// The label of the group item each record falls in.
std::vector<std::string> RecordLabels(const PivotCache& cache) {
  const std::vector<std::string> texts = GroupTexts(cache.fields()[kGroup]);
  std::vector<std::string> labels;
  for (const double n : GroupNumbers(cache)) {
    labels.push_back(n >= 0.0 && n < static_cast<double>(texts.size()) ? texts[static_cast<std::size_t>(n)] : "?");
  }
  return labels;
}

TEST(NumericGroup, AutoBoundsTakeTheDataAndLabelEachRange) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 55, 140, 300}), 2U);
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  const PivotCache& cache = CacheOf(wb);
  ASSERT_EQ(cache.fields().size(), 4U);
  const PivotCacheField& group = cache.fields()[kGroup];
  EXPECT_EQ(group.name, "Amt2");
  EXPECT_FALSE(group.is_database_field);
  EXPECT_EQ(group.field_group->base.value_or(99U), kAmt);
  EXPECT_EQ(cache.fields()[kAmt].field_group->par.value_or(99U), kGroup);
  const RangeGroup& range = group.field_group->range;
  EXPECT_FALSE(range.is_date);
  EXPECT_EQ(range.by, GroupBy::Range);
  EXPECT_EQ(range.start_num, 20.0);
  EXPECT_EQ(range.end_num, 300.0);
  EXPECT_EQ(range.interval, 100.0);
  EXPECT_TRUE(range.auto_start);
  EXPECT_TRUE(range.auto_end);
  EXPECT_EQ(GroupTexts(group), (std::vector<std::string>{"<20", "20-119", "120-219", "220-319", ">320"}));
  EXPECT_EQ(GroupNumbers(cache), (std::vector<double>{1, 1, 2, 3}));

  const PivotTable& grouped = *wb.sheet(0).pivot_tables()[0];
  const PivotTable& sibling = *wb.sheet(0).pivot_tables()[1];
  for (const PivotTable* table : {&grouped, &sibling}) {
    ASSERT_EQ(table->fields().size(), 4U);
    const PivotField& field = table->fields()[kGroup];
    EXPECT_EQ(field.source_name, "Amt2");
    EXPECT_TRUE(field.sort.manual);
    ASSERT_EQ(field.items.size(), 5U);
    for (std::uint32_t x = 0; x < 5U; ++x) {
      EXPECT_EQ(field.items[x].cache_index, x);
    }
    EXPECT_TRUE(table->layout_dirty());
  }
  EXPECT_EQ(grouped.row_field_order(), (std::vector<std::uint32_t>{3U, 0U}));
  EXPECT_EQ(grouped.fields()[kAmt].axis, PivotAxis::None);
  EXPECT_EQ(sibling.row_field_order(), (std::vector<std::uint32_t>{3U, 1U, 0U}));
  EXPECT_EQ(sibling.fields()[kAmt].axis, PivotAxis::Row);
}

TEST(NumericGroup, NegativeStartLabelsCountFromIt) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({-30, 400, 950}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  EXPECT_EQ(GroupTexts(CacheOf(wb).fields()[kGroup]),
            (std::vector<std::string>{"<-30", "-30-69", "70-169", "170-269", "270-369", "370-469", "470-569", "570-669",
                                      "670-769", "770-869", "870-969", ">970"}));
  EXPECT_EQ(GroupNumbers(CacheOf(wb)), (std::vector<double>{1, 5, 10}));
}

TEST(NumericGroup, FixedBoundsPutOutliersInTheEndItems) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({-5, 50, 150, 200, 250}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, 0.0, 200.0, 100.0)));
  const PivotCache& cache = CacheOf(wb);
  const RangeGroup& range = cache.fields()[kGroup].field_group->range;
  EXPECT_FALSE(range.auto_start);
  EXPECT_FALSE(range.auto_end);
  // The end on a range edge closes the last range, which then reads up to it.
  EXPECT_EQ(GroupTexts(cache.fields()[kGroup]), (std::vector<std::string>{"<0", "0-99", "100-200", ">200"}));
  EXPECT_EQ(GroupNumbers(cache), (std::vector<double>{0, 1, 2, 2, 3}));
}

TEST(NumericGroup, AnEndOnARangeEdgeClosesTheLastIntegerRange) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  // Mac Excel: 10..120 by 10 ends with `110-120` and `>120`; by 50 with `110-159` and `>160`.
  Workbook by10 = AmountWorkbook(Numbers({10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(by10, 0U, 0U, kAmt, std::nullopt, std::nullopt, 10.0)));
  EXPECT_EQ(GroupTexts(CacheOf(by10).fields()[kGroup]),
            (std::vector<std::string>{"<10", "10-19", "20-29", "30-39", "40-49", "50-59", "60-69", "70-79", "80-89",
                                      "90-99", "100-109", "110-120", ">120"}));
  EXPECT_EQ(RecordLabels(CacheOf(by10)).back(), "110-120");
  Workbook by50 = AmountWorkbook(Numbers({10, 120}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(by50, 0U, 0U, kAmt, std::nullopt, std::nullopt, 50.0)));
  EXPECT_EQ(GroupTexts(CacheOf(by50).fields()[kGroup]),
            (std::vector<std::string>{"<10", "10-59", "60-109", "110-159", ">160"}));
  // A start of 1 leaves the end inside the last range: `111-120`, `>121`.
  Workbook from1 = AmountWorkbook(Numbers({10, 120}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(from1, 0U, 0U, kAmt, 1.0, std::nullopt, 10.0)));
  const std::vector<std::string> texts = GroupTexts(CacheOf(from1).fields()[kGroup]);
  ASSERT_EQ(texts.size(), 14U);
  EXPECT_EQ(texts[1], "1-10");
  EXPECT_EQ(texts[12], "111-120");
  EXPECT_EQ(texts[13], ">121");
}

TEST(NumericGroup, RegroupingReplacesTheGroupField) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  // Through the group field, as on a group cell.
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kGroup, std::nullopt, std::nullopt, 50.0)));
  const PivotCache& cache = CacheOf(wb);
  ASSERT_EQ(cache.fields().size(), 4U);
  EXPECT_EQ(cache.fields()[kGroup].name, "Amt2");
  EXPECT_EQ(cache.fields()[kGroup].field_group->group_items.size(), 8U);
  const PivotTable& table = *wb.sheet(0).pivot_tables()[0];
  EXPECT_EQ(table.row_field_order(), (std::vector<std::uint32_t>{3U, 0U}));
  EXPECT_EQ(table.fields()[kGroup].items.size(), 8U);
}

TEST(NumericGroup, ClearRestoresTheBaseWhereItsGroupWas) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}), 2U);
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  ASSERT_TRUE(static_cast<bool>(clear_numeric_group(wb, 0U, 1U, kGroup)));
  const PivotCache& cache = CacheOf(wb);
  ASSERT_EQ(cache.fields().size(), 3U);
  EXPECT_FALSE(cache.fields()[kAmt].field_group.has_value());
  for (const auto& table : wb.sheet(0).pivot_tables()) {
    ASSERT_EQ(table->fields().size(), 3U);
    EXPECT_EQ(table->row_field_order(), (std::vector<std::uint32_t>{1U, 0U}));
    EXPECT_EQ(table->fields()[kAmt].axis, PivotAxis::Row);
  }
  // Clearing what is not grouped changes nothing; an index out of range fails.
  EXPECT_TRUE(static_cast<bool>(clear_numeric_group(wb, 0U, 0U, kAmt)));
  EXPECT_FALSE(static_cast<bool>(clear_numeric_group(wb, 0U, 0U, 9U)));
}

// Mac Excel measurements over base A (0.3 .. 3.6 by 0.3), base B (-2.5 .. 2.5
// by 0.5) and base A scaled by 1e-5 and by 1e15.
const std::vector<double> kBaseA = {0.3, 0.6, 0.9, 1.2, 1.5, 1.8, 2.1, 2.4, 2.7, 3.0, 3.3, 3.6};
const std::vector<double> kBaseB = {-2.5, -2.0, -1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5, 2.0, 2.5};
const std::vector<double> kBaseATimes1e5 = {3e-6,   6e-6,   9e-6,   1.2e-5, 1.5e-5, 1.8e-5,
                                            2.1e-5, 2.4e-5, 2.7e-5, 3e-5,   3.3e-5, 3.6e-5};
const std::vector<double> kBaseATimes1e15 = {3e14,   6e14,   9e14,   1.2e15, 1.5e15, 1.8e15,
                                             2.1e15, 2.4e15, 2.7e15, 3e15,   3.3e15, 3.6e15};

struct FractionalCase {
  const char* name;
  const std::vector<double>& values;
  std::optional<double> start;
  std::optional<double> end;
  double interval;
  std::vector<std::string> items;
  std::vector<std::string> record_labels;
};

std::vector<Value> ValuesOf(const std::vector<double>& values) {
  std::vector<Value> out;
  for (const double v : values) {
    out.push_back(Value::number(v));
  }
  return out;
}

TEST(NumericGroup, FractionalGroupingMatchesExcel) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const std::vector<FractionalCase> cases = {
      {"c01",
       kBaseA,
       std::nullopt,
       std::nullopt,
       0.5,
       {"<0.3", "0.3-0.8", "0.8-1.3", "1.3-1.8", "1.8-2.3", "2.3-2.8", "2.8-3.3", "3.3-3.8", ">3.8"},
       {"0.3-0.8", "0.3-0.8", "0.8-1.3", "0.8-1.3", "1.3-1.8", "1.8-2.3", "1.8-2.3", "2.3-2.8", "2.3-2.8", "2.8-3.3",
        "3.3-3.8", "3.3-3.8"}},
      {"c02",
       kBaseA,
       std::nullopt,
       std::nullopt,
       0.25,
       {"<0.3", "0.3-0.55", "0.55-0.8", "0.8-1.05", "1.05-1.3", "1.3-1.55", "1.55-1.8", "1.8-2.05", "2.05-2.3",
        "2.3-2.55", "2.55-2.8", "2.8-3.05", "3.05-3.3", "3.3-3.55", "3.55-3.8", ">3.8"},
       {"0.3-0.55", "0.55-0.8", "0.8-1.05", "1.05-1.3", "1.3-1.55", "1.8-2.05", "2.05-2.3", "2.3-2.55", "2.55-2.8",
        "2.8-3.05", "3.3-3.55", "3.55-3.8"}},
      {"c03",
       kBaseA,
       std::nullopt,
       std::nullopt,
       1.5,
       {"<0.3", "0.3-1.8", "1.8-3.3", "3.3-4.8", ">4.8"},
       {"0.3-1.8", "0.3-1.8", "0.3-1.8", "0.3-1.8", "0.3-1.8", "1.8-3.3", "1.8-3.3", "1.8-3.3", "1.8-3.3", "1.8-3.3",
        "3.3-4.8", "3.3-4.8"}},
      {"c04",
       kBaseA,
       0.25,
       3.75,
       0.5,
       {"<0.25", "0.25-0.75", "0.75-1.25", "1.25-1.75", "1.75-2.25", "2.25-2.75", "2.75-3.25", "3.25-3.75", ">3.75"},
       {"0.25-0.75", "0.25-0.75", "0.75-1.25", "0.75-1.25", "1.25-1.75", "1.75-2.25", "1.75-2.25", "2.25-2.75",
        "2.25-2.75", "2.75-3.25", "3.25-3.75", "3.25-3.75"}},
      {"c05",
       kBaseA,
       0.0,
       4.0,
       1.5,
       {"<0", "0-1.5", "1.5-3", "3-4.5", ">4.5"},
       {"0-1.5", "0-1.5", "0-1.5", "0-1.5", "1.5-3", "1.5-3", "1.5-3", "1.5-3", "1.5-3", "3-4.5", "3-4.5", "3-4.5"}},
      {"c06",
       kBaseB,
       std::nullopt,
       std::nullopt,
       0.5,
       {"<-2.5", "-2.5--2", "-2--1.5", "-1.5--1", "-1--0.5", "-0.5-0", "0-0.5", "0.5-1", "1-1.5", "1.5-2", "2-2.5",
        ">2.5"},
       {"-2.5--2", "-2--1.5", "-1.5--1", "-1--0.5", "-0.5-0", "0-0.5", "0.5-1", "1-1.5", "1.5-2", "2-2.5", "2-2.5"}},
      {"c07",
       kBaseB,
       -2.25,
       2.25,
       0.75,
       {"<-2.25", "-2.25--1.5", "-1.5--0.75", "-0.75-0", "0-0.75", "0.75-1.5", "1.5-2.25", ">2.25"},
       {"<-2.25", "-2.25--1.5", "-1.5--0.75", "-1.5--0.75", "-0.75-0", "0-0.75", "0-0.75", "0.75-1.5", "1.5-2.25",
        "1.5-2.25", ">2.25"}},
      {"c08",
       kBaseA,
       std::nullopt,
       std::nullopt,
       0.1,
       {"<0.3",    "0.3-0.4", "0.4-0.5", "0.5-0.6", "0.6-0.7", "0.7-0.8", "0.8-0.9", "0.9-1",   "1-1.1",
        "1.1-1.2", "1.2-1.3", "1.3-1.4", "1.4-1.5", "1.5-1.6", "1.6-1.7", "1.7-1.8", "1.8-1.9", "1.9-2",
        "2-2.1",   "2.1-2.2", "2.2-2.3", "2.3-2.4", "2.4-2.5", "2.5-2.6", "2.6-2.7", "2.7-2.8", "2.8-2.9",
        "2.9-3",   "3-3.1",   "3.1-3.2", "3.2-3.3", "3.3-3.4", "3.4-3.5", "3.5-3.6", ">3.6"},
       {"0.3-0.4", "0.6-0.7", "0.9-1", "1.2-1.3", "1.5-1.6", "1.8-1.9", "2.1-2.2", "2.4-2.5", "2.7-2.8", "3-3.1",
        "3.3-3.4", "3.5-3.6"}},
      {"c09",
       kBaseA,
       1.0,
       3.0,
       0.5,
       {"<1", "1-1.5", "1.5-2", "2-2.5", "2.5-3", ">3"},
       {"<1", "<1", "<1", "1-1.5", "1.5-2", "1.5-2", "2-2.5", "2-2.5", "2.5-3", "2.5-3", ">3", ">3"}},
      {"corner a",
       kBaseA,
       0.3,
       3.4,
       0.5,
       {"<0.3", "0.3-0.8", "0.8-1.3", "1.3-1.8", "1.8-2.3", "2.3-2.8", "2.8-3.3", "3.3-3.8", ">3.8"},
       {"0.3-0.8", "0.3-0.8", "0.8-1.3", "0.8-1.3", "1.3-1.8", "1.8-2.3", "1.8-2.3", "2.3-2.8", "2.3-2.8", "2.8-3.3",
        "3.3-3.8", ">3.8"}},
      {"corner b",
       kBaseA,
       1.0,
       1.0,
       0.5,
       {"<1", "1-1.5", ">1.5"},
       {"<1", "<1", "<1", ">1.5", ">1.5", ">1.5", ">1.5", ">1.5", ">1.5", ">1.5", ">1.5", ">1.5"}},
      {"corner c",
       kBaseATimes1e5,
       std::nullopt,
       std::nullopt,
       1e-05,
       {"<0.000003", "0.000003-0.000013", "0.000013-0.000023", "0.000023-0.000033", "0.000033-0.000043", ">0.000043"},
       {"0.000003-0.000013", "0.000003-0.000013", "0.000003-0.000013", "0.000003-0.000013", "0.000013-0.000023",
        "0.000013-0.000023", "0.000013-0.000023", "0.000023-0.000033", "0.000023-0.000033", "0.000023-0.000033",
        "0.000033-0.000043", "0.000033-0.000043"}},
      {"corner d",
       kBaseATimes1e15,
       std::nullopt,
       std::nullopt,
       1000000000000000.0,
       {"<300000000000000", "300000000000000-1300000000000000", "1300000000000000-2300000000000000",
        "2300000000000000-3300000000000000", "3300000000000000-4300000000000000", ">4300000000000000"},
       {"300000000000000-1300000000000000", "300000000000000-1300000000000000", "300000000000000-1300000000000000",
        "300000000000000-1300000000000000", "1300000000000000-2300000000000000", "1300000000000000-2300000000000000",
        "1300000000000000-2300000000000000", "2300000000000000-3300000000000000", "2300000000000000-3300000000000000",
        "2300000000000000-3300000000000000", "3300000000000000-4300000000000000", "3300000000000000-4300000000000000"}},
  };
  for (const FractionalCase& c : cases) {
    Workbook wb = AmountWorkbook(ValuesOf(c.values));
    const Expected<void, Error> grouped = set_numeric_group(wb, 0U, 0U, kAmt, c.start, c.end, c.interval);
    ASSERT_TRUE(static_cast<bool>(grouped)) << c.name << ": " << grouped.error().message;
    const PivotCache& cache = CacheOf(wb);
    EXPECT_EQ(GroupTexts(cache.fields()[kGroup]), c.items) << c.name;
    EXPECT_EQ(RecordLabels(cache), c.record_labels) << c.name;
    const RangeGroup& range = cache.fields()[kGroup].field_group->range;
    EXPECT_EQ(range.start_num, c.start.value_or(c.values.front())) << c.name;
    EXPECT_EQ(range.end_num, c.end.value_or(c.values.back())) << c.name;
    EXPECT_EQ(range.interval, c.interval) << c.name;
    EXPECT_EQ(range.auto_start, !c.start.has_value()) << c.name;
    EXPECT_EQ(range.auto_end, !c.end.has_value()) << c.name;
    EXPECT_EQ(wb.sheet(0).pivot_tables()[0]->fields()[kGroup].items.size(), c.items.size()) << c.name;
  }
}

TEST(NumericGroup, LabelFormAroundTheIntegerAndExponentThresholds) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  std::vector<double> e14;
  std::vector<double> e15;
  std::vector<double> e300;
  for (const double v : kBaseA) {
    e14.push_back(v * 1e14);
    e15.push_back(v * 1e15);
    e300.push_back(v * 1e300);
  }
  // Whole settings at 1e14 keep the integer form; the last range ends at the end on its edge.
  Workbook integer = AmountWorkbook(ValuesOf(e14));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(integer, 0U, 0U, kAmt, 1e14, 5e14, 1e14)));
  EXPECT_EQ(GroupTexts(CacheOf(integer).fields()[kGroup]),
            (std::vector<std::string>{"<100000000000000", "100000000000000-199999999999999",
                                      "200000000000000-299999999999999", "300000000000000-399999999999999",
                                      "400000000000000-500000000000000", ">500000000000000"}));
  // At 1e15 they take the decimal form.
  Workbook decimal = AmountWorkbook(ValuesOf(e15));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(decimal, 0U, 0U, kAmt, 1e15, 5e15, 1e15)));
  EXPECT_EQ(GroupTexts(CacheOf(decimal).fields()[kGroup]),
            (std::vector<std::string>{"<1000000000000000", "1000000000000000-2000000000000000",
                                      "2000000000000000-3000000000000000", "3000000000000000-4000000000000000",
                                      "4000000000000000-5000000000000000", ">5000000000000000"}));
  EXPECT_EQ(
      RecordLabels(CacheOf(decimal)),
      (std::vector<std::string>{
          "<1000000000000000", "<1000000000000000", "<1000000000000000", "1000000000000000-2000000000000000",
          "1000000000000000-2000000000000000", "1000000000000000-2000000000000000", "2000000000000000-3000000000000000",
          "2000000000000000-3000000000000000", "2000000000000000-3000000000000000", "3000000000000000-4000000000000000",
          "3000000000000000-4000000000000000", "3000000000000000-4000000000000000"}));
  Workbook wide = AmountWorkbook(ValuesOf(e15));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wide, 0U, 0U, kAmt, 2e15, 6e15, 2e15)));
  EXPECT_EQ(GroupTexts(CacheOf(wide).fields()[kGroup]),
            (std::vector<std::string>{"<2000000000000000", "2000000000000000-4000000000000000",
                                      "4000000000000000-6000000000000000", ">6000000000000000"}));
  // At 1e300 the edges take the exponent form.
  Workbook huge = AmountWorkbook(ValuesOf(e300));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(huge, 0U, 0U, kAmt, std::nullopt, std::nullopt, 1e300)));
  EXPECT_EQ(GroupTexts(CacheOf(huge).fields()[kGroup]),
            (std::vector<std::string>{"<3E+299", "3E+299-1.3E+300", "1.3E+300-2.3E+300", "2.3E+300-3.3E+300",
                                      "3.3E+300-4.3E+300", ">4.3E+300"}));
  EXPECT_EQ(GroupNumbers(CacheOf(huge)), (std::vector<double>{1, 1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4}));
}

// Groups `values` over the given bounds and returns the group items.
std::vector<std::string> GroupItems(const std::vector<double>& values, std::optional<double> start,
                                    std::optional<double> end, double interval) {
  Workbook wb = AmountWorkbook(ValuesOf(values));
  const Expected<void, Error> grouped = set_numeric_group(wb, 0U, 0U, kAmt, start, end, interval);
  EXPECT_TRUE(static_cast<bool>(grouped)) << (grouped ? std::string() : grouped.error().message);
  return grouped ? GroupTexts(CacheOf(wb).fields()[kGroup]) : std::vector<std::string>{};
}

std::vector<double> Scaled(double scale, double offset = 0.0) {
  std::vector<double> out;
  for (const double v : kBaseA) {
    out.push_back(offset + v * scale);
  }
  return out;
}

TEST(NumericGroup, TheIntegerFormIsChosenPerRange) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  EXPECT_EQ(GroupItems(Scaled(1e13, 9e14), 9e14, 9.3e14, 1e13),
            (std::vector<std::string>{"<900000000000000", "900000000000000-909999999999999",
                                      "910000000000000-919999999999999", "920000000000000-930000000000000",
                                      ">930000000000000"}));
  EXPECT_EQ(GroupItems(Scaled(100.0, 999999999999000.0), 999999999999000.0, 999999999999300.0, 100.0),
            (std::vector<std::string>{"<999999999999000", "999999999999000-999999999999099",
                                      "999999999999100-999999999999199", "999999999999200-999999999999300",
                                      ">999999999999300"}));
  EXPECT_EQ(GroupItems(Scaled(100.0, 1e15), 1e15, 1e15 + 300.0, 100.0),
            (std::vector<std::string>{"<1000000000000000", "1000000000000000-1000000000000100",
                                      "1000000000000100-1000000000000200", "1000000000000200-1000000000000300",
                                      ">1000000000000300"}));
  // The forms mix within one field.
  EXPECT_EQ(GroupItems(Scaled(100.0, 999999999999000.0), 999999999999800.0, 1e15 + 100.0, 100.0),
            (std::vector<std::string>{"<999999999999800", "999999999999800-999999999999899",
                                      "999999999999900-999999999999999", "1000000000000000-1000000000000100",
                                      ">1000000000000100"}));
  // A range starting below 1e15 whose upper edge passes it takes the decimal form.
  EXPECT_EQ(GroupItems(Scaled(1e15), 9.5e14, 1.15e15, 1e14),
            (std::vector<std::string>{"<950000000000000", "950000000000000-1050000000000000",
                                      "1050000000000000-1150000000000000", ">1150000000000000"}));
}

TEST(NumericGroup, EachLabelNumberIsSpelledInFullUpToTwentyCharacters) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  EXPECT_EQ(GroupItems(Scaled(1e16), std::nullopt, std::nullopt, 1e16),
            (std::vector<std::string>{"<3000000000000000", "3000000000000000-13000000000000000",
                                      "13000000000000000-23000000000000000", "23000000000000000-33000000000000000",
                                      "33000000000000000-43000000000000000", ">43000000000000000"}));
  EXPECT_EQ(
      GroupItems({9.5e19, 1.05e20}, 9e19, 1.1e20, 1e19),
      (std::vector<std::string>{"<90000000000000000000", "90000000000000000000-1E+20", "1E+20-1.1E+20", ">1.1E+20"}));
  const std::vector<std::string> e20 = GroupItems(Scaled(1e20), std::nullopt, std::nullopt, 1e20);
  ASSERT_EQ(e20.size(), 6U);
  EXPECT_EQ(e20[1], "30000000000000000000-1.3E+20");
  EXPECT_EQ(GroupItems(Scaled(1e21), std::nullopt, std::nullopt, 1e21)[1], "3E+20-1.3E+21");
  EXPECT_EQ(GroupItems(Scaled(1e-9), std::nullopt, std::nullopt, 1e-9)[1], "0.0000000003-0.0000000013");
  EXPECT_EQ(GroupItems(Scaled(1e-17), std::nullopt, std::nullopt, 1e-17),
            (std::vector<std::string>{"<0.000000000000000003", "0.000000000000000003-0.000000000000000013",
                                      "0.000000000000000013-0.000000000000000023",
                                      "0.000000000000000023-0.000000000000000033",
                                      "0.000000000000000033-0.000000000000000043", ">0.000000000000000043"}));
  EXPECT_EQ(GroupItems(Scaled(1e-18), std::nullopt, std::nullopt, 1e-18),
            (std::vector<std::string>{"<3E-19", "3E-19-1.3E-18", "1.3E-18-2.3E-18", "2.3E-18-3.3E-18",
                                      "3.3E-18-4.3E-18", ">4.3E-18"}));
  EXPECT_EQ(GroupItems(Scaled(1e-20), std::nullopt, std::nullopt, 1e-20).front(), "<3E-21");
  // A minus sign does not count toward the 20 characters.
  EXPECT_EQ(GroupItems({-5e18}, -1e19, 0.0, 1e19),
            (std::vector<std::string>{"<-10000000000000000000", "-10000000000000000000-0", ">0"}));
  EXPECT_EQ(GroupItems({-1e-17}, -1.2e-17, 0.0, 1.23e-17),
            (std::vector<std::string>{"<-0.000000000000000012", "-0.000000000000000012-3E-19", ">3E-19"}));
  // The forms mix within one label: 0.0000000000000000243 would be 21 characters.
  EXPECT_EQ(GroupItems({1.5e-17}, 1.2e-17, 2e-17, 1.23e-17),
            (std::vector<std::string>{"<0.000000000000000012", "0.000000000000000012-2.43E-17", ">2.43E-17"}));
}

TEST(NumericGroup, LabelsUseTheWorkbookLocaleDecimalSeparator) {
  for (const ExcelLocale locale : {ExcelLocale::kDeDE, ExcelLocale::kFrFR}) {
    const ExcelProfile profile{ExcelHost::kMac365, locale};
    const eval::ScopedEvalProfile scope(profile);
    Workbook a = AmountWorkbook(ValuesOf(kBaseA), 1U, profile);
    ASSERT_TRUE(static_cast<bool>(set_numeric_group(a, 0U, 0U, kAmt, std::nullopt, std::nullopt, 0.5)));
    EXPECT_EQ(GroupTexts(CacheOf(a).fields()[kGroup]),
              (std::vector<std::string>{"<0,3", "0,3-0,8", "0,8-1,3", "1,3-1,8", "1,8-2,3", "2,3-2,8", "2,8-3,3",
                                        "3,3-3,8", ">3,8"}));
    // The stored bounds stay numbers.
    EXPECT_EQ(CacheOf(a).fields()[kGroup].field_group->range.start_num, 0.3);
    Workbook b = AmountWorkbook(ValuesOf(kBaseB), 1U, profile);
    ASSERT_TRUE(static_cast<bool>(set_numeric_group(b, 0U, 0U, kAmt, -2.25, 2.25, 0.75)));
    EXPECT_EQ(GroupTexts(CacheOf(b).fields()[kGroup]),
              (std::vector<std::string>{"<-2,25", "-2,25--1,5", "-1,5--0,75", "-0,75-0", "0-0,75", "0,75-1,5",
                                        "1,5-2,25", ">2,25"}));
    // The integer form has no decimal point to localize.
    Workbook whole = AmountWorkbook(Numbers({20, 300}), 1U, profile);
    ASSERT_TRUE(static_cast<bool>(set_numeric_group(whole, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
    EXPECT_EQ(GroupTexts(CacheOf(whole).fields()[kGroup]),
              (std::vector<std::string>{"<20", "20-119", "120-219", "220-319", ">320"}));
  }

  // German, measured: the exponent form localizes its mantissa too, and the
  // 20-character rule counts the comma like the point.
  const ExcelProfile de{ExcelHost::kMac365, ExcelLocale::kDeDE};
  const eval::ScopedEvalProfile scope(de);
  const auto german = [&de](double scale) {
    Workbook wb = AmountWorkbook(ValuesOf(Scaled(scale)), 1U, de);
    EXPECT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, scale)));
    return GroupTexts(CacheOf(wb).fields()[kGroup]);
  };
  EXPECT_EQ(german(1e300), (std::vector<std::string>{"<3E+299", "3E+299-1,3E+300", "1,3E+300-2,3E+300",
                                                     "2,3E+300-3,3E+300", "3,3E+300-4,3E+300", ">4,3E+300"}));
  EXPECT_EQ(german(1e-20), (std::vector<std::string>{"<3E-21", "3E-21-1,3E-20", "1,3E-20-2,3E-20", "2,3E-20-3,3E-20",
                                                     "3,3E-20-4,3E-20", ">4,3E-20"}));
  EXPECT_EQ(german(1e-17), (std::vector<std::string>{
                               "<0,000000000000000003", "0,000000000000000003-0,000000000000000013",
                               "0,000000000000000013-0,000000000000000023", "0,000000000000000023-0,000000000000000033",
                               "0,000000000000000033-0,000000000000000043", ">0,000000000000000043"}));
}

TEST(NumericGroup, WholeSettingsBelow1e15KeepTheIntegerForm) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  // A whole start and end with a fractional interval take the decimal form.
  Workbook half = AmountWorkbook(Numbers({1, 2}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(half, 0U, 0U, kAmt, std::nullopt, std::nullopt, 0.5)));
  EXPECT_EQ(GroupTexts(CacheOf(half).fields()[kGroup]), (std::vector<std::string>{"<1", "1-1.5", "1.5-2", ">2"}));
  // A fractional start with a whole interval takes it too.
  Workbook offset = AmountWorkbook(Numbers({20.5, 300}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(offset, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  EXPECT_EQ(GroupTexts(CacheOf(offset).fields()[kGroup]),
            (std::vector<std::string>{"<20.5", "20.5-120.5", "120.5-220.5", "220.5-320.5", ">320.5"}));
  EXPECT_EQ(GroupNumbers(CacheOf(offset)), (std::vector<double>{1, 3}));
}

TEST(NumericGroup, TooManyItemsAreRefusedAndChangeNothing) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({0, 1}));
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 0.5)));
  const std::vector<std::string> before = GroupTexts(CacheOf(wb).fields()[kGroup]);
  for (const Expected<void, Error>& result : {
           set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 1e-300),
           set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 5e-324),
           set_numeric_group(wb, 0U, 0U, kAmt, -1.7e308, 1.7e308, 1.0),
           set_numeric_group(wb, 0U, 0U, kAmt, 0.0, 1048576.0, 1.0),
       }) {
    ASSERT_FALSE(static_cast<bool>(result));
    EXPECT_EQ(result.error().code, FormulonErrorCode::kSecurityResourceLimit);
  }
  ASSERT_EQ(CacheOf(wb).fields().size(), 4U);
  EXPECT_EQ(GroupTexts(CacheOf(wb).fields()[kGroup]), before);
}

TEST(NumericGroup, InvalidSettingsAreRefused) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  for (const Expected<void, Error>& result : {
           set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, 0.0),
           set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, std::nullopt, -100.0),
           set_numeric_group(wb, 0U, 0U, kAmt, 300.0, 20.0, 100.0),
           // A setting is refused before the base data is looked at.
           set_numeric_group(wb, 0U, 0U, kRegion, std::nullopt, std::nullopt, 0.0),
       }) {
    ASSERT_FALSE(static_cast<bool>(result));
    EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument);
  }
  EXPECT_EQ(set_numeric_group(wb, 1U, 0U, kAmt, std::nullopt, std::nullopt, 100.0).error().code,
            FormulonErrorCode::kSheetIndexOutOfRange);
  EXPECT_EQ(set_numeric_group(wb, 0U, 3U, kAmt, std::nullopt, std::nullopt, 100.0).error().code,
            FormulonErrorCode::kIndexOutOfRange);
  EXPECT_EQ(CacheOf(wb).fields().size(), 3U);
}

TEST(NumericGroup, BaseDataThatIsNotNumbersIsAPrecondition) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  const Expected<void, Error> text = set_numeric_group(wb, 0U, 0U, kRegion, std::nullopt, std::nullopt, 100.0);
  ASSERT_FALSE(static_cast<bool>(text));
  EXPECT_EQ(text.error().code, FormulonErrorCode::kPreconditionFailed);
  Workbook blank = AmountWorkbook({Value::number(20), Value::blank()});
  const Expected<void, Error> with_blank = set_numeric_group(blank, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0);
  ASSERT_FALSE(static_cast<bool>(with_blank));
  EXPECT_EQ(with_blank.error().code, FormulonErrorCode::kPreconditionFailed);
  Workbook empty = AmountWorkbook({});
  const Expected<void, Error> no_number = set_numeric_group(empty, 0U, 0U, kAmt, std::nullopt, 10.0, 1.0);
  ASSERT_FALSE(static_cast<bool>(no_number));
  EXPECT_EQ(no_number.error().code, FormulonErrorCode::kPreconditionFailed);
  for (const Workbook* book : {&wb, &blank, &empty}) {
    EXPECT_EQ(CacheOf(*book).fields().size(), 3U);
  }
  // With both bounds given, an empty field groups.
  EXPECT_TRUE(static_cast<bool>(set_numeric_group(empty, 0U, 0U, kAmt, 0.0, 10.0, 1.0)));

  // A missing shared item does not count as a present value.
  Workbook stale = AmountWorkbook({});
  PivotCacheField& amt = (*stale.mutable_pivot_caches().front()).mutable_fields()[kAmt];
  amt.shared_items = {Value::number(10.0), Value::text("gone")};
  amt.shared_item_is_missing = {false, true};
  EXPECT_TRUE(static_cast<bool>(set_numeric_group(stale, 0U, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
}

TEST(NumericGroup, StateReadsTheGroupingBack) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  Workbook wb = AmountWorkbook(Numbers({20, 300}));
  EXPECT_FALSE(numeric_group_state(CacheOf(wb), kAmt).has_value());
  ASSERT_TRUE(static_cast<bool>(set_numeric_group(wb, 0U, 0U, kAmt, std::nullopt, 400.0, 100.0)));
  for (const std::size_t field : {kAmt, kGroup}) {
    const std::optional<NumericGroupState> state = numeric_group_state(CacheOf(wb), field);
    ASSERT_TRUE(state.has_value());
    EXPECT_EQ(state->start, 20.0);
    EXPECT_EQ(state->end, 400.0);
    EXPECT_EQ(state->interval, 100.0);
    EXPECT_TRUE(state->auto_start);
    EXPECT_FALSE(state->auto_end);
  }
  EXPECT_FALSE(numeric_group_state(CacheOf(wb), kRegion).has_value());
  EXPECT_FALSE(date_group_state(CacheOf(wb), kAmt, /*date1904=*/false).has_value());
}

double Day(int y, unsigned m, unsigned d) {
  return date_time::serial_from_ymd(y, m, d);
}

PivotCache DateCache() {
  PivotCache cache;
  cache.set_cache_id(1U);
  cache.mutable_fields().push_back(PivotCacheField{"Date", {}});
  for (const double serial : {Day(2024, 1U, 15U), Day(2024, 3U, 10U)}) {
    PivotCacheRecord record;
    record.cells = {Value::number(serial)};
    cache.mutable_records().push_back(std::move(record));
  }
  return cache;
}

TEST(NumericGroup, DateStateReadsTheDateGroupingBack) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache cache = DateCache();
  EXPECT_FALSE(date_group_state(cache, 0U, /*date1904=*/false).has_value());
  DateGroupSpec spec;
  spec.periods = date_period_bit(GroupBy::Months) | date_period_bit(GroupBy::Years);
  spec.end = Day(2024, 12U, 31U);
  ASSERT_TRUE(static_cast<bool>(group_cache_by_date(cache, 0U, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  for (const std::size_t field : {0U, 1U, 2U}) {
    const std::optional<DateGroupState> state = date_group_state(cache, field, /*date1904=*/false);
    ASSERT_TRUE(state.has_value());
    EXPECT_EQ(state->periods, spec.periods);
    EXPECT_EQ(state->calendar, CalendarSystem::Gregorian);
    EXPECT_EQ(state->interval_days, 1U);
    EXPECT_EQ(state->start, Day(2024, 1U, 15U));
    EXPECT_EQ(state->end, Day(2024, 12U, 31U));
    EXPECT_TRUE(state->auto_start);
    EXPECT_FALSE(state->auto_end);
  }
  EXPECT_FALSE(numeric_group_state(cache, 0U).has_value());

  PivotCache days = DateCache();
  spec = DateGroupSpec{};
  spec.periods = date_period_bit(GroupBy::Days);
  spec.interval_days = 7U;
  ASSERT_TRUE(static_cast<bool>(group_cache_by_date(days, 0U, spec, /*date1904=*/false, mac_365_ja_jp_profile())));
  const std::optional<DateGroupState> state = date_group_state(days, 0U, /*date1904=*/false);
  ASSERT_TRUE(state.has_value());
  EXPECT_EQ(state->interval_days, 7U);
  EXPECT_EQ(state->end, Day(2024, 3U, 11U));
  EXPECT_TRUE(state->auto_end);
}

}  // namespace
}  // namespace formulon::pivot
