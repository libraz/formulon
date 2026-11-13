//
// Workbooks Mac Excel 365 (ja-JP) saved after grouping a pivot's date field
// by months, quarters, years and days, after ungrouping it again, and after
// renaming an item. Each pivot is evaluated as loaded, and its row labels and
// values must equal what Excel drew: one entry per rendered row, a parent
// row carrying its subtotal, then the grand total.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <initializer_list>
#include <limits>
#include <memory>
#include <optional>
#include <ostream>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const char* file) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_options/" + file;
  std::vector<std::uint8_t> out;
  FILE* handle = std::fopen(path.c_str(), "rb");
  if (handle == nullptr) {
    return out;
  }
  std::fseek(handle, 0, SEEK_END);
  const long size = std::ftell(handle);
  std::fseek(handle, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), handle) != out.size()) {
      out.clear();
    }
  }
  std::fclose(handle);
  return out;
}

/// One rendered row: its label and one value per data field.
struct Row {
  Row(std::string row_label, std::vector<double> row_values)
      : label(std::move(row_label)), values(std::move(row_values)) {}
  // From a range: GCC 14 -Werror=stringop-overflow misjudges a vector built from a one-element braced list.
  Row(std::string row_label, std::initializer_list<double> row_values)
      : label(std::move(row_label)), values(row_values.begin(), row_values.end()) {}

  std::string label;
  std::vector<double> values;

  bool operator==(const Row& other) const { return label == other.label && values == other.values; }
};

std::ostream& operator<<(std::ostream& os, const Row& row) {
  os << row.label;
  for (const double v : row.values) {
    os << " | " << v;
  }
  return os;
}

std::vector<double> Numbers(const std::vector<Value>& values) {
  std::vector<double> out;
  for (const Value& v : values) {
    out.push_back(v.is_number() ? v.as_number() : std::numeric_limits<double>::quiet_NaN());
  }
  return out;
}

void Flatten(const pivot::PivotResult& result, const std::vector<pivot::AxisHierarchyNode>& nodes,
             std::vector<std::string>& path, std::size_t& leaf, std::vector<Row>& out) {
  for (const pivot::AxisHierarchyNode& node : nodes) {
    path.push_back(node.label);
    if (node.children.empty()) {
      out.push_back({node.label, Numbers(result.values[leaf][0])});
      ++leaf;
    } else {
      std::vector<double> subtotal;
      for (const pivot::RowSubtotal& candidate : result.row_subtotals) {
        if (candidate.labels == path) {
          subtotal = Numbers(candidate.values);
          break;
        }
      }
      out.push_back({node.label, std::move(subtotal)});
      Flatten(result, node.children, path, leaf, out);
    }
    path.pop_back();
  }
}

/// The rows the pivot named `table_name` on `sheet_name` evaluates to under
/// the ja-JP profile the files were saved in.
std::vector<Row> EvaluatedRows(const char* file, const char* sheet_name, const char* table_name) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file);
  EXPECT_FALSE(bytes.empty()) << "could not read " << file;
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  if (!read_or) {
    return {};
  }
  const Workbook& wb = read_or.value().workbook;
  const Sheet* sheet = wb.sheet_by_name(sheet_name);
  EXPECT_NE(sheet, nullptr) << sheet_name;
  if (sheet == nullptr) {
    return {};
  }
  const pivot::PivotTable* table = nullptr;
  for (const std::unique_ptr<pivot::PivotTable>& candidate : sheet->pivot_tables()) {
    if (candidate != nullptr && candidate->name() == table_name) {
      table = candidate.get();
    }
  }
  EXPECT_NE(table, nullptr) << table_name;
  const pivot::PivotCache* cache = table != nullptr ? wb.find_pivot_cache(table->pivot_cache_id()) : nullptr;
  EXPECT_NE(cache, nullptr);
  if (cache == nullptr) {
    return {};
  }
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(mac_365_ja_jp_profile());
  auto result_or = pivot::evaluate(*table, *cache, options, pivot::PivotFilterEnv{std::nullopt, wb.date1904()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return {};
  }
  const pivot::PivotResult& result = result_or.value();
  std::vector<Row> rows;
  std::vector<std::string> path;
  std::size_t leaf = 0;
  Flatten(result, result.rows, path, leaf, rows);
  rows.push_back({options.grand_total_label, Numbers(result.grand_totals)});
  return rows;
}

// "月" / "年" / "日" / "第 1 四半期" / "総計".
#define MONTH "\xE6\x9C\x88"
#define YEAR "\xE5\xB9\xB4"
#define DAY "\xE6\x97\xA5"
#define QUARTER1 "\xE7\xAC\xAC 1 \xE5\x9B\x9B\xE5\x8D\x8A\xE6\x9C\x9F"
#define TOTAL "\xE7\xB7\x8F\xE8\xA8\x88"

constexpr const char* kSheet = "Sheet5";
constexpr const char* kTable =
    "\xE3\x83\x94\xE3\x83\x9C\xE3\x83\x83\xE3\x83\x88\xE3\x83\x86\xE3\x83\xBC\xE3\x83\x96\xE3\x83\xAB"
    "4";  // ピボットテーブル4

const std::vector<Row> kUngroupedDates = {
    {"2024/1/15", {100}}, {"2024/1/25", {80}},  {"2024/2/8", {150}},
    {"2024/2/20", {50}},  {"2024/3/10", {200}}, {TOTAL, {580}},
};

TEST(PivotDateGroupFixture, UngroupedDateFieldIsLabelledInItsNumberFormat) {
  EXPECT_EQ(EvaluatedRows("date_field_ungrouped.xlsx", kSheet, kTable), kUngroupedDates);
}

TEST(PivotDateGroupFixture, UngroupedAgainDateFieldIsLabelledInItsNumberFormat) {
  EXPECT_EQ(EvaluatedRows("date_group_ungrouped.xlsx", kSheet, kTable), kUngroupedDates);
}

TEST(PivotDateGroupFixture, Months) {
  const std::vector<Row> expected = {{"1" MONTH, {180}}, {"2" MONTH, {200}}, {"3" MONTH, {200}}, {TOTAL, {580}}};
  EXPECT_EQ(EvaluatedRows("date_group_months.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, MonthsAndYears) {
  const std::vector<Row> expected = {
      {"2024" YEAR, {580}}, {"1" MONTH, {180}}, {"2" MONTH, {200}}, {"3" MONTH, {200}}, {TOTAL, {580}},
  };
  EXPECT_EQ(EvaluatedRows("date_group_months_years.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, QuartersMonthsAndYears) {
  const std::vector<Row> expected = {
      {"2024" YEAR, {580}}, {QUARTER1, {580}},  {"1" MONTH, {180}},
      {"2" MONTH, {200}},   {"3" MONTH, {200}}, {TOTAL, {580}},
  };
  EXPECT_EQ(EvaluatedRows("date_group_quarters_months_years.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, Days) {
  const std::vector<Row> expected = {
      {"1" MONTH "15" DAY, {100}}, {"1" MONTH "25" DAY, {80}},  {"2" MONTH "8" DAY, {150}},
      {"2" MONTH "20" DAY, {50}},  {"3" MONTH "10" DAY, {200}}, {TOTAL, {580}},
  };
  EXPECT_EQ(EvaluatedRows("date_group_days.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, DaysInIntervalsOfSeven) {
  const std::vector<Row> expected = {
      {"2024/1/15 - 2024/1/21", {100}}, {"2024/1/22 - 2024/1/28", {80}}, {"2024/2/5 - 2024/2/11", {150}},
      {"2024/2/19 - 2024/2/25", {50}},  {"2024/3/4 - 2024/3/10", {200}}, {TOTAL, {580}},
  };
  EXPECT_EQ(EvaluatedRows("date_group_days_interval7.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, Years) {
  const std::vector<Row> expected = {{"2024" YEAR, {580}}, {TOTAL, {580}}};
  EXPECT_EQ(EvaluatedRows("date_group_years.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, Quarters) {
  const std::vector<Row> expected = {{QUARTER1, {580}}, {TOTAL, {580}}};
  EXPECT_EQ(EvaluatedRows("date_group_quarters.xlsx", kSheet, kTable), expected);
}

TEST(PivotDateGroupFixture, RenamedItemShowsItsCaptionInItsOwnPlace) {
  const std::vector<Row> expected = {
      {"zz", {500, 5}},
      {"North", {300, 3}},
      {"South", {700, 7}},
      {TOTAL, {1500, 15}},
  };
  EXPECT_EQ(EvaluatedRows("item_caption.xlsx", "Tabelle1", "PivotTable1"), expected);
}

// Workbooks Excel saved with the years / 7-day group field on rows, before and after moving the
// earliest (C2) and latest (C9) source date and refreshing. The engine's refresh of the "before"
// book must give Excel's group items and the row field's item order.
struct DateGroupRefreshCase {
  const char* before;
  const char* after;
  std::optional<double> c2;
  std::optional<double> c9;
};

Workbook LoadPivotOptions(const std::string& file) {
  const std::vector<std::uint8_t> bytes = ReadFixture(file.c_str());
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << file;
  Workbook wb = std::move(read_or.value().workbook);
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

std::vector<std::string> GroupLabels(const Workbook& wb, std::uint32_t cache_id, std::size_t field) {
  std::vector<std::string> out;
  const pivot::PivotCache* cache = wb.find_pivot_cache(cache_id);
  if (cache == nullptr || field >= cache->fields().size() || !cache->fields()[field].field_group.has_value()) {
    ADD_FAILURE() << "cache " << cache_id << " has no group field " << field;
    return out;
  }
  for (const Value& item : cache->fields()[field].field_group->group_items) {
    out.emplace_back(item.is_text() ? std::string(item.as_text()) : std::string("<not text>"));
  }
  return out;
}

std::vector<std::uint32_t> FieldItemOrder(const Workbook& wb, std::size_t field) {
  std::vector<std::uint32_t> out;
  if (wb.sheet(0).pivot_tables().empty() || field >= wb.sheet(0).pivot_tables()[0]->fields().size()) {
    ADD_FAILURE() << "no pivot field " << field;
    return out;
  }
  for (const pivot::PivotItem& item : wb.sheet(0).pivot_tables()[0]->fields()[field].items) {
    if (item.has_cache_index) {
      out.push_back(item.cache_index);
    }
  }
  return out;
}

TEST(PivotDateGroupFixture, RefreshMovingYearsAndSevenDayBoundsMatchesExcel) {
  constexpr std::size_t kGroupField = 4U;
  const DateGroupRefreshCase cases[] = {
      {"years_before", "years_start", 45078.0, std::nullopt},  // 2023-06-01
      {"years_before", "years_end", std::nullopt, 45719.0},    // 2025-03-03
      {"years_before", "years_both", 45078.0, 45719.0},
      {"days7_before", "days7_start", 45301.0, std::nullopt},  // 2024-01-10
      {"days7_before", "days7_end", std::nullopt, 45514.0},    // 2024-08-10
      {"days7_before", "days7_both", 45301.0, 45514.0},
  };
  for (const DateGroupRefreshCase& c : cases) {
    SCOPED_TRACE(c.after);
    Workbook wb = LoadPivotOptions(std::string("refresh_dategroup/") + c.before + ".xlsx");
    const Workbook excel = LoadPivotOptions(std::string("refresh_dategroup/") + c.after + ".xlsx");
    std::size_t data = wb.sheet_count();
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      if (wb.sheet(s).name() == "Sheet1") {
        data = s;
      }
    }
    ASSERT_LT(data, wb.sheet_count());
    for (const auto& [row, serial] : {std::pair<std::uint32_t, std::optional<double>>{1U, c.c2}, {8U, c.c9}}) {
      if (serial.has_value()) {
        ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(data, row, 2U, Value::number(*serial))));
      }
    }
    ASSERT_FALSE(wb.sheet(0).pivot_tables().empty());
    const std::uint32_t cache_id = wb.sheet(0).pivot_tables()[0]->pivot_cache_id();
    const auto refreshed = wb.refresh_pivot_cache(cache_id, eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
    // Excel renumbers its caches on save; each book's pivot names its own.
    const std::uint32_t excel_cache_id = excel.sheet(0).pivot_tables()[0]->pivot_cache_id();
    EXPECT_EQ(GroupLabels(wb, cache_id, kGroupField), GroupLabels(excel, excel_cache_id, kGroupField));
    EXPECT_EQ(FieldItemOrder(wb, kGroupField), FieldItemOrder(excel, kGroupField));
  }
}

}  // namespace
}  // namespace formulon
