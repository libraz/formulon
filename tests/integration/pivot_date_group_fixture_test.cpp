//
// Workbooks Mac Excel 365 (ja-JP) saved after grouping a pivot's date field
// by months, quarters, years and days, after ungrouping it again, and after
// renaming an item. Each pivot is evaluated as loaded, and its row labels and
// values must equal what Excel drew: one entry per rendered row, a parent
// row carrying its subtotal, then the grand total.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <limits>
#include <memory>
#include <optional>
#include <ostream>
#include <string>
#include <utility>
#include <vector>

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

}  // namespace
}  // namespace formulon
