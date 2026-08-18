//
// Checks how a blank source cell reaches the axis, against a real Excel
// 365 (ja-JP) workbook (`tests/fixtures/excel/pivot_blank_item.xlsx`).
//
// Source table (`Sheet1`, A1:B5) — Region / Amt, with A3 left empty:
//   North 100
//         200
//   South 300
//   North  50
//
// Excel's cached render (`Sheet2`, A3:B7) is what the expectations below
// are measured against:
//   行ラベル   合計 / Amt
//   North             150
//   South             300
//   (空白)            200
//   総計              650
//
// Two properties of Excel's bytes are why this fixture exists, and both
// contradict a plausible reading of the schema:
//
//   * The pivot field spells the blank as an ordinary `<item x="1"/>`.
//     There is a `t="blank"` item type, and it is *not* what Excel writes
//     for an empty source cell — the blank lives in the cache, as an
//     `<m/>` shared item under `containsBlank="1"`. An implementation
//     that keyed off the item type would see no blank at all here.
//   * The blank sorts last, after South. Its label begins with `(`, which
//     orders before every letter, so a placeholder resolved before the
//     sort would have put it first. Excel orders the axis by the source
//     value and names the node afterwards.
//
// The render also pins `blank_item_label` for ja-JP: Excel wrote the
// literal `(空白)` into the grid, which is the string
// `pivot_layout_options_for` supplies.

#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "eval/pivot_locale.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
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

// Sheet 0 is the report tab Excel created ("Sheet2"); the source table
// sits behind it on "Sheet1".
constexpr std::size_t kReportSheet = 0;

// Region aggregates, as Excel rendered them.
constexpr double kNorth = 150.0;
constexpr double kSouth = 300.0;
constexpr double kBlank = 200.0;

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
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_blank_item.xlsx";
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

const pivot::PivotTable* ReportPivot(const Workbook& workbook) {
  if (kReportSheet >= workbook.sheet_count() || workbook.sheet(kReportSheet).pivot_tables().empty()) {
    return nullptr;
  }
  return workbook.sheet(kReportSheet).pivot_tables().front().get();
}

// Evaluates the fixture's single pivot under `options` and returns its
// row labels paired with the one data-field aggregate on each row.
std::vector<std::pair<std::string, double>> EvaluateRows(const Workbook& workbook,
                                                         const pivot::PivotLayoutOptions& options) {
  std::vector<std::pair<std::string, double>> rows;
  const pivot::PivotTable* table = ReportPivot(workbook);
  EXPECT_NE(table, nullptr);
  if (table == nullptr) {
    return rows;
  }
  const pivot::PivotCache* cache = workbook.find_pivot_cache(table->pivot_cache_id());
  EXPECT_NE(cache, nullptr);
  if (cache == nullptr) {
    return rows;
  }
  auto result_or = pivot::evaluate(*table, *cache, options);
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
// (a) Where Excel puts the blank.
// ---------------------------------------------------------------------------

TEST(PivotBlankItemFixture, TheCacheCarriesTheBlankAsASharedItem) {
  const Workbook workbook = LoadFixture();
  const pivot::PivotTable* table = ReportPivot(workbook);
  ASSERT_NE(table, nullptr);
  const pivot::PivotCache* cache = workbook.find_pivot_cache(table->pivot_cache_id());
  ASSERT_NE(cache, nullptr);
  ASSERT_FALSE(cache->fields().empty());

  const pivot::PivotCacheField& region = cache->fields().front();
  EXPECT_EQ(region.name, "Region");
  // Excel wrote `<s v="North"/><m/><s v="South"/>` — the blank keeps a
  // slot of its own, so record indices stay aligned with it present.
  ASSERT_EQ(region.shared_items.size(), 3U);
  EXPECT_TRUE(region.shared_items[0].is_text());
  EXPECT_TRUE(region.shared_items[1].is_blank());
  EXPECT_TRUE(region.shared_items[2].is_text());

  EXPECT_TRUE(region.shared_items_hints.has_contains_blank);
  EXPECT_TRUE(region.shared_items_hints.contains_blank);
}

// `t="blank"` exists as an item type and is not what an empty source cell
// produces. Pinning this keeps the classification from being "simplified"
// later into reading the marker, which would find no blank on this file.
TEST(PivotBlankItemFixture, TheFieldItemIsOrdinaryAndNotMarkedBlank) {
  const Workbook workbook = LoadFixture();
  const pivot::PivotTable* table = ReportPivot(workbook);
  ASSERT_NE(table, nullptr);
  ASSERT_FALSE(table->fields().empty());

  const pivot::PivotField& region = table->fields().front();
  // North, South and the blank — the trailing `<item t="default"/>` is the
  // subtotal marker and is not an axis item.
  ASSERT_EQ(region.items.size(), 3U);
  for (const pivot::PivotItem& item : region.items) {
    EXPECT_TRUE(item.visible) << "item '" << item.name << "' carries h=\"1\"";
    EXPECT_TRUE(item.has_cache_index) << "item '" << item.name << "' lost its `x` attribute";
  }
  // Excel's item order is its axis order, and the blank's cache slot (1)
  // appears last rather than in cache order.
  EXPECT_EQ(region.items[0].cache_index, 0U);
  EXPECT_EQ(region.items[1].cache_index, 2U);
  EXPECT_EQ(region.items[2].cache_index, 1U);
}

// ---------------------------------------------------------------------------
// (b) Evaluation reproduces Excel's render.
// ---------------------------------------------------------------------------

TEST(PivotBlankItemFixture, TheBlankGroupSortsLastAndAggregatesItsRows) {
  const Workbook workbook = LoadFixture();
  const std::vector<std::pair<std::string, double>> rows = EvaluateRows(workbook, pivot::PivotLayoutOptions{});
  ASSERT_EQ(rows.size(), 3U);
  EXPECT_EQ(rows[0].first, "North");
  EXPECT_DOUBLE_EQ(rows[0].second, kNorth);
  EXPECT_EQ(rows[1].first, "South");
  EXPECT_DOUBLE_EQ(rows[1].second, kSouth);
  // Last, despite a label that would sort first.
  EXPECT_EQ(rows[2].first, pivot::PivotLayoutOptions{}.blank_item_label);
  EXPECT_DOUBLE_EQ(rows[2].second, kBlank);
}

// The label Excel cached in the grid for the ja-JP profile, which is the
// one observation `blank_item_label`'s spelling rests on.
TEST(PivotBlankItemFixture, TheJaJpProfileNamesTheBlankGroupAsExcelDid) {
  const Workbook workbook = LoadFixture();
  eval::ExcelProfile profile;
  profile.locale = eval::ExcelLocale::kJaJP;
  const std::vector<std::pair<std::string, double>> rows =
      EvaluateRows(workbook, eval::pivot_layout_options_for(profile));
  ASSERT_EQ(rows.size(), 3U);
  EXPECT_EQ(rows[2].first, "(空白)");
  EXPECT_DOUBLE_EQ(rows[2].second, kBlank);
}

}  // namespace
}  // namespace formulon
