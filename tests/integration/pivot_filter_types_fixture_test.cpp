//
// Every pivot filter type against Mac Excel 365 (ja-JP) re-saves
// (`tests/fixtures/excel/pivot_filters/`, one file per case).
//
// Each file holds one pivot: Region on rows, Sum of Amt and Sum of Qty, over
// the items Apple 100/1, Apricot 300/3, Banana 400/4, Cherry 200/2 and
// *star 50/10 (the `ft_n*` files use the numeric labels 1, 2, 3, 4 and 10
// instead). The expected rows and grand totals are Excel's cached render.
//
// Two properties are checked per file: the evaluation keeps Excel's rows and
// totals, and the `<filters>` block written for the same filters added
// through the model (no source bytes) is the block Excel saved.

#include <cstdint>
#include <cstdio>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_table_writer.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
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

struct Row {
  const char* label;
  double amt;
  double qty;
};

struct FilterCase {
  const char* file;
  std::vector<Row> rows;
  double total_amt;
  double total_qty;
};

const Row kApple{"Apple", 100, 1};
const Row kApricot{"Apricot", 300, 3};
const Row kBanana{"Banana", 400, 4};
const Row kCherry{"Cherry", 200, 2};
const Row kStar{"*star", 50, 10};
const Row kN1{"1", 100, 1};
const Row kN2{"2", 300, 3};
const Row kN3{"3", 400, 4};
const Row kN4{"4", 200, 2};
const Row kN10{"10", 50, 10};

const std::vector<FilterCase>& Cases() {
  static const std::vector<FilterCase> cases = {
      {"ft_l01_eq", {kApple}, 100, 1},
      {"ft_l02_eq_case", {kApple}, 100, 1},
      {"ft_l03_eq_star", {kApple, kApricot}, 400, 4},
      {"ft_l04_eq_qmark", {kApple}, 100, 1},
      {"ft_l05_eq_fullwidth", {kApple}, 100, 1},
      {"ft_l06_eq_literal_star", {kStar}, 50, 10},
      {"ft_l06b_eq_star_unescaped", {kStar}, 50, 10},
      {"ft_l07_ne", {kApricot, kBanana, kCherry, kStar}, 950, 19},
      {"ft_l08_bw", {kApple, kApricot}, 400, 4},
      {"ft_l09_bw_case", {kApple, kApricot}, 400, 4},
      {"ft_l10_nbw", {kBanana, kCherry, kStar}, 650, 16},
      {"ft_l11_ew", {kApple}, 100, 1},
      {"ft_l12_new", {kApricot, kBanana, kCherry, kStar}, 950, 19},
      {"ft_l13_ct", {kApricot}, 300, 3},
      {"ft_l14_ct_star", {kApple, kApricot, kBanana, kCherry, kStar}, 1050, 20},
      {"ft_l15_nct", {kBanana, kCherry, kStar}, 650, 16},
      {"ft_l16_gt", {kCherry}, 200, 2},
      {"ft_l17_gt_wild", {kApple, kApricot, kCherry, kStar}, 650, 16},
      {"ft_l18_ge", {kBanana, kCherry}, 600, 6},
      {"ft_l19_lt", {kApple, kApricot, kStar}, 450, 14},
      {"ft_l20_le", {kApple, kApricot, kBanana, kStar}, 850, 18},
      {"ft_l21_between", {kApricot, kBanana, kCherry}, 900, 9},
      {"ft_l22_notbetween", {kApple, kStar}, 150, 11},
      {"ft_m01_two_filters_mff0", {kApricot}, 300, 3},
      {"ft_m02_two_filters_mff1", {kApricot}, 300, 3},
      {"ft_n01_gt3", {kN4, kN10}, 250, 12},
      {"ft_n02_between", {kN2, kN3, kN4}, 900, 9},
      {"ft_n03_eq10", {kN10}, 50, 10},
      {"ft_n04_lt3", {kN1, kN2}, 400, 4},
      {"ft_n05_ge10", {kN10}, 50, 10},
      {"ft_n06_bw1", {kN1, kN10}, 150, 11},
      {"ft_t01_top2", {kApricot, kBanana}, 700, 7},
      {"ft_t02_bottom2", {kApple, kStar}, 150, 11},
      {"ft_t03_toppct50", {kApricot, kBanana}, 700, 7},
      {"ft_t04_bottompct20", {kApple, kCherry, kStar}, 350, 13},
      {"ft_t05_topsum700", {kApricot, kBanana}, 700, 7},
      {"ft_t06_bottomsum350", {kApple, kCherry, kStar}, 350, 13},
      {"ft_t07_toppct100", {kApple, kApricot, kBanana, kCherry, kStar}, 1050, 20},
      {"ft_t08_top10_of5", {kApple, kApricot, kBanana, kCherry, kStar}, 1050, 20},
      {"ft_v01_eq", {kApricot}, 300, 3},
      {"ft_v02_ne", {kApple, kBanana, kCherry, kStar}, 750, 17},
      {"ft_v03_ge", {kApricot, kBanana}, 700, 7},
      {"ft_v04_lt", {kApple, kCherry, kStar}, 350, 13},
      {"ft_v05_le", {kApple, kApricot, kCherry, kStar}, 650, 16},
      {"ft_v06_between", {kApple, kApricot, kCherry}, 600, 6},
      {"ft_v07_notbetween", {kBanana, kStar}, 450, 14},
      {"ft_v08_gt", {kBanana}, 400, 4},
  };
  return cases;
}

std::string FixturePath(const char* file) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_filters/" + file + ".xlsx";
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
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

// The `<filters count=...>` block of a pivot table part; empty when absent.
std::string FiltersBlock(const std::string& xml) {
  const std::size_t begin = xml.find("<filters count=");
  const std::size_t end = xml.rfind("</filters>");
  if (begin == std::string::npos || end == std::string::npos || end < begin) {
    return {};
  }
  return xml.substr(begin, end + std::string("</filters>").size() - begin);
}

const pivot::PivotTable* FirstPivot(const Workbook& wb) {
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (!wb.sheet(i).pivot_tables().empty()) {
      return wb.sheet(i).pivot_tables().front().get();
    }
  }
  return nullptr;
}

TEST(PivotFilterTypesFixture, EvaluationMatchesTheExcelFilteredRender) {
  ExcelProfile profile;
  ASSERT_TRUE(parse_excel_profile_id("mac-365-ja_JP", &profile));
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(profile);
  std::size_t checked = 0;
  for (const FilterCase& c : Cases()) {
    SCOPED_TRACE(c.file);
    const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath(c.file));
    ASSERT_FALSE(bytes.empty());
    auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const Workbook& wb = loaded.value().workbook;
    const pivot::PivotTable* table = FirstPivot(wb);
    ASSERT_NE(table, nullptr);
    EXPECT_FALSE(table->active_filters().empty());
    EXPECT_TRUE(table->retained_filters().empty());
    const pivot::PivotCache* cache = wb.find_pivot_cache(table->pivot_cache_id());
    ASSERT_NE(cache, nullptr);

    auto result_or = pivot::evaluate(*table, *cache, options);
    ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
    const pivot::PivotResult& result = result_or.value();
    ASSERT_EQ(result.rows.size(), c.rows.size());
    ASSERT_EQ(result.values.size(), c.rows.size());
    for (std::size_t i = 0; i < c.rows.size(); ++i) {
      EXPECT_EQ(result.rows[i].label, c.rows[i].label) << "row " << i;
      ASSERT_FALSE(result.values[i].empty());
      ASSERT_EQ(result.values[i][0].size(), 2U);
      EXPECT_DOUBLE_EQ(result.values[i][0][0].as_number(), c.rows[i].amt) << c.rows[i].label;
      EXPECT_DOUBLE_EQ(result.values[i][0][1].as_number(), c.rows[i].qty) << c.rows[i].label;
    }
    ASSERT_EQ(result.grand_totals.size(), 2U);
    EXPECT_DOUBLE_EQ(result.grand_totals[0].as_number(), c.total_amt);
    EXPECT_DOUBLE_EQ(result.grand_totals[1].as_number(), c.total_qty);
    ++checked;
  }
  EXPECT_EQ(checked, 47U);
}

// The filters rebuilt from the model alone, ids renumbered from 1, spell the
// block Excel saved.
TEST(PivotFilterTypesFixture, GeneratedFiltersSpellTheExcelBlock) {
  for (const FilterCase& c : Cases()) {
    SCOPED_TRACE(c.file);
    const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath(c.file));
    ASSERT_FALSE(bytes.empty());
    auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const pivot::PivotTable* loaded_table = FirstPivot(loaded.value().workbook);
    ASSERT_NE(loaded_table, nullptr);

    io::ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(io::ByteSpan{bytes.data(), bytes.size()})));
    auto part_or = zip.read_entry("xl/pivotTables/pivotTable1.xml");
    ASSERT_TRUE(static_cast<bool>(part_or)) << "missing pivotTable1.xml";
    const std::string excel_block = FiltersBlock(std::string(part_or.value().begin(), part_or.value().end()));
    ASSERT_FALSE(excel_block.empty());

    std::unique_ptr<pivot::PivotTable> table = loaded_table->clone();
    for (pivot::PivotFilter& filter : table->mutable_active_filters()) {
      ASSERT_FALSE(filter.field_name.empty());
      filter.source.reset();
    }
    EXPECT_EQ(FiltersBlock(io::write_pivot_table_definition(*table)), excel_block);
  }
}

}  // namespace
}  // namespace formulon
