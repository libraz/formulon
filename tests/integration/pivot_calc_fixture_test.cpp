//
// Workbooks Mac Excel 365 saved with calculated fields and calculated items
// (tests/fixtures/excel/pivot_calc/), and with `showAll` fields that a filter
// on another field leaves empty (tests/fixtures/excel/pivot_options/op_*).
//
// The cells inside each pivot's range are Excel's own. The pivot evaluated
// with a formula host and projected under the ja-JP profile the files were
// saved in must draw the same grid, numbers compared bit for bit.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <iomanip>
#include <map>
#include <ostream>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "eval/pivot_formula.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "pivot/value_order.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

struct CalcFixture {
  const char* dir;  // under tests/fixtures/excel/
  const char* file;
};

void PrintTo(const CalcFixture& fx, std::ostream* os) {
  *os << fx.dir << "/" << fx.file;
}

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

std::string Render(const Value& v) {
  if (v.is_number()) {
    std::ostringstream os;
    os << std::setprecision(17) << v.as_number();
    return os.str();
  }
  return pivot::display_string(v);
}

constexpr std::uint32_t kScanRows = 64;
constexpr std::uint32_t kScanCols = 16;

class PivotCalcFixture : public ::testing::TestWithParam<CalcFixture> {};

TEST_P(PivotCalcFixture, EngineDrawsTheGridExcelSaved) {
  const CalcFixture& fx = GetParam();
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + fx.dir + "/" + fx.file;
  const std::vector<std::uint8_t> bytes = ReadFileBytes(path);
  ASSERT_FALSE(bytes.empty()) << "could not read " << path;
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  const Workbook& workbook = read_or.value().workbook;

  const Sheet* sheet = nullptr;
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    if (!workbook.sheet(i).pivot_tables().empty()) {
      sheet = &workbook.sheet(i);
    }
  }
  ASSERT_NE(sheet, nullptr) << fx.file << ": no pivot";
  const pivot::PivotTable& table = *sheet->pivot_tables().front();
  const pivot::PivotCache* cache = workbook.find_pivot_cache(table.pivot_cache_id());
  ASSERT_NE(cache, nullptr);

  ExcelProfile profile;
  ASSERT_TRUE(parse_excel_profile_id("mac-365-ja_JP", &profile));
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(profile);
  eval::PivotFormulaEvaluator host;
  pivot::PivotFilterEnv env;
  env.formulas = &host;
  auto result_or = pivot::evaluate(table, *cache, options, env);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = pivot::layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;

  const pivot::PivotCells& grid = cells_or.value();
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::string> engine;
  for (const pivot::PivotCell& cell : grid.cells) {
    if (!cell.value.is_blank()) {
      engine[{cell.row, cell.col}] = Render(cell.value);
    }
  }
  // The report sheet holds nothing but the report, so every non-blank cell
  // in its corner is Excel's; the engine's range must reach all of them.
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::string> excel;
  for (std::uint32_t r = 0; r < kScanRows; ++r) {
    for (std::uint32_t c = 0; c < kScanCols; ++c) {
      const Cell* cell = sheet->cell_at(r, c);
      if (cell != nullptr && !cell->cached_value.is_blank()) {
        excel[{r, c}] = Render(cell->cached_value);
      }
    }
  }

  std::ostringstream diff;
  for (const auto& [at, text] : excel) {
    const auto it = engine.find(at);
    if (it == engine.end() || it->second != text) {
      diff << "  R" << (at.first + 1) << "C" << (at.second + 1) << ": excel=" << text
           << " engine=" << (it == engine.end() ? std::string() : it->second) << "\n";
    }
  }
  for (const auto& [at, text] : engine) {
    if (excel.find(at) == excel.end()) {
      diff << "  R" << (at.first + 1) << "C" << (at.second + 1) << ": excel= engine=" << text << "\n";
    }
  }
  EXPECT_TRUE(diff.str().empty()) << fx.file << "\n" << diff.str();
}

std::string FixtureName(const ::testing::TestParamInfo<CalcFixture>& info) {
  std::string name;
  for (const char* p = info.param.file; *p != '\0' && *p != '.'; ++p) {
    name.push_back((*p >= 'a' && *p <= 'z') || (*p >= 'A' && *p <= 'Z') || (*p >= '0' && *p <= '9') ? *p : '_');
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(ExcelSaved, PivotCalcFixture,
                         ::testing::Values(CalcFixture{"pivot_calc", "cf1_price_amt_per_qty.xlsx"},
                                           CalcFixture{"pivot_calc", "cf2_div_zero_error_caption.xlsx"},
                                           CalcFixture{"pivot_calc", "cf3_quoted_names_if.xlsx"},
                                           CalcFixture{"pivot_calc", "cf4_function_average_count.xlsx"},
                                           CalcFixture{"pivot_calc", "cf5_nested_calc_field.xlsx"},
                                           CalcFixture{"pivot_calc", "ci1_item_sum_two_items.xlsx"},
                                           CalcFixture{"pivot_calc", "ci2_item_with_count.xlsx"},
                                           CalcFixture{"pivot_calc", "ci3_item_positional_refs.xlsx"},
                                           CalcFixture{"pivot_calc", "ci4_items_row_and_column.xlsx"},
                                           CalcFixture{"pivot_calc", "ci4b_intersection_row_item_first.xlsx"},
                                           CalcFixture{"pivot_calc", "ci4c_intersection_column_item_first.xlsx"},
                                           CalcFixture{"pivot_options", "op_emptyrow_show_all.xlsx"},
                                           CalcFixture{"pivot_options", "op_emptycol_show_all.xlsx"},
                                           CalcFixture{"pivot_options", "op_showall_only.xlsx"}),
                         FixtureName);

}  // namespace
}  // namespace formulon
