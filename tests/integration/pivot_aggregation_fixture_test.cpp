//
// Compares non-Sum pivot aggregation against the values Mac Excel saved.
//
// Each fixture under `tests/fixtures/excel/pivot_aggregation/` is a workbook
// whose pivot Excel refreshed and saved, so the sheet cells inside the pivot's
// `ref` are Excel's own numbers. The cases cover filtered record order (hidden
// items, a selected page item, label and value filters), `subtotalHiddenItems`,
// the total strips of deep hierarchies, and Sum totals beside non-Sum custom
// subtotals. Numbers are compared bit for bit.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <iomanip>
#include <map>
#include <ostream>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

struct AggregationFixture {
  const char* file;  // under tests/fixtures/excel/pivot_aggregation/
};

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
  if (v.is_blank()) {
    return "";
  }
  if (v.is_text()) {
    return std::string(v.as_text());
  }
  if (v.is_number()) {
    std::ostringstream os;
    os << std::setprecision(17) << v.as_number();
    return os.str();
  }
  return "#ERR";
}

void PrintTo(const AggregationFixture& fx, std::ostream* os) {
  *os << fx.file;
}

class PivotAggregationFixture : public ::testing::TestWithParam<AggregationFixture> {};

TEST_P(PivotAggregationFixture, EngineMatchesTheValuesExcelSaved) {
  const AggregationFixture& fx = GetParam();
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_aggregation/" + fx.file;
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
  auto result_or = pivot::evaluate(table, *cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = pivot::layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;

  // Label cells compare by rendered text: Excel stores a numeric item as a number.
  const pivot::PivotCells& grid = cells_or.value();
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::string> engine;
  for (const pivot::PivotCell& cell : grid.cells) {
    if (!cell.value.is_blank()) {
      engine[{cell.row, cell.col}] = Render(cell.value);
    }
  }
  std::ostringstream diff;
  for (std::uint32_t r = grid.top; r < grid.top + grid.rows; ++r) {
    for (std::uint32_t c = grid.left; c < grid.left + grid.cols; ++c) {
      const Cell* cell = sheet->cell_at(r, c);
      const std::string excel = cell == nullptr ? std::string() : Render(cell->cached_value);
      const auto it = engine.find({r, c});
      const std::string ours = it == engine.end() ? std::string() : it->second;
      if (excel != ours) {
        diff << "  R" << (r + 1) << "C" << (c + 1) << ": excel=" << excel << " engine=" << ours << "\n";
      }
    }
  }
  EXPECT_TRUE(diff.str().empty()) << fx.file << "\n" << diff.str();
}

std::string FixtureName(const ::testing::TestParamInfo<AggregationFixture>& info) {
  std::string name;
  for (const char* p = info.param.file; *p != '\0' && *p != '.'; ++p) {
    name.push_back((*p >= 'a' && *p <= 'z') || (*p >= 'A' && *p <= 'Z') || (*p >= '0' && *p <= '9') ? *p : '_');
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(ExcelSaved, PivotAggregationFixture,
                         ::testing::Values(AggregationFixture{"HA1r__max.xlsx"},
                                           AggregationFixture{"HQ1r__product.xlsx"},
                                           AggregationFixture{"G1__average.xlsx"}, AggregationFixture{"Gx__count.xlsx"},
                                           AggregationFixture{"OF1__sum.xlsx"}, AggregationFixture{"D1__max.xlsx"},
                                           AggregationFixture{"D2__max.xlsx"}, AggregationFixture{"OV__max.xlsx"},
                                           AggregationFixture{"K4__average.xlsx"}, AggregationFixture{"L22__var.xlsx"},
                                           AggregationFixture{"V1Z22__sum.xlsx"}, AggregationFixture{"V3Z22__sum.xlsx"},
                                           AggregationFixture{"V4Z21__sum.xlsx"}, AggregationFixture{"C9__sum.xlsx"},
                                           AggregationFixture{"W1K1__var.xlsx"},
                                           AggregationFixture{"FHQ__average.xlsx"}),
                         FixtureName);

}  // namespace
}  // namespace formulon
