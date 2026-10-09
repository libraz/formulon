//
// Compares the pivot layout projection against the grid Excel itself saved.
//
// Each fixture under `tests/fixtures/excel/pivot_layout/` is a workbook whose
// pivot Mac Excel 365 laid out through its own UI automation (row / column /
// Values placement and report form) and then saved, so the sheet cells inside
// the pivot's `ref` are Excel's render. The test projects the same pivot under
// the profile of the UI locale the file was laid out in and requires the two
// grids to agree cell for cell. The Values caption is the one stored in the
// file, which is why the `de` sources keep "Werte" under a ja-JP UI.

#include <algorithm>
#include <cstdint>
#include <cstdio>
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

struct LayoutFixture {
  const char* file;     // under tests/fixtures/excel/pivot_layout/
  const char* profile;  // the UI locale Excel laid the pivot out in
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
    return "'" + std::string(v.as_text()) + "'";
  }
  if (v.is_number()) {
    std::ostringstream os;
    os << v.as_number();
    return os.str();
  }
  return "?";
}

void PrintTo(const LayoutFixture& fx, std::ostream* os) {
  *os << fx.file << " under " << fx.profile;
}

class PivotLayoutFixture : public ::testing::TestWithParam<LayoutFixture> {};

TEST_P(PivotLayoutFixture, ProjectionMatchesTheGridExcelSaved) {
  const LayoutFixture& fx = GetParam();
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_layout/" + fx.file;
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
  ASSERT_TRUE(parse_excel_profile_id(fx.profile, &profile));
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(profile);
  auto result_or = pivot::evaluate(table, *cache, options);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  auto cells_or = pivot::layout(table, result_or.value(), options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message;

  // The engine grid, then every cell of the authored `ref` the engine left out.
  std::map<std::pair<std::uint32_t, std::uint32_t>, std::string> engine;
  for (const pivot::PivotCell& cell : cells_or.value().cells) {
    engine[{cell.row, cell.col}] = Render(cell.value);
  }
  const std::uint32_t top = table.anchor_row();
  const std::uint32_t left = table.anchor_col();
  const std::uint32_t bottom = std::max(top + table.span_rows(), top + cells_or.value().rows);
  const std::uint32_t right = std::max(left + table.span_cols(), left + cells_or.value().cols);
  std::ostringstream diff;
  for (std::uint32_t r = top; r < bottom; ++r) {
    for (std::uint32_t c = left; c < right; ++c) {
      const Cell* cell = sheet->cell_at(r, c);
      const std::string excel = cell == nullptr ? std::string() : Render(cell->cached_value);
      const auto it = engine.find({r, c});
      const std::string ours = it == engine.end() ? std::string() : it->second;
      if (excel != ours) {
        diff << "  R" << (r + 1) << "C" << (c + 1) << ": excel=" << excel << " engine=" << ours << "\n";
      }
    }
  }
  EXPECT_TRUE(diff.str().empty()) << fx.file << " under " << fx.profile << "\n" << diff.str();
}

std::string FixtureName(const ::testing::TestParamInfo<LayoutFixture>& info) {
  std::string name;
  for (const char* p = info.param.file; *p != '\0' && *p != '.'; ++p) {
    name.push_back((*p >= 'a' && *p <= 'z') || (*p >= 'A' && *p <= 'Z') || (*p >= '0' && *p <= '9') ? *p : '_');
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(ExcelSaved, PivotLayoutFixture,
                         ::testing::Values(LayoutFixture{"ja_c0_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_c1_values_on_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_c2_cols_values_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_c3_cols_values_cols.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_c4_tabular.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d0_outline_values_cols.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d1_outline_rows_only.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d2_tabular_rows_only.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d3_tabular_values_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d4_outline_values_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d5_values_first_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d6_values_first_cols.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d7_two_rows_values_cols.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d8_two_rows_values_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d9_two_cols_values_cols.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d10_tabular_two_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d11_outline_two_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d12_values_cols_no_rows.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"ja_d13_values_only.xlsx", "mac-365-ja_JP"},
                                           LayoutFixture{"en_c0_rows.xlsx", "mac-365-en_US"},
                                           LayoutFixture{"en_c1_values_on_rows.xlsx", "mac-365-en_US"},
                                           LayoutFixture{"en_c2_cols_values_rows.xlsx", "mac-365-en_US"},
                                           LayoutFixture{"en_c3_cols_values_cols.xlsx", "mac-365-en_US"},
                                           LayoutFixture{"en_c4_tabular.xlsx", "mac-365-en_US"}),
                         FixtureName);

}  // namespace
}  // namespace formulon
