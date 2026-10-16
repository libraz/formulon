//
// Compares the "show values as" modes against the values Mac Excel saved.
//
// Each fixture under `tests/fixtures/excel/pivot_show_as/` is a workbook whose
// pivot Excel refreshed and saved, so the sheet cells inside the pivot's `ref`
// are Excel's own numbers. The cases cover percent / index / difference /
// running-total / parent modes, error and blank cells, and all-zero totals.
// Numbers are compared bit for bit.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <iomanip>
#include <ostream>
#include <set>
#include <sstream>
#include <string>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
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

struct ShowAsFixture {
  const char* file;  // under tests/fixtures/excel/pivot_show_as/
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

void PrintTo(const ShowAsFixture& fx, std::ostream* os) {
  *os << fx.file;
}

class PivotShowAsFixture : public ::testing::TestWithParam<ShowAsFixture> {};

TEST_P(PivotShowAsFixture, EngineMatchesTheValuesExcelSaved) {
  const ShowAsFixture& fx = GetParam();
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_show_as/" + fx.file;
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

  // Excel keeps an item list without `sortType` in its authored order, so cells
  // are compared as a multiset of rendered values rather than by position.
  const std::uint32_t top = table.anchor_row();
  const std::uint32_t left = table.anchor_col();
  std::multiset<std::string> engine;
  for (const pivot::PivotCell& cell : cells_or.value().cells) {
    if (!cell.value.is_blank()) {
      engine.insert(Render(cell.value));
    }
  }
  // Page-field headers sit above the pivot's `ref`.
  for (const pivot::PivotPageSelection& page : result_or.value().page_selections) {
    for (const std::string& label : {page.field_label, page.item_label}) {
      if (const auto it = engine.find(label); it != engine.end()) {
        engine.erase(it);
      }
    }
  }
  std::multiset<std::string> excel;
  const std::uint32_t bottom = std::max(top + table.span_rows(), top + cells_or.value().rows);
  const std::uint32_t right = std::max(left + table.span_cols(), left + cells_or.value().cols);
  for (std::uint32_t r = top; r < bottom; ++r) {
    for (std::uint32_t c = left; c < right; ++c) {
      const Cell* cell = sheet->cell_at(r, c);
      if (cell != nullptr && !cell->cached_value.is_blank()) {
        excel.insert(Render(cell->cached_value));
      }
    }
  }
  std::ostringstream diff;
  for (const std::string& v : excel) {
    if (engine.count(v) < excel.count(v)) {
      diff << "  missing from engine: " << v << "\n";
    }
  }
  for (const std::string& v : engine) {
    if (excel.count(v) < engine.count(v)) {
      diff << "  extra in engine: " << v << "\n";
    }
  }
  EXPECT_TRUE(diff.str().empty()) << fx.file << "\n" << diff.str();
}

std::string FixtureName(const ::testing::TestParamInfo<ShowAsFixture>& info) {
  std::string name;
  for (const char* p = info.param.file; *p != '\0' && *p != '.'; ++p) {
    name.push_back((*p >= 'a' && *p <= 'z') || (*p >= 'A' && *p <= 'Z') || (*p >= '0' && *p <= '9') ? *p : '_');
  }
  return name;
}

// The x14 `pivotShowAs` extension Excel wrote survives a save and reload.
TEST(PivotShowAsRoundTrip, X14ModeAndBaseFieldSurviveSave) {
  struct Case {
    const char* file;
    pivot::ShowValuesAs mode;
  };
  const Case cases[] = {{"K1_parentRow__sum.xlsx", pivot::ShowValuesAs::PercentOfParentRow},
                        {"K1_parentRegion__sum.xlsx", pivot::ShowValuesAs::PercentOfParent},
                        {"K1_rankDescGrp__sum.xlsx", pivot::ShowValuesAs::RankDescending},
                        {"S8_pctRunR__sum.xlsx", pivot::ShowValuesAs::PercentOfRunningTotal}};
  const auto data_field_of = [](const Workbook& workbook) -> const pivot::PivotDataField* {
    for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
      if (!workbook.sheet(i).pivot_tables().empty()) {
        return &workbook.sheet(i).pivot_tables().front()->data_fields().front();
      }
    }
    return nullptr;
  };
  for (const Case& c : cases) {
    const std::vector<std::uint8_t> bytes =
        ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_show_as/" + c.file);
    ASSERT_FALSE(bytes.empty()) << c.file;
    auto first = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
    ASSERT_TRUE(static_cast<bool>(first)) << first.error().message;
    const pivot::PivotDataField* before = data_field_of(first.value().workbook);
    ASSERT_NE(before, nullptr) << c.file;
    EXPECT_EQ(before->show_as, c.mode) << c.file;

    auto saved = io::write_ooxml(first.value().workbook);
    ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
    auto second = io::read_ooxml(io::ByteSpan{saved.value().data(), saved.value().size()});
    ASSERT_TRUE(static_cast<bool>(second)) << second.error().message;
    const pivot::PivotDataField* after = data_field_of(second.value().workbook);
    ASSERT_NE(after, nullptr) << c.file;
    EXPECT_EQ(after->show_as, c.mode) << c.file;
    EXPECT_EQ(after->show_as_base_field, before->show_as_base_field) << c.file;
    EXPECT_EQ(after->show_as_base_item, before->show_as_base_item) << c.file;
  }
}

INSTANTIATE_TEST_SUITE_P(
    ExcelSaved, PivotShowAsFixture,
    ::testing::Values(ShowAsFixture{"S8_pct_total__sum.xlsx"}, ShowAsFixture{"S8_index__sum.xlsx"},
                      ShowAsFixture{"S8_pctdiffR1__sum.xlsx"}, ShowAsFixture{"S8_pctofR1__sum.xlsx"},
                      ShowAsFixture{"S8_runR__sum.xlsx"}, ShowAsFixture{"S8_diffP2__sum.xlsx"},
                      ShowAsFixture{"S8_pctRunR__sum.xlsx"}, ShowAsFixture{"S8e_pctdiffPrevR__sum.xlsx"},
                      ShowAsFixture{"S8e_runR__sum.xlsx"}, ShowAsFixture{"S8g_index__sum.xlsx"},
                      ShowAsFixture{"S8g_pctdiffR1__sum.xlsx"}, ShowAsFixture{"K1_parentRow__sum.xlsx"},
                      ShowAsFixture{"K1_parentRegion__sum.xlsx"}, ShowAsFixture{"K1_rankDescGrp__sum.xlsx"}),
    FixtureName);

}  // namespace
}  // namespace formulon
