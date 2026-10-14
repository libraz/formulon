#include <chrono>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_cache.h"
#include "chart/chart_model.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

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

bool IsXlsb(const std::string& name) {
  return name.size() > 5 && name.substr(name.size() - 5) == ".xlsb";
}

Workbook LoadBytes(const std::vector<std::uint8_t>& bytes, bool xlsb) {
  Workbook wb = Workbook::create_empty();
  if (!bytes.empty()) {
    const io::ByteSpan span{bytes.data(), bytes.size()};
    if (xlsb) {
      auto read = io::xlsb::read_xlsb(span);
      EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
      if (read) {
        wb = std::move(read.value().workbook);
      }
    } else {
      auto read = io::read_ooxml(span);
      EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
      if (read) {
        wb = std::move(read.value().workbook);
      }
    }
  }
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

Workbook LoadFixture(const std::string& name) {
  return LoadBytes(ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name), IsXlsb(name));
}

/// Saves `wb` in the container named by `xlsb` and reads it back.
Workbook SaveAndReload(const Workbook& wb, bool xlsb) {
  auto written = xlsb ? io::xlsb::write_xlsb(wb) : io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(written)) << (written ? "" : written.error().message);
  return written ? LoadBytes(written.value(), xlsb) : Workbook::create_empty();
}

bool Contains(const std::string& text, const std::string& needle) {
  return text.find(needle) != std::string::npos;
}

std::string PartText(const Workbook& wb, const std::string& path) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  EXPECT_NE(part, nullptr) << path;
  return part == nullptr ? std::string() : std::string(part->bytes.begin(), part->bytes.end());
}

std::vector<std::string> ChartPaths(const Workbook& wb) {
  std::vector<std::string> paths;
  auto refs = all_chart_parts(wb);
  EXPECT_TRUE(static_cast<bool>(refs));
  if (refs) {
    for (const SheetChartPartRef& entry : refs.value()) {
      paths.push_back(entry.ref.part_path);
    }
  }
  return paths;
}

// An Excel-saved chart saved without edits keeps its part bytes.
TEST(ChartSave, UneditedFixturesKeepChartPartBytes) {
  const char* const names[] = {"chart_area.xlsx",       "chart_bar.xlsx",    "chart_cache.xlsx",  "chart_chartex.xlsx",
                               "chart_chartsheet.xlsx", "chart_column.xlsx", "chart_combo.xlsx",  "chart_doughnut.xlsx",
                               "chart_data.xlsx",       "chart_filter.xlsx", "chart_format.xlsx", "chart_line.xlsx",
                               "chart_pie.xlsx",        "chart_pivot.xlsx",  "chart_refs.xlsx",   "chart_scatter.xlsx"};
  for (const char* name : names) {
    const Workbook wb = LoadFixture(name);
    const std::vector<std::string> paths = ChartPaths(wb);
    EXPECT_FALSE(paths.empty()) << name;
    auto refreshed = refreshed_chart_parts(wb);
    ASSERT_TRUE(static_cast<bool>(refreshed)) << name;
    for (const PassthroughPart& part : refreshed.value()) {
      ADD_FAILURE() << name << " refreshed " << part.path;
    }
    const Workbook saved = SaveAndReload(wb, false);
    for (const std::string& path : paths) {
      const PassthroughPart* before = io::ooxml::find_passthrough_part(wb, path);
      const PassthroughPart* after = io::ooxml::find_passthrough_part(saved, path);
      ASSERT_NE(before, nullptr) << name << " " << path;
      ASSERT_NE(after, nullptr) << name << " " << path;
      EXPECT_EQ(before->bytes, after->bytes) << name << " " << path;
    }
  }
}

TEST(ChartSave, UneditedXlsbKeepsChartPartBytes) {
  const Workbook wb = LoadFixture("chart_xlsb.xlsb");
  const std::vector<std::string> paths = ChartPaths(wb);
  EXPECT_FALSE(paths.empty());
  const Workbook saved = SaveAndReload(wb, true);
  for (const std::string& path : paths) {
    EXPECT_EQ(PartText(wb, path), PartText(saved, path)) << path;
  }
}

// A cell edit reaches the saved file only; the in-memory part keeps its old cache.
TEST(ChartSave, EditedCellRefreshesSavedCacheOnly) {
  for (const bool xlsb : {false, true}) {
    SCOPED_TRACE(xlsb ? "xlsb" : "xlsx");
    Workbook wb = LoadFixture(xlsb ? "chart_xlsb.xlsb" : "chart_column.xlsx");
    const std::vector<std::string> paths = ChartPaths(wb);
    ASSERT_FALSE(paths.empty());
    std::vector<std::string> before;
    for (const std::string& path : paths) {
      before.push_back(PartText(wb, path));
    }
    // Every numeric cell near the origin takes a marker no fixture cache holds.
    for (std::uint32_t row = 0; row < 12; ++row) {
      for (std::uint32_t col = 0; col < 8; ++col) {
        if (wb.sheet(0).resolve_cell_value(row, col).is_number()) {
          wb.sheet(0).set_cell_value(row, col, Value::number(424242));
        }
      }
    }
    const Workbook saved = SaveAndReload(wb, xlsb);
    bool any_refreshed = false;
    for (std::size_t i = 0; i < paths.size(); ++i) {
      EXPECT_EQ(PartText(wb, paths[i]), before[i]) << "in-memory part changed: " << paths[i];
      any_refreshed = any_refreshed || Contains(PartText(saved, paths[i]), "<c:v>424242</c:v>");
    }
    EXPECT_TRUE(any_refreshed);
  }
}

// Saving never recalculates: a dirty cell in a manual-calc book is written as stored.
TEST(ChartSave, ManualCalcDirtyCellWritesStoredValue) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::vector<std::string> paths = ChartPaths(wb);
  ASSERT_FALSE(paths.empty());
  wb.set_calc_mode(CalcMode::kManual);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 1, Value::number(5))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 1, Value::number(1))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 3, 1, "=B2*1000+777")));
  const Value stored = wb.sheet(0).resolve_cell_value(3, 1);
  const Workbook saved = SaveAndReload(wb, false);
  EXPECT_EQ(wb.sheet(0).resolve_cell_value(3, 1).is_number(), stored.is_number());
  for (const std::string& path : paths) {
    EXPECT_FALSE(Contains(PartText(saved, path), "<c:v>5777</c:v>")) << path;
  }
}

// Save time of a 100,000-point series, printed for the record.
TEST(ChartSave, HundredThousandPointChartSaveTime) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::vector<std::string> paths = ChartPaths(wb);
  ASSERT_FALSE(paths.empty());
  std::string text;
  std::string path;
  for (const std::string& candidate : paths) {
    if (Contains(PartText(wb, candidate), "Sheet1!$B$2:$B$5")) {
      path = candidate;
      text = PartText(wb, candidate);
    }
  }
  ASSERT_FALSE(path.empty());
  const std::string from = "Sheet1!$B$2:$B$5";
  for (std::size_t at = text.find(from); at != std::string::npos; at = text.find(from, at)) {
    text.replace(at, from.size(), "Sheet1!$B$2:$B$100001");
  }
  ASSERT_TRUE(
      static_cast<bool>(wb.replace_passthrough_part(path, std::vector<std::uint8_t>(text.begin(), text.end()))));
  for (std::uint32_t row = 1; row <= 100000; ++row) {
    wb.sheet(0).set_cell_value(row, 1, Value::number(static_cast<double>(row) * 1.5));
  }
  const auto start = std::chrono::steady_clock::now();
  auto written = io::write_ooxml(wb);
  const auto stop = std::chrono::steady_clock::now();
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(stop - start).count();
  std::printf("[ChartSave] 100000-point chart save: %lld ms\n", static_cast<long long>(ms));
  const Workbook saved = LoadBytes(written.value(), false);
  EXPECT_TRUE(Contains(PartText(saved, path), "<c:ptCount val=\"100000\"/>"));
}

}  // namespace
}  // namespace formulon
