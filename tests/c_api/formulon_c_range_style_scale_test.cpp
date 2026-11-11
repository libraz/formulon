// A hyperlink over A1:Z100000 (2.6M blank formatted cells) holds no cell per
// blank address in memory, and survives an .xlsx and an .xlsb save and load
// with every coordinate, its corners' formats and a bounded footprint. Runs
// in its own SLOW executable so the peak-RSS reading covers this run alone.

#include <sys/resource.h>

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

constexpr std::size_t kMiB = 1024U * 1024U;
constexpr std::uint64_t kGiB = 1024ULL * 1024ULL * 1024ULL;
constexpr std::uint32_t kLastRow = 99999U;
constexpr std::uint32_t kLastCol = 25U;

std::size_t MemoryBytes(const fm_workbook_t* wb) {
  std::size_t bytes = 0U;
  EXPECT_EQ(fm_workbook_get_memory_usage(wb, &bytes), 0);
  return bytes;
}

std::size_t CellCount(const fm_workbook_t* wb) {
  std::size_t count = 0U;
  EXPECT_EQ(fm_sheet_get_cell_count(wb, 0, &count), 0);
  return count;
}

std::uint32_t EffectiveXf(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_effective_style_info_t style = FM_STRUCT_INIT(fm_effective_style_info_t);
  EXPECT_EQ(fm_sheet_get_effective_style(wb, 0, row, col, &style), 0);
  return style.xf_index;
}

// Peak resident set size in bytes (`ru_maxrss` is bytes on macOS, KiB on Linux).
std::uint64_t PeakRssBytes() {
  struct rusage usage {};
  getrusage(RUSAGE_SELF, &usage);
  const auto peak = static_cast<std::uint64_t>(usage.ru_maxrss);
#if defined(__APPLE__)
  return peak;
#else
  return peak * 1024U;
#endif
}

// The coordinates whose format the round trip must keep: the four corners,
// A1 and one inner cell.
const std::vector<std::pair<std::uint32_t, std::uint32_t>>& Probes() {
  static const std::vector<std::pair<std::uint32_t, std::uint32_t>> probes = {
      {0U, 0U}, {0U, kLastCol}, {kLastRow, 0U}, {kLastRow, kLastCol}, {50000U, 13U}};
  return probes;
}

class FormulonCApiRangeStyleScale : public ::testing::TestWithParam<fm_workbook_format_t> {};

TEST_P(FormulonCApiRangeStyleScale, HyperlinkOverA1ToZ100000SurvivesSaveAndLoad) {
  WorkbookGuard empty;
  ASSERT_EQ(fm_workbook_create_default(&empty.handle), 0);
  const std::size_t baseline = MemoryBytes(empty.handle);

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::size_t before = MemoryBytes(wb.handle);
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = fm_cell_rect_t{0U, 0U, kLastRow, kLastCol};
  link.target = "https://example.com";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb.handle, 0, &link), 0);
  EXPECT_LT(MemoryBytes(wb.handle) - before, 32U * kMiB);
  EXPECT_EQ(CellCount(wb.handle), 2600000U);
  std::vector<std::uint32_t> xfs;
  for (const auto& [row, col] : Probes()) {
    xfs.push_back(EffectiveXf(wb.handle, row, col));
  }
  EXPECT_NE(xfs[4], 0U);

  const std::uint64_t rss_before = PeakRssBytes();
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, GetParam(), &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  EXPECT_LT(PeakRssBytes() - rss_before, kGiB);

  EXPECT_EQ(CellCount(loaded.handle), 2600000U);
  for (std::size_t i = 0; i < Probes().size(); ++i) {
    EXPECT_EQ(EffectiveXf(loaded.handle, Probes()[i].first, Probes()[i].second), xfs[i]) << "probe " << i;
  }
  EXPECT_LT(MemoryBytes(loaded.handle) - baseline, 32U * kMiB);
}

INSTANTIATE_TEST_SUITE_P(Formats, FormulonCApiRangeStyleScale,
                         ::testing::Values(FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB),
                         [](const ::testing::TestParamInfo<fm_workbook_format_t>& param_info) {
                           return param_info.param == FM_WORKBOOK_FORMAT_XLSX ? std::string("Xlsx")
                                                                              : std::string("Xlsb");
                         });

}  // namespace
