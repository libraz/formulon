// A save that writes modelled state from the model supersedes the reader's retained .xlsb original of it: the
// original is not counted as a lost part, and an .xlsb save gives the regenerated part the original's name.

#include <algorithm>
#include <cstdint>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

std::vector<std::string> PivotAndStyleParts(const std::vector<std::uint8_t>& package) {
  std::vector<std::string> out;
  for (const std::string& name : tc::ListParts(package)) {
    if (name.rfind("xl/pivotCache/", 0) == 0 || name.rfind("xl/pivotTables/", 0) == 0 || name == "xl/styles.bin") {
      out.push_back(name);
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

const char kBook[] = "/excel/xlsb_parity/pivot/formats_areas.xlsb";

TEST(XlsbRetainedOriginal, ARefreshedGroupAndItsStylesReplaceTheirOriginals) {
  const tc::Bytes source = tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + kBook);
  auto read = read_xlsb(SpanOf(source));
  ASSERT_TRUE(read.has_value()) << read.error().message;
  Workbook& wb = read.value().workbook;
  ASSERT_FALSE(wb.pivot_caches().empty());
  for (const auto& cache : wb.pivot_caches()) {
    const auto refreshed = wb.refresh_pivot_cache(cache->cache_id(), eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  }
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  EXPECT_EQ(saved.value().diagnostics.dropped_part_count, 0U);
  EXPECT_EQ(PivotAndStyleParts(saved.value().bytes), PivotAndStyleParts(source));
}

// The .xlsx save writes styles and pivots as XML, so the .bin originals are left out without counting.
TEST(XlsbRetainedOriginal, AnXlsxSaveDoesNotCountTheRetainedOriginals) {
  const tc::Bytes source = tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + kBook);
  auto read = read_xlsb(SpanOf(source));
  ASSERT_TRUE(read.has_value()) << read.error().message;
  auto saved = write_ooxml_with_result(read.value().workbook);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.deferred_feature_count, 0U);
  EXPECT_EQ(saved.value().diagnostics.dropped_part_count, 0U);
}

// A part the reader kept without modelling it is lost when a generated part takes its path, and counted.
TEST(XlsbRetainedOriginal, AnUnmodelledPartUnderAGeneratedPathIsCounted) {
  Workbook wb = Workbook::create();
  wb.set_passthrough_parts({PassthroughPart(
      "xl/workbook.bin", "application/vnd.ms-excel.sheet.binary.macroEnabled.main", std::vector<std::uint8_t>{0U})});
  auto saved = write_xlsb_with_result(wb);
  ASSERT_TRUE(saved.has_value()) << saved.error().message;
  EXPECT_EQ(saved.value().diagnostics.dropped_part_count, 1U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
