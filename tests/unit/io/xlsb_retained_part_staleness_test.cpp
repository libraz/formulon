//
// Save invariant for XLSB retained parts (pivot table, pivot cache, styles):
// `write_xlsb_with_result` never writes a retained part whose model twin was
// mutated since load. A mutated pivot group is encoded from the model, or left
// out and counted when the bridges refuse it. Uses the same real Mac Excel 365
// fixture as `xlsb_pivot_fixture_test.cpp`
// (`tests/fixtures/excel/xlsb_pivot_base.xlsb`), which retains a pivot table,
// its cache, and a styles.bin part.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "styles.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath() {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_pivot_base.xlsb";
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
  const std::vector<std::uint8_t> bytes = ReadFileBytes(FixturePath());
  auto result_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_xlsb failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

TEST(XlsbRetainedPartStaleness, UnmodifiedRoundTripSavesCleanly) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
}

/// The parts of `package` whose names start with `prefix`, by name.
std::vector<std::pair<std::string, std::vector<std::uint8_t>>> PartsUnder(const std::vector<std::uint8_t>& package,
                                                                          const std::string& prefix) {
  std::vector<std::pair<std::string, std::vector<std::uint8_t>>> out;
  io::ZipReader zip;
  if (!zip.open(io::ByteSpan{package.data(), package.size()})) {
    ADD_FAILURE() << "package does not open";
    return out;
  }
  for (const std::string& name : zip.list_entries()) {
    if (name.compare(0, prefix.size(), prefix) == 0 && name.size() > 4U &&
        name.compare(name.size() - 4U, 4U, ".bin") == 0) {
      auto bytes = zip.read_entry(name);
      if (bytes) {
        out.emplace_back(name, std::move(bytes.value()));
      }
    }
  }
  return out;
}

Workbook Reload(const std::vector<std::uint8_t>& bytes) {
  auto result_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  return result_or ? std::move(result_or.value().workbook) : Workbook::create_empty();
}

TEST(XlsbRetainedPartStaleness, AMutatedPivotTableIsWrittenFromTheModel) {
  const std::vector<std::uint8_t> source = ReadFileBytes(FixturePath());
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  wb.sheet(0).mutable_pivot_tables()[0]->set_data_caption("Changed");

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const Workbook back = Reload(result.value().bytes);
  ASSERT_EQ(back.sheet(0).pivot_tables().size(), 1U);
  EXPECT_EQ(back.sheet(0).pivot_tables()[0]->data_caption(), "Changed");
  const auto retained = PartsUnder(source, "xl/pivotTables/");
  const auto written = PartsUnder(result.value().bytes, "xl/pivotTables/");
  ASSERT_EQ(retained.size(), 1U);
  ASSERT_EQ(written.size(), 1U);
  EXPECT_NE(written[0].second, retained[0].second);
}

TEST(XlsbRetainedPartStaleness, AMutatedPivotCacheIsWrittenFromTheModel) {
  const std::vector<std::uint8_t> source = ReadFileBytes(FixturePath());
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  ASSERT_FALSE(wb.mutable_pivot_caches()[0]->fields().empty());
  wb.mutable_pivot_caches()[0]->mutable_fields()[0].name = "Changed";

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const Workbook back = Reload(result.value().bytes);
  ASSERT_EQ(back.pivot_caches().size(), 1U);
  ASSERT_FALSE(back.pivot_caches()[0]->fields().empty());
  EXPECT_EQ(back.pivot_caches()[0]->fields()[0].name, "Changed");
  const auto retained = PartsUnder(source, "xl/pivotCache/pivotCacheDefinition");
  const auto written = PartsUnder(result.value().bytes, "xl/pivotCache/pivotCacheDefinition");
  ASSERT_EQ(retained.size(), 1U);
  ASSERT_EQ(written.size(), 1U);
  EXPECT_NE(written[0].second, retained[0].second);
}

// A mutated group the bridges refuse is left out whole and counted; its stale
// retained bytes are never written in its place.
TEST(XlsbRetainedPartStaleness, AMutatedPivotTheBridgesRefuseIsLeftOutAndCounted) {
  const Workbook unchanged = LoadFixture();
  auto baseline = io::xlsb::write_xlsb_with_result(unchanged);
  ASSERT_TRUE(static_cast<bool>(baseline)) << baseline.error().message;

  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  // No Excel-saved table carries alternative text, so the table bridge refuses it.
  wb.sheet(0).mutable_pivot_tables()[0]->set_alt_text("Changed");

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.deferred_feature_count,
            baseline.value().diagnostics.deferred_feature_count + 1U);
  EXPECT_TRUE(PartsUnder(result.value().bytes, "xl/pivotTables/").empty());
  EXPECT_TRUE(PartsUnder(result.value().bytes, "xl/pivotCache/").empty());
  const Workbook back = Reload(result.value().bytes);
  EXPECT_TRUE(back.sheet(0).pivot_tables().empty());
  EXPECT_TRUE(back.pivot_caches().empty());
}

// A stale styles.bin is regenerated from the model rather than failing the
// save, and the pivot parts still ride along as passthrough.
TEST(XlsbRetainedPartStaleness, MutatingAStyleRegeneratesStylesBin) {
  Workbook wb = LoadFixture();
  StylesTable& styles = wb.mutable_styles();
  ASSERT_FALSE(styles.fonts.empty());
  const double new_size = styles.fonts[0].size + 1.0;
  styles.fonts[0].size = new_size;

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  auto reread = io::xlsb::read_xlsb(io::ByteSpan{result.value().bytes.data(), result.value().bytes.size()});
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  const Workbook& back = reread.value().workbook;
  ASSERT_FALSE(back.styles().fonts.empty());
  EXPECT_EQ(back.styles().fonts[0].size, new_size);
  EXPECT_EQ(back.sheet(0).pivot_tables().size(), 1U);
}

// Cell values are not embedded in the retained pivot cache parts -- the
// cache is a snapshot the reader decoded into `PivotCache::records()`, a
// container entirely separate from `Sheet::rows()` -- so mutating a cell
// well outside the pivot's source range and rendered grid must not trip
// the freshness gate.
TEST(XlsbRetainedPartStaleness, MutatingAnUnrelatedCellValueStillSavesCleanly) {
  Workbook wb = LoadFixture();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 50, 50, Value::number(42.0))));

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
}

}  // namespace
}  // namespace formulon
