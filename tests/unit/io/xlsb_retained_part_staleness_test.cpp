//
// Fail-closed save invariant for XLSB retained parts (pivot table, pivot
// cache, styles): `write_xlsb_with_result` must never silently re-emit a
// retained part whose model twin has been mutated since load. Uses the
// same real Mac Excel 365 fixture as `xlsb_pivot_fixture_test.cpp`
// (`tests/fixtures/excel/xlsb_pivot_base.xlsb`), which round-trips a
// pivot table, its cache, and a styles.bin part all as passthrough.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
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

TEST(XlsbRetainedPartStaleness, MutatingThePivotTableFailsTheSave) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  wb.sheet(0).mutable_pivot_tables()[0]->set_data_caption("Changed");

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRetainedPartStale);
}

TEST(XlsbRetainedPartStaleness, MutatingThePivotCacheFailsTheSave) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  ASSERT_FALSE(wb.mutable_pivot_caches()[0]->fields().empty());
  wb.mutable_pivot_caches()[0]->mutable_fields()[0].name = "Changed";

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRetainedPartStale);
}

TEST(XlsbRetainedPartStaleness, MutatingAStyleFailsTheSave) {
  Workbook wb = LoadFixture();
  StylesTable& styles = wb.mutable_styles();
  ASSERT_FALSE(styles.fonts.empty());
  styles.fonts[0].size += 1.0;

  auto result = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRetainedPartStale);
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
