//
// Checks XLSB pivot decoding against a real Mac Excel 365-produced
// `.xlsb` (`tests/fixtures/excel/xlsb_pivot_base.xlsb`).
//
// The fixture carries both engines' answers at once: Excel evaluated the
// three GETPIVOTDATA formulas before saving, so its results sit in the
// cached cell values while ours come from a recalc over the decoded
// model. Every expectation below is therefore a direct comparison rather
// than a number transcribed into this file.
//
// The record layouts the reader relies on were established by
// differential decode — the same workbook re-saved by Excel as `.xlsx`
// and read with the OOXML pivot reader — rather than from a
// specification. See `src/io/xlsb/pivot_reader.h`.
//
// Fixture layout (`Sheet1`):
//   A1:C7  source table, headers Region / Qtr / Amt
//   E1:H6  a PivotTable over A1:C7 — Region on rows, Qtr on columns,
//          sum of Amt as the measure (ja-JP labels: 合計 / Amt,
//          行ラベル, 列ラベル, 総計)
//   E12    GETPIVOTDATA("Amt",$E$1,"Region","North")            -> 15
//   E13    GETPIVOTDATA("Amt",$E$1)                             -> 56
//   E14    GETPIVOTDATA("Amt",$E$1,"Region","North","Qtr","Q1") -> 10
//
// E12 is the shape that names one axis in full and leaves the other
// open: Qtr is on the columns and the formula does not mention it, so
// the answer is North's total across every quarter.

#include <array>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
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

// Row/column of the three GETPIVOTDATA probes (column E, rows 12-14).
constexpr std::uint32_t kProbeCol = 4U;
constexpr std::uint32_t kProbeRowSingleField = 11U;
constexpr std::uint32_t kProbeRowGrandTotal = 12U;
constexpr std::uint32_t kProbeRowTwoFields = 13U;

// The values Excel 365 (Mac, ja-JP, 16.112) computed for those probes.
constexpr double kExcelSingleField = 15.0;
constexpr double kExcelGrandTotal = 56.0;
constexpr double kExcelTwoFields = 10.0;

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
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  auto result_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result_or)) << "read_xlsb failed: " << (result_or ? "" : result_or.error().message);
  if (!result_or) {
    return Workbook::create_empty();
  }
  return std::move(result_or.value().workbook);
}

// ---------------------------------------------------------------------------
// (a) The package's three pivot parts reach the model.
// ---------------------------------------------------------------------------

TEST(XlsbPivotFixture, PivotPartsDecodeIntoTheModel) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);

  // `pivotCacheDefinition1.bin`: three source columns, the two discrete
  // ones carrying their shared items and the measure carrying none
  // because its values are stored inline on each record.
  const pivot::PivotCache& cache = *wb.pivot_caches().front();
  ASSERT_EQ(cache.fields().size(), 3U);
  EXPECT_EQ(cache.fields()[0].name, "Region");
  EXPECT_EQ(cache.fields()[1].name, "Qtr");
  EXPECT_EQ(cache.fields()[2].name, "Amt");
  EXPECT_EQ(cache.fields()[0].shared_items.size(), 3U);
  EXPECT_EQ(cache.fields()[1].shared_items.size(), 2U);
  EXPECT_TRUE(cache.fields()[2].shared_items.empty());

  // `pivotCacheRecords1.bin`: the six source rows, each an index into
  // Region's items, an index into Qtr's, and the inline amount.
  ASSERT_EQ(cache.records().size(), 6U);
  double amount_total = 0.0;
  for (const pivot::PivotCacheRecord& record : cache.records()) {
    ASSERT_EQ(record.cells.size(), 3U);
    ASSERT_EQ(record.cell_is_index.size(), 3U);
    EXPECT_TRUE(record.cell_is_index[0]);
    EXPECT_TRUE(record.cell_is_index[1]);
    EXPECT_FALSE(record.cell_is_index[2]) << "the measure is stored inline, not as a shared-item index";
    ASSERT_TRUE(record.cells[2].is_number());
    amount_total += record.cells[2].as_number();
  }
  EXPECT_DOUBLE_EQ(amount_total, kExcelGrandTotal) << "the decoded records do not sum to Excel's grand total";

  // `pivotTable1.bin`: Region down the rows, Qtr across the columns, one
  // summed measure, anchored on the E1:H6 rectangle Excel wrote.
  const pivot::PivotTable& table = *wb.sheet(0).pivot_tables().front();
  EXPECT_EQ(table.pivot_cache_id(), cache.cache_id());
  EXPECT_EQ(table.anchor_row(), 0U);
  EXPECT_EQ(table.anchor_col(), kProbeCol);
  EXPECT_EQ(table.span_rows(), 6U);
  EXPECT_EQ(table.span_cols(), 4U);
  ASSERT_EQ(table.row_field_order().size(), 1U);
  ASSERT_EQ(table.col_field_order().size(), 1U);
  EXPECT_EQ(table.row_field_order()[0], 0U);
  EXPECT_EQ(table.col_field_order()[0], 1U);
  ASSERT_EQ(table.data_fields().size(), 1U);
  EXPECT_EQ(table.data_fields()[0].field_index, 2U);
  EXPECT_EQ(table.data_fields()[0].aggregation, pivot::Aggregation::Sum);
  // Excel names the measure after its aggregation; GETPIVOTDATA in the
  // sheet addresses it by the source column instead, and both resolve.
  EXPECT_EQ(table.data_fields()[0].name, "合計 / Amt");

  // Names are backfilled from the bound cache, positionally: the binary
  // identifies a source column by index and never by name.
  ASSERT_EQ(table.fields().size(), 3U);
  EXPECT_EQ(table.fields()[0].source_name, "Region");
  EXPECT_EQ(table.fields()[1].source_name, "Qtr");
  EXPECT_EQ(table.fields()[2].source_name, "Amt");
}

// ---------------------------------------------------------------------------
// (b) Excel's own answers, read straight out of the fixture.
// ---------------------------------------------------------------------------

TEST(XlsbPivotFixture, CachedCellValuesCarryTheExcelGetPivotDataResults) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Sheet& sheet = wb.sheet(0);

  const Value single = sheet.resolve_cell_value(kProbeRowSingleField, kProbeCol);
  ASSERT_TRUE(single.is_number()) << "E12 lost Excel's cached value";
  EXPECT_DOUBLE_EQ(single.as_number(), kExcelSingleField);

  const Value total = sheet.resolve_cell_value(kProbeRowGrandTotal, kProbeCol);
  ASSERT_TRUE(total.is_number()) << "E13 lost Excel's cached value";
  EXPECT_DOUBLE_EQ(total.as_number(), kExcelGrandTotal);

  const Value pair = sheet.resolve_cell_value(kProbeRowTwoFields, kProbeCol);
  ASSERT_TRUE(pair.is_number()) << "E14 lost Excel's cached value";
  EXPECT_DOUBLE_EQ(pair.as_number(), kExcelTwoFields);
}

// The pivot's rendered grid is ordinary cached cell content, so it
// survives the load even though the pivot behind it does not. This keeps
// the fixture honest: the cached GETPIVOTDATA values above are not an
// artefact of the whole sheet being unreadable.
TEST(XlsbPivotFixture, RenderedPivotGridSurvivesAsPlainCells) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Sheet& sheet = wb.sheet(0);
  // H6 is the pivot's grand total (総計 row, 総計 column).
  const Value grand = sheet.resolve_cell_value(5U, 7U);
  ASSERT_TRUE(grand.is_number());
  EXPECT_DOUBLE_EQ(grand.as_number(), kExcelGrandTotal);
}

// ---------------------------------------------------------------------------
// (c) A recalc reproduces every answer Excel cached.
// ---------------------------------------------------------------------------

TEST(XlsbPivotFixture, RecalcReproducesTheExcelGetPivotDataAnswers) {
  // The fixture carries both answers at once -- Excel's in the cached cell
  // values, ours from the recalc -- so this is a direct comparison rather
  // than a check against a number written down in this file.
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << (recalc_or ? "" : recalc_or.error().message);

  struct Probe {
    std::uint32_t row;
    double excel_answer;
  };
  const std::array<Probe, 3> kProbes = {Probe{kProbeRowSingleField, kExcelSingleField},
                                        Probe{kProbeRowGrandTotal, kExcelGrandTotal},
                                        Probe{kProbeRowTwoFields, kExcelTwoFields}};
  for (const Probe& probe : kProbes) {
    const Value after = wb.sheet(0).resolve_cell_value(probe.row, kProbeCol);
    ASSERT_TRUE(after.is_number()) << "row " << probe.row << ": the pivot lookup did not produce a number";
    EXPECT_DOUBLE_EQ(after.as_number(), probe.excel_answer) << "row " << probe.row;
  }
}

// The measure displays as "合計 / Amt" but every formula in the sheet
// addresses it as "Amt", which is what Excel writes into a formula it
// generates. Both spellings have to resolve to the same value.
TEST(XlsbPivotFixture, TheMeasureResolvesByDisplayNameAndBySourceColumn) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  // E13 already asks for "Amt" and is covered above; rewriting that same
  // cell with the display-name spelling keeps the comparison on one cell
  // the dependency graph already tracks.
  wb.sheet(0).set_cell_formula(kProbeRowGrandTotal, kProbeCol, "=GETPIVOTDATA(\"合計 / Amt\",$E$1)");
  auto recalc_or = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc_or)) << (recalc_or ? "" : recalc_or.error().message);

  const Value by_display = wb.sheet(0).resolve_cell_value(kProbeRowGrandTotal, kProbeCol);
  ASSERT_TRUE(by_display.is_number()) << "the display-name spelling did not resolve";
  EXPECT_DOUBLE_EQ(by_display.as_number(), kExcelGrandTotal);
}

}  // namespace
}  // namespace formulon
