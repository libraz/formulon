//
// Pins the XLSB pivot gap against a real Mac Excel 365-produced `.xlsb`
// (`tests/fixtures/excel/xlsb_pivot_base.xlsb`).
//
// The XLSB reader decodes no pivot records, so a pivot that lives in a
// `.xlsb` never reaches the evaluated pivot model and GETPIVOTDATA over
// it resolves against nothing. That is a divergence from Excel rather
// than a formula defect, and this fixture is what makes it measurable:
// Excel evaluated the same three formulas before saving, so the file
// carries both answers at once — Excel's in the cached cell values, ours
// from a recalc.
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
// When XLSB pivot record decoding lands, the recalc expectations below
// become the cached values and this file collapses into a plain equality
// check. Until then it is the only place the measured Excel answers are
// written down.

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
// (a) The package really does carry a pivot, and we really do drop it.
// ---------------------------------------------------------------------------

TEST(XlsbPivotFixture, PivotPartsAreNotDecodedIntoTheModel) {
  Workbook wb = LoadFixture();
  ASSERT_EQ(wb.sheet_count(), 1U);
  // The `.xlsb` holds `pivotCacheDefinition1.bin`, `pivotCacheRecords1.bin`
  // and `pivotTable1.bin`; none of the three is decoded, so nothing is
  // wired onto the workbook. Asserting zero (rather than skipping the
  // check) is what turns "not implemented" into a fact a future decoder
  // has to overturn deliberately.
  EXPECT_TRUE(wb.pivot_caches().empty());
  EXPECT_TRUE(wb.sheet(0).pivot_tables().empty());
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
// (c) The divergence itself: recalc replaces every answer with #REF!.
// ---------------------------------------------------------------------------

TEST(XlsbPivotFixture, RecalcTurnsGetPivotDataIntoRefErrorWhileExcelAnswers) {
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
    ASSERT_TRUE(after.is_error()) << "row " << probe.row << ": expected the pivot lookup to fail, got a value";
    EXPECT_EQ(after.as_error(), ErrorCode::Ref)
        << "row " << probe.row << ": Excel answered " << probe.excel_answer << " for this probe";
  }
}

}  // namespace
}  // namespace formulon
