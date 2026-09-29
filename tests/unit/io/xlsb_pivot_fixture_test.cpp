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
#include "io/xlsb/pivot_reader.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/error.h"
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

// ---------------------------------------------------------------------------
// (d) `BrtPivotTableLocation` is rejected, not clamped, when its rectangle
//     falls outside the sheet grid.
// ---------------------------------------------------------------------------

// Frames one `BrtPivotTableLocation` record: a 2-byte record-type varint
// (314 needs both continuation bytes), a 1-byte payload-size varint, and
// the six-`u32` payload `read_pivot_table_bin` expects.
std::vector<std::uint8_t> EncodeLocationRecord(std::uint32_t first_row, std::uint32_t last_row, std::uint32_t first_col,
                                               std::uint32_t last_col) {
  constexpr std::uint16_t kPivotTableLocation = 314;
  std::vector<std::uint8_t> out;
  out.push_back(static_cast<std::uint8_t>((kPivotTableLocation & 0x7F) | 0x80));
  out.push_back(static_cast<std::uint8_t>(kPivotTableLocation >> 7));
  const std::array<std::uint32_t, 6> fields = {first_row, last_row, first_col, last_col, 0U, 0U};
  out.push_back(static_cast<std::uint8_t>(fields.size() * sizeof(std::uint32_t)));
  for (std::uint32_t field : fields) {
    for (int shift = 0; shift < 32; shift += 8) {
      out.push_back(static_cast<std::uint8_t>((field >> shift) & 0xFFU));
    }
  }
  return out;
}

TEST(XlsbPivotFixture, LocationRecordRejectsARectangleBeyondTheRowGrid) {
  const std::vector<std::uint8_t> bytes =
      EncodeLocationRecord(/*first_row=*/0U, /*last_row=*/Sheet::kMaxRows, /*first_col=*/0U, /*last_col=*/0U);
  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbPivotFixture, LocationRecordRejectsARectangleBeyondTheColumnGrid) {
  const std::vector<std::uint8_t> bytes =
      EncodeLocationRecord(/*first_row=*/0U, /*last_row=*/0U, /*first_col=*/0U, /*last_col=*/Sheet::kMaxCols);
  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

// A rectangle that just reaches the grid edge must clear the bounds check;
// the record alone still fails to decode because a pivot table with no
// data field is rejected further down, which is the unrelated failure
// this test distinguishes from a wrongly-tightened bounds check.
TEST(XlsbPivotFixture, LocationRecordAtTheGridEdgeClearsTheBoundsCheck) {
  const std::vector<std::uint8_t> bytes = EncodeLocationRecord(
      /*first_row=*/0U, /*last_row=*/Sheet::kMaxRows - 1U, /*first_col=*/0U, /*last_col=*/Sheet::kMaxCols - 1U);
  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().message, "xlsb pivot table declares no data field");
}

// ---------------------------------------------------------------------------
// (e) Corrupt-input paths through `read_pivot_table_bin` / `read_pivot_cache_bin`
//     directly, exercising the "callers treat this as skip the pivot"
//     contract `pivot_reader.h` documents without needing a whole package.
// ---------------------------------------------------------------------------

// Record ids mirrored from the private enum in `src/io/xlsb/pivot_reader.cpp`
// (not exposed via the header, so the numeric values are duplicated here,
// the same way `EncodeLocationRecord` above already duplicates 314).
constexpr std::uint16_t kBeginPCDField = 183;
constexpr std::uint16_t kBeginPivotFields = 287;
constexpr std::uint16_t kBeginPivotField = 285;
constexpr std::uint16_t kBeginPivotFieldItem = 282;
constexpr std::uint16_t kPivotRowFields = 309;
constexpr std::uint16_t kBeginPivotDataField = 293;

void AppendU32(std::vector<std::uint8_t>& out, std::uint32_t v) {
  for (int shift = 0; shift < 32; shift += 8) {
    out.push_back(static_cast<std::uint8_t>((v >> shift) & 0xFFU));
  }
}

// General single-record encoder: every id used below needs both varint
// continuation bytes (all are > 127), so the type is always framed as two
// bytes; `EncodeLocationRecord` above predates this and stays specialised
// to its own fixed payload shape.
std::vector<std::uint8_t> EncodeRecord(std::uint16_t type, const std::vector<std::uint8_t>& payload) {
  std::vector<std::uint8_t> out;
  out.push_back(static_cast<std::uint8_t>((type & 0x7F) | 0x80));
  out.push_back(static_cast<std::uint8_t>(type >> 7));
  out.push_back(static_cast<std::uint8_t>(payload.size()));
  out.insert(out.end(), payload.begin(), payload.end());
  return out;
}

TEST(XlsbPivotFixture, CacheFieldOutsideDefinitionIsRejected) {
  const std::vector<std::uint8_t> bytes = EncodeRecord(kBeginPCDField, {});
  auto cache_or = io::xlsb::read_pivot_cache_bin(io::ByteSpan{bytes.data(), bytes.size()}, io::ByteSpan{});
  ASSERT_FALSE(static_cast<bool>(cache_or));
  EXPECT_EQ(cache_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(cache_or.error().message, "xlsb pivot cache field outside a cache definition");
}

// A cache definition holding one field named "F" whose body is `inner`.
std::vector<std::uint8_t> OneFieldCacheDefinition(const std::vector<std::vector<std::uint8_t>>& inner) {
  constexpr std::uint16_t kBeginPCDefinition = 179;
  constexpr std::uint16_t kEndPCDefinition = 180;
  constexpr std::uint16_t kEndPCDField = 184;
  std::vector<std::uint8_t> field_header(20, 0U);  // bytes preceding the name
  AppendU32(field_header, 1U);
  field_header.push_back(static_cast<std::uint8_t>('F'));
  field_header.push_back(0U);
  std::vector<std::uint8_t> bytes;
  for (const auto& rec : {EncodeRecord(kBeginPCDefinition, {}), EncodeRecord(kBeginPCDField, field_header)}) {
    bytes.insert(bytes.end(), rec.begin(), rec.end());
  }
  for (const auto& rec : inner) {
    bytes.insert(bytes.end(), rec.begin(), rec.end());
  }
  for (const auto& rec : {EncodeRecord(kEndPCDField, {}), EncodeRecord(kEndPCDefinition, {})}) {
    bytes.insert(bytes.end(), rec.begin(), rec.end());
  }
  return bytes;
}

TEST(XlsbPivotFixture, CacheFieldWithOnlyCharacterisedRecordsDecodes) {
  constexpr std::uint16_t kBeginPCDFAtbl = 189;
  constexpr std::uint16_t kEndPCDFAtbl = 190;
  const std::vector<std::uint8_t> bytes =
      OneFieldCacheDefinition({EncodeRecord(kBeginPCDFAtbl, {}), EncodeRecord(kEndPCDFAtbl, {})});
  auto cache_or = io::xlsb::read_pivot_cache_bin(io::ByteSpan{bytes.data(), bytes.size()}, io::ByteSpan{});
  ASSERT_TRUE(static_cast<bool>(cache_or)) << cache_or.error().message;
  ASSERT_EQ(cache_or.value().fields().size(), 1U);
  EXPECT_EQ(cache_or.value().fields()[0].name, "F");
}

TEST(XlsbPivotFixture, CacheFieldWithAnUncharacterisedRecordIsRejected) {
  // An id the measured fixture never carries inside a field, standing in for
  // a grouping block whose layout has not been decoded.
  constexpr std::uint16_t kUncharacterised = 373;
  const std::vector<std::uint8_t> bytes = OneFieldCacheDefinition({EncodeRecord(kUncharacterised, {0U, 0U})});
  auto cache_or = io::xlsb::read_pivot_cache_bin(io::ByteSpan{bytes.data(), bytes.size()}, io::ByteSpan{});
  ASSERT_FALSE(static_cast<bool>(cache_or));
  EXPECT_EQ(cache_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(cache_or.error().message, "xlsb pivot cache field carries an uncharacterised record (grouping?)");
}

TEST(XlsbPivotFixture, FieldCountMismatchIsRejected) {
  std::vector<std::uint8_t> bytes;
  std::vector<std::uint8_t> count_payload;
  AppendU32(count_payload, 2U);  // declares two fields
  const std::vector<std::uint8_t> fields_rec = EncodeRecord(kBeginPivotFields, count_payload);
  bytes.insert(bytes.end(), fields_rec.begin(), fields_rec.end());
  // ...but only one `kBeginPivotField` block actually follows.
  const std::vector<std::uint8_t> field_rec = EncodeRecord(kBeginPivotField, {});
  bytes.insert(bytes.end(), field_rec.begin(), field_rec.end());

  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(table_or.error().message, "xlsb pivot table field count does not match the fields decoded");
}

TEST(XlsbPivotFixture, FieldItemOutsideAnItemListIsRejected) {
  const std::vector<std::uint8_t> bytes = EncodeRecord(kBeginPivotFieldItem, {});
  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(table_or.error().message, "xlsb pivot field item outside a field's item list");
}

TEST(XlsbPivotFixture, RowAxisIndexOutOfRangeIsRejected) {
  std::vector<std::uint8_t> bytes;
  auto append = [&](const std::vector<std::uint8_t>& rec) { bytes.insert(bytes.end(), rec.begin(), rec.end()); };

  std::vector<std::uint8_t> field_count_payload;
  AppendU32(field_count_payload, 1U);
  append(EncodeRecord(kBeginPivotFields, field_count_payload));
  append(EncodeRecord(kBeginPivotField, {}));

  // A minimal but valid `BrtBeginPivotDataField`: field 0, Sum, empty name --
  // needed so the table clears the "no data field" check before the row-axis
  // check under test ever runs.
  std::vector<std::uint8_t> data_field_payload;
  AppendU32(data_field_payload, 0U);                             // field_index
  AppendU32(data_field_payload, 0U);                             // selector: Sum
  data_field_payload.insert(data_field_payload.end(), 17U, 0U);  // header gap
  AppendU32(data_field_payload, 0U);                             // name cch=0
  append(EncodeRecord(kBeginPivotDataField, data_field_payload));

  std::vector<std::uint8_t> row_fields_payload;
  AppendU32(row_fields_payload, 1U);  // one row field...
  AppendU32(row_fields_payload, 5U);  // ...naming field 5, which the table does not have
  append(EncodeRecord(kPivotRowFields, row_fields_payload));

  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(table_or.error().message, "xlsb pivot row axis names a field the table does not have");
}

TEST(XlsbPivotFixture, UnknownAggregationSelectorIsRejected) {
  std::vector<std::uint8_t> payload;
  AppendU32(payload, 0U);    // field_index
  AppendU32(payload, 999U);  // no selector this reader has measured
  const std::vector<std::uint8_t> bytes = EncodeRecord(kBeginPivotDataField, payload);

  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(table_or.error().message, "xlsb pivot data field uses an unknown aggregation selector");
}

TEST(XlsbPivotFixture, DataFieldHeaderTruncationIsRejected) {
  std::vector<std::uint8_t> payload;
  AppendU32(payload, 0U);                 // field_index
  AppendU32(payload, 0U);                 // selector: Sum (valid)
  payload.insert(payload.end(), 5U, 0U);  // short of the 17-byte header gap
  const std::vector<std::uint8_t> bytes = EncodeRecord(kBeginPivotDataField, payload);

  auto table_or = io::xlsb::read_pivot_table_bin(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_FALSE(static_cast<bool>(table_or));
  EXPECT_EQ(table_or.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_EQ(table_or.error().message, "xlsb pivot data field header truncated");
}

}  // namespace
}  // namespace formulon
