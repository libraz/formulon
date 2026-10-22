// Pivot style and advanced filters in XLSB: BrtTableStyleClient and the
// SXFILTERS block of a pivot table part decode into the same
// `PivotStyleInfo` and filter lists an `.xlsx` pivot's
// `<pivotTableStyleInfo>` and `<filters>` produce. The XLSB writer has no
// pivot output, so the part itself stays passthrough and is checked fresh
// against the decoded model on save.

#include <cstdint>
#include <map>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_table_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pivot/pivot_table.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kPivotPart = "xl/pivotTables/pivotTable1.bin";
constexpr std::uint16_t kBrtTableStyleClient = 513;

// GETPIVOTDATA probes of the fixture (see xlsb_pivot_fixture_test.cpp).
constexpr std::uint32_t kProbeCol = 4U;
constexpr std::uint32_t kProbeRowNorth = 11U;
constexpr std::uint32_t kProbeRowGrandTotal = 12U;

std::vector<std::uint8_t> Fixture() {
  return ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_pivot_base.xlsb");
}

std::vector<std::uint8_t> Entry(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry)) << path;
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// One BrtBeginSXFILTER .. BrtEndSXFilter block over field 0 whose AutoFilter
/// is `criterion` (framed records inside one filter column).
std::vector<std::uint8_t> SxFilter(std::uint32_t type, std::uint32_t id, std::uint32_t measure,
                                   const std::string& value1, const std::vector<std::uint8_t>& criterion,
                                   const std::vector<std::uint8_t>& trailer = {}) {
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, 0U);           // isxvd
  emit_u32(payload, 0xFFFFFFFFU);  // isxvdMProp
  emit_u32(payload, type);         // sxft
  emit_u32(payload, 0xFFFFFFFFU);  // evalOrder
  emit_u32(payload, id);           // dwId
  emit_u32(payload, measure);      // isxdiMeasure
  emit_u32(payload, 0xFFFFFFFFU);  // isxthMeasure
  emit_u16(payload, value1.empty() ? 0U : 0x0004U);
  if (!value1.empty()) {
    emit_xlwidestring(payload, value1);
  }
  emit_record(out, 601, payload);
  payload.clear();
  emit_rfx(payload, MergeRange{0, 0, 0, 0});
  emit_record(out, 161, payload);
  payload.clear();
  emit_u32(payload, 0U);
  emit_u16(payload, 0U);
  emit_record(out, 163, payload);
  out.insert(out.end(), criterion.begin(), criterion.end());
  emit_record(out, 164, ByteSpan{});
  emit_record(out, 162, ByteSpan{});
  out.insert(out.end(), trailer.begin(), trailer.end());
  emit_record(out, 602, ByteSpan{});
  return out;
}

std::vector<std::uint8_t> ValuesCriterion(const std::string& value) {
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, 0U);
  emit_u32(payload, 0U);
  emit_record(out, 165, payload);
  payload.clear();
  emit_xlwidestring(payload, value);
  emit_record(out, 167, payload);
  emit_record(out, 166, ByteSpan{});
  return out;
}

std::vector<std::uint8_t> Top10Criterion(double n) {
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_u8(payload, 0x05U);  // fTop | fApplied
  emit_double(payload, n);
  emit_double(payload, n);
  emit_record(out, 170, payload);
  return out;
}

/// The fixture with `filters` (SXFILTER blocks) inserted after the pivot's
/// BrtTableStyleClient.
std::vector<std::uint8_t> WithFilters(const std::vector<std::vector<std::uint8_t>>& filters) {
  const std::vector<std::uint8_t> package = Fixture();
  const std::vector<std::uint8_t> part = Entry(package, kPivotPart);
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(part)) {
    append_record(patched, part, rec);
    if (rec.type == kBrtTableStyleClient) {
      std::vector<std::uint8_t> payload;
      emit_u32(payload, static_cast<std::uint32_t>(filters.size()));
      emit_record(patched, 599, payload);
      for (const std::vector<std::uint8_t>& filter : filters) {
        patched.insert(patched.end(), filter.begin(), filter.end());
      }
      emit_record(patched, 600, ByteSpan{});
    }
  }
  return RepackWithEntries(package, {{kPivotPart, patched}});
}

TEST(XlsbPivotFidelity, StyleClientDecodesAsPivotTableStyleInfo) {
  const Workbook wb = Load(Fixture());
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  const auto& style = wb.sheet(0).pivot_tables()[0]->style_info();
  ASSERT_TRUE(style.has_value());
  // Excel writes the same state to `.xlsx` as name="PivotStyleLight16"
  // showRowHeaders="1" showColHeaders="1" showRowStripes="0"
  // showColStripes="0" showLastColumn="1".
  EXPECT_EQ(style->name, "PivotStyleLight16");
  EXPECT_TRUE(style->show_row_headers);
  EXPECT_TRUE(style->show_col_headers);
  EXPECT_FALSE(style->show_row_stripes);
  EXPECT_FALSE(style->show_col_stripes);
  EXPECT_TRUE(style->show_last_column);
}

TEST(XlsbPivotFidelity, CaptionFilterDecodesLikeTheXlsxElement) {
  Workbook wb = Load(WithFilters({SxFilter(4U, 1U, 0xFFFFFFFFU, "North", ValuesCriterion("North"))}));
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  const pivot::PivotTable& table = *wb.sheet(0).pivot_tables()[0];
  ASSERT_EQ(table.retained_filters().size(), 1U);
  EXPECT_EQ(table.authored_caption_filters().size(), 1U);

  // The same filter as Windows Excel 365 wrote it to `.xlsx`.
  const std::vector<std::uint8_t> xlsx =
      ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_caption_filter.xlsx");
  auto twin = read_pivot_table_definition(Entry(xlsx, "xl/pivotTables/pivotTable1.xml"));
  ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
  ASSERT_EQ(twin.value().retained_filters().size(), 1U);
  EXPECT_EQ(table.retained_filters()[0].xml, twin.value().retained_filters()[0].xml);

  // The filter applies: only North is left under the grand total.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value north = wb.sheet(0).resolve_cell_value(kProbeRowNorth, kProbeCol);
  const Value grand = wb.sheet(0).resolve_cell_value(kProbeRowGrandTotal, kProbeCol);
  ASSERT_TRUE(north.is_number());
  ASSERT_TRUE(grand.is_number());
  EXPECT_DOUBLE_EQ(grand.as_number(), north.as_number());
}

TEST(XlsbPivotFidelity, TopCountFilterBecomesAnActiveFilter) {
  const Workbook wb = Load(WithFilters({SxFilter(1U, 2U, 0U, "", Top10Criterion(2.0))}));
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  const pivot::PivotTable& table = *wb.sheet(0).pivot_tables()[0];
  ASSERT_EQ(table.active_filters().size(), 1U);
  EXPECT_EQ(table.active_filters()[0].type, pivot::FilterType::ValueTop10);
  EXPECT_TRUE(table.retained_filters().empty());
}

TEST(XlsbPivotFidelity, FilteredPivotStaysFreshOnAnXlsbSave) {
  const std::vector<std::uint8_t> source =
      WithFilters({SxFilter(4U, 1U, 0xFFFFFFFFU, "North", ValuesCriterion("North"))});
  const Workbook wb = Load(source);
  ASSERT_EQ(wb.sheet(0).pivot_tables().size(), 1U);
  // The retained part matches the decoded model, so the save keeps it.
  auto xlsb = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  EXPECT_EQ(Entry(xlsb.value(), kPivotPart), Entry(source, kPivotPart));
  const Workbook back = Load(xlsb.value());
  ASSERT_EQ(back.sheet(0).pivot_tables().size(), 1U);
  EXPECT_EQ(back.sheet(0).pivot_tables()[0]->retained_filters().size(), 1U);
}

TEST(XlsbPivotFidelity, UnmodelledFilterLeavesThePivotAlone) {
  std::vector<std::uint8_t> frt;
  emit_record(frt, 35, std::vector<std::uint8_t>{0x02, 0x0E, 0x00, 0x00});
  emit_record(frt, 2079, std::vector<std::uint8_t>{0x01, 0x00, 0x00, 0x00});  // fWholeDay
  emit_record(frt, 36, ByteSpan{});
  const std::vector<std::uint8_t> source =
      WithFilters({SxFilter(4U, 1U, 0xFFFFFFFFU, "North", ValuesCriterion("North"), frt)});
  const Workbook wb = Load(source);
  EXPECT_TRUE(wb.sheet(0).pivot_tables().empty());
  auto xlsb = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  EXPECT_EQ(Entry(xlsb.value(), kPivotPart), Entry(source, kPivotPart));
}

TEST(XlsbPivotFidelity, FilterCountMismatchLeavesThePivotAlone) {
  std::vector<std::uint8_t> filter = SxFilter(4U, 1U, 0xFFFFFFFFU, "North", ValuesCriterion("North"));
  // Two filters declared, one present.
  const std::vector<std::uint8_t> package = Fixture();
  const std::vector<std::uint8_t> part = Entry(package, kPivotPart);
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(part)) {
    append_record(patched, part, rec);
    if (rec.type == kBrtTableStyleClient) {
      std::vector<std::uint8_t> payload;
      emit_u32(payload, 2U);
      emit_record(patched, 599, payload);
      patched.insert(patched.end(), filter.begin(), filter.end());
      emit_record(patched, 600, ByteSpan{});
    }
  }
  const Workbook wb = Load(RepackWithEntries(package, {{kPivotPart, patched}}));
  EXPECT_TRUE(wb.sheet(0).pivot_tables().empty());
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
