// Pivot style and advanced filters in XLSB: BrtTableStyleClient and the
// SXFILTERS block of a pivot table part decode into the same
// `PivotStyleInfo` and filter lists an `.xlsx` pivot's
// `<pivotTableStyleInfo>` and `<filters>` produce. The XLSB writer has no
// pivot output, so the part itself stays passthrough and is checked fresh
// against the decoded model on save.

#include <algorithm>
#include <cstdint>
#include <filesystem>
#include <map>
#include <string>
#include <utility>
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
  ASSERT_EQ(table.active_filters().size(), 1U);
  EXPECT_EQ(table.active_filters()[0].type, pivot::FilterType::LabelEqual);
  EXPECT_TRUE(table.retained_filters().empty());

  // The same filter as Windows Excel 365 wrote it to `.xlsx`.
  const std::vector<std::uint8_t> xlsx =
      ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/pivot_caption_filter.xlsx");
  auto twin = read_pivot_table_definition(Entry(xlsx, "xl/pivotTables/pivotTable1.xml"));
  ASSERT_TRUE(static_cast<bool>(twin)) << twin.error().message;
  ASSERT_EQ(twin.value().active_filters().size(), 1U);
  ASSERT_TRUE(table.active_filters()[0].source.has_value());
  ASSERT_TRUE(twin.value().active_filters()[0].source.has_value());
  EXPECT_EQ(table.active_filters()[0].source->xml, twin.value().active_filters()[0].source->xml);

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
  EXPECT_EQ(back.sheet(0).pivot_tables()[0]->active_filters().size(), 1U);
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

/// (idSx, cache definition part) per BrtBeginPivotCacheID of `package`'s
/// workbook part, resolving each relationship id through workbook.bin.rels.
std::vector<std::pair<std::uint32_t, std::string>> PivotCacheIds(const std::vector<std::uint8_t>& package) {
  const std::vector<std::uint8_t> rels_bytes = Entry(package, "xl/_rels/workbook.bin.rels");
  const std::string rels(rels_bytes.begin(), rels_bytes.end());
  const auto attr = [](const std::string& element, const std::string& name) {
    const std::string key = " " + name + "=\"";
    const std::size_t at = element.find(key);
    if (at == std::string::npos) {
      return std::string();
    }
    const std::size_t from = at + key.size();
    return element.substr(from, element.find('"', from) - from);
  };
  const auto target_of = [&](const std::string& id) {
    for (std::size_t at = rels.find("<Relationship "); at != std::string::npos;
         at = rels.find("<Relationship ", at + 1U)) {
      const std::string element = rels.substr(at, rels.find('>', at) - at);
      if (attr(element, "Id") == id) {
        return attr(element, "Target");
      }
    }
    return std::string();
  };
  std::vector<std::pair<std::uint32_t, std::string>> out;
  const std::vector<std::uint8_t> book = Entry(package, "xl/workbook.bin");
  for (const FramedRecord& rec : split_records(book)) {
    if (rec.type != 386U) {
      continue;
    }
    ByteSpan cursor{rec.payload.data + 4, rec.payload.size - 4U};
    auto id = read_xlwidestring(cursor);
    EXPECT_TRUE(static_cast<bool>(id));
    out.emplace_back(load_u32(rec.payload.data), id ? target_of(id.value()) : std::string());
  }
  return out;
}

TEST(XlsbPivotFidelity, SaveKeepsThePivotCacheListOfEveryTwin) {
  std::vector<std::string> paths = {std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_pivot_base.xlsb"};
  for (const auto& entry :
       std::filesystem::directory_iterator(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_pivot_twins")) {
    if (entry.path().extension() == ".xlsb") {
      paths.push_back(entry.path().string());
    }
  }
  ASSERT_GT(paths.size(), 80U);
  for (const std::string& path : paths) {
    SCOPED_TRACE(path);
    const std::vector<std::uint8_t> source = ReadFixtureFile(path);
    const std::vector<std::pair<std::uint32_t, std::string>> expected = PivotCacheIds(source);
    ASSERT_FALSE(expected.empty());
    auto xlsb = write_xlsb(Load(source));
    ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
    EXPECT_EQ(PivotCacheIds(xlsb.value()), expected);
    // The list sits ahead of BrtFileRecover, the future-record blocks and BrtEndBook, as Excel places it.
    const std::vector<FramedRecord> recs = split_records(Entry(xlsb.value(), "xl/workbook.bin"));
    const auto list_end = std::find_if(recs.begin(), recs.end(), [](const FramedRecord& r) { return r.type == 385U; });
    const auto trailer = std::find_if(recs.begin(), recs.end(), [](const FramedRecord& r) {
      return r.type == 155U || r.type == 35U || r.type == 132U;
    });
    EXPECT_LT(list_end, trailer);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
