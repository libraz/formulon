// The worksheet default-format record (BrtWsFmtInfo) held against Excel 16.115's own xlsb and xlsx saves of the same
// sources (tests/fixtures/excel/xlsb_parity/format), the xlsx reader's normalisation of what Excel normalises on
// open, the encoder's stand-ins for what Excel does not hold, and the reader's skip of what it does not open.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "support/xlsb_package_validator.h"
#include "utils/structured_log.h"
#include "workbook.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using test::Bytes;

constexpr std::uint16_t kBrtWsFmtInfo = 485;

// Sources Excel opened and saved in both formats, its xlsb re-saved as well.
const char* const kOpened[] = {"base",         "base_0",       "base_fraction", "base_255",     "width_0",
                               "width_255",    "width_quant",  "height_0",      "height_409_5", "height_410",
                               "height_large", "height_quant", "outline_row_7"};

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/format/" + name;
}

/// The first worksheet's BrtWsFmtInfo payload; nullopt when it has none.
std::optional<Bytes> FmtInfo(const Bytes& package) {
  const std::optional<Bytes> part = test::ReadPart(package, "xl/worksheets/sheet1.bin");
  EXPECT_TRUE(part.has_value());
  if (!part) {
    return std::nullopt;
  }
  for (const test::IndexedRecord& record : test::SplitPart(*part)) {
    if (record.type == kBrtWsFmtInfo) {
      return record.payload;
    }
  }
  return std::nullopt;
}

Workbook ReadXlsx(const std::string& name) {
  auto read = read_ooxml(reader_test_support::SpanOf(test::ReadBytes(Fixture(name))));
  EXPECT_TRUE(static_cast<bool>(read)) << name << ": " << (read ? "" : read.error().message);
  return std::move(read.value().workbook);
}

XlsbReadResult ReadXlsb(const Bytes& bytes) {
  auto read = read_xlsb(reader_test_support::SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return std::move(read.value());
}

Bytes WriteXlsb(const Workbook& workbook) {
  auto written = write_xlsb_with_result(workbook);
  EXPECT_TRUE(static_cast<bool>(written)) << (written ? "" : written.error().message);
  return written ? written.value().bytes : Bytes{};
}

void ExpectSameDefaults(const SheetFormatDefaults& got, const SheetFormatDefaults& want) {
  EXPECT_EQ(got.base_col_width, want.base_col_width);
  EXPECT_EQ(got.has_default_col_width, want.has_default_col_width);
  EXPECT_EQ(got.default_col_width, want.default_col_width);
  EXPECT_EQ(got.has_default_row_height, want.has_default_row_height);
  EXPECT_EQ(got.default_row_height, want.default_row_height);
  EXPECT_EQ(got.custom_height, want.custom_height);
  EXPECT_EQ(got.zero_height, want.zero_height);
  EXPECT_EQ(got.thick_top, want.thick_top);
  EXPECT_EQ(got.thick_bottom, want.thick_bottom);
  EXPECT_EQ(got.outline_level_row, want.outline_level_row);
  EXPECT_EQ(got.outline_level_col, want.outline_level_col);
}

/// Whether Excel's xlsx open of the case holds 409.6 pt where its xlsb open holds 409.5 pt.
bool ClampsDifferently(const std::string& name) {
  return name == "height_410" || name == "height_large";
}

TEST(XlsbParityFormatDefaults, ConvertedMatchesExcel) {
  for (const char* name : kOpened) {
    SCOPED_TRACE(name);
    const std::optional<Bytes> ours = FmtInfo(WriteXlsb(ReadXlsx(std::string(name) + ".xlsx")));
    const std::optional<Bytes> excel = FmtInfo(test::ReadBytes(Fixture(std::string(name) + ".xlsb")));
    ASSERT_TRUE(ours.has_value());
    ASSERT_TRUE(excel.has_value());
    EXPECT_EQ(*ours, *excel);
  }
}

TEST(XlsbParityFormatDefaults, ReadMatchesXlsxTwin) {
  for (const char* name : kOpened) {
    SCOPED_TRACE(name);
    const XlsbReadResult xlsb = ReadXlsb(test::ReadBytes(Fixture(std::string(name) + ".xlsb")));
    EXPECT_EQ(xlsb.skipped_feature_count, 0U);
    SheetFormatDefaults want = ReadXlsx(std::string(name) + ".xlsx").sheet(0).format_defaults();
    if (ClampsDifferently(name)) {
      EXPECT_EQ(want.default_row_height, 409.6);
      want.default_row_height = 409.5;
    }
    ExpectSameDefaults(xlsb.workbook.sheet(0).format_defaults(), want);
  }
}

TEST(XlsbParityFormatDefaults, ResaveMatchesExcel) {
  for (const char* name : kOpened) {
    SCOPED_TRACE(name);
    const XlsbReadResult xlsb = ReadXlsb(test::ReadBytes(Fixture(std::string(name) + ".xlsb")));
    const std::optional<Bytes> ours = FmtInfo(WriteXlsb(xlsb.workbook));
    const std::optional<Bytes> excel = FmtInfo(test::ReadBytes(Fixture(std::string(name) + ".2.xlsb")));
    ASSERT_TRUE(ours.has_value());
    ASSERT_TRUE(excel.has_value());
    EXPECT_EQ(*ours, *excel);
  }
}

TEST(XlsbParityFormatDefaults, XlsxReadNormalisesAsExcel) {
  // Mac Excel rewrites the base width and an uncustomised height from the Normal font, so only the custom heights
  // and the base width's whole part are compared with its twin.
  for (const char* name : {"height_0", "height_409_5", "height_410", "height_large", "height_quant"}) {
    SCOPED_TRACE(name);
    const SheetFormatDefaults got = ReadXlsx(std::string(name) + "_src.xlsx").sheet(0).format_defaults();
    const SheetFormatDefaults want = ReadXlsx(std::string(name) + ".xlsx").sheet(0).format_defaults();
    EXPECT_EQ(got.default_row_height, want.default_row_height);
    EXPECT_EQ(got.custom_height, want.custom_height);
    EXPECT_EQ(got.zero_height, want.zero_height);
  }
  EXPECT_EQ(ReadXlsx("base_fraction_src.xlsx").sheet(0).format_defaults().base_col_width, 8.0);
}

TEST(XlsbParityFormatDefaults, ExcelRefusedSourcesReadAsWritten) {
  EXPECT_EQ(ReadXlsx("base_256_src.xlsx").sheet(0).format_defaults().base_col_width, 256.0);
  EXPECT_EQ(ReadXlsx("width_256_src.xlsx").sheet(0).format_defaults().default_col_width, 256.0);
  EXPECT_EQ(ReadXlsx("width_record_max_src.xlsx").sheet(0).format_defaults().default_col_width, 65535.0 / 256.0);
  const Workbook row8 = ReadXlsx("outline_row_8_src.xlsx");
  ASSERT_FALSE(row8.sheet(0).layout().row_overrides.empty());
  EXPECT_EQ(row8.sheet(0).layout().row_overrides.front().outline_level, 8U);
}

TEST(XlsbParityFormatDefaults, XlsxSaveKeepsOutlineSummaryAtSeven) {
  for (const char* name : {"outline_row_8_src.xlsx", "outline_col_8_src.xlsx"}) {
    auto written = write_ooxml(ReadXlsx(name));
    ASSERT_TRUE(static_cast<bool>(written)) << name;
    const std::optional<Bytes> part = test::ReadPart(written.value(), "xl/worksheets/sheet1.xml");
    ASSERT_TRUE(part.has_value()) << name;
    const std::string xml(part->begin(), part->end());
    EXPECT_EQ(xml.find("outlineLevelRow=\"8\""), std::string::npos) << name;
    EXPECT_EQ(xml.find("outlineLevelCol=\"8\""), std::string::npos) << name;
    EXPECT_NE(xml.find("outlineLevel=\"8\""), std::string::npos) << name;
    EXPECT_TRUE(xml.find("outlineLevelRow=\"7\"") != std::string::npos ||
                xml.find("outlineLevelCol=\"7\"") != std::string::npos)
        << name;
  }
}

/// The per-kind counts of the deferred-feature log records of the saves made while it lives.
class DeferredKinds {
 public:
  DeferredKinds() {
    set_structured_log_sink(&Collect, &records_);
    set_structured_log_min_level(StructuredLogLevel::kWarn);
  }
  ~DeferredKinds() {
    set_structured_log_min_level(StructuredLogLevel::kOff);
    set_structured_log_sink(nullptr);
  }
  DeferredKinds(const DeferredKinds&) = delete;
  DeferredKinds& operator=(const DeferredKinds&) = delete;

  /// The items the saves reported deferred under `kind`.
  std::size_t count(std::string_view kind) const {
    const std::string key = "\"kind\":\"" + std::string(kind) + "\"";
    std::size_t total = 0U;
    for (const std::string& record : records_) {
      const std::size_t count_at = record.find("\"count\":");
      if (record.find(key) != std::string::npos && count_at != std::string::npos) {
        total += std::strtoul(record.c_str() + count_at + 8U, nullptr, 10);
      }
    }
    return total;
  }

 private:
  static void Collect(std::string_view record, void* user) {
    if (record.find("xlsb.writer.deferred") != std::string_view::npos) {
      static_cast<std::vector<std::string>*>(user)->emplace_back(record);
    }
  }
  std::vector<std::string> records_;
};

/// BrtWsFmtInfo's fields, decoded from its fixed 12-byte payload.
struct FmtFields {
  std::uint32_t dx_g_col = 0U;
  std::uint16_t base_col_width = 0U;
  std::uint16_t row_height_twips = 0U;
  std::uint32_t flags = 0U;
};

FmtFields DecodeFmt(const Bytes& payload) {
  EXPECT_EQ(payload.size(), 12U);
  FmtFields out;
  if (payload.size() != 12U) {
    return out;
  }
  const auto le = [&](std::size_t at, std::size_t width) {
    std::uint32_t v = 0U;
    for (std::size_t i = 0; i < width; ++i) {
      v |= static_cast<std::uint32_t>(payload[at + i]) << (8U * i);
    }
    return v;
  };
  out.dx_g_col = le(0U, 4U);
  out.base_col_width = static_cast<std::uint16_t>(le(4U, 2U));
  out.row_height_twips = static_cast<std::uint16_t>(le(6U, 2U));
  out.flags = le(8U, 4U);
  return out;
}

constexpr std::uint32_t kAbsentWidth = 0xFFFFFFFFU;

TEST(XlsbParityFormatDefaults, EncoderWritesStandInsAndReportsThem) {
  struct Case {
    void (*edit)(SheetFormatDefaults&);
    std::uint32_t dx_g_col;
    std::uint16_t base_col_width;
    std::uint16_t row_height_twips;
  };
  const Case cases[] = {
      {[](SheetFormatDefaults& d) { d.base_col_width = 8.5; }, kAbsentWidth, 8U, 300U},
      {[](SheetFormatDefaults& d) { d.base_col_width = 256.0; }, kAbsentWidth, 8U, 300U},
      {[](SheetFormatDefaults& d) {
         d.has_default_col_width = true;
         d.default_col_width = 255.5;
       },
       kAbsentWidth, 8U, 300U},
      {[](SheetFormatDefaults& d) {
         d.has_default_row_height = true;
         d.default_row_height = 409.65;
       },
       kAbsentWidth, 8U, 300U},
  };
  for (const Case& c : cases) {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("Sheet1");
    SheetFormatDefaults& defaults = wb.sheet(0).mutable_format_defaults();
    defaults.base_col_width = 8.0;
    defaults.has_default_row_height = true;
    defaults.default_row_height = 15.0;
    c.edit(defaults);
    const DeferredKinds deferred;
    auto written = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    EXPECT_EQ(deferred.count("format_defaults"), 1U);
    EXPECT_GE(written.value().diagnostics.deferred_feature_count, 1U);
    const std::optional<Bytes> payload = FmtInfo(written.value().bytes);
    ASSERT_TRUE(payload.has_value());
    const FmtFields fields = DecodeFmt(*payload);
    EXPECT_EQ(fields.dx_g_col, c.dx_g_col);
    EXPECT_EQ(fields.base_col_width, c.base_col_width);
    EXPECT_EQ(fields.row_height_twips, c.row_height_twips);
    EXPECT_EQ(fields.flags, 0U);
  }

  // Outline level 8 is clamped to 7 in the summary and not reported.
  Workbook rows = Workbook::create_empty();
  rows.add_sheet("Sheet1");
  RowLayout row;
  row.row = 2U;
  row.outline_level = 8U;
  rows.sheet(0).mutable_layout().row_overrides.push_back(row);
  Workbook columns = Workbook::create_empty();
  columns.add_sheet("Sheet1");
  ColumnLayout column;
  column.first = 1U;
  column.last = 1U;
  column.outline_level = 8U;
  columns.sheet(0).mutable_layout().columns.push_back(column);
  const std::pair<const Workbook*, std::uint32_t> outlines[] = {{&rows, 7U << 16U}, {&columns, 7U << 24U}};
  for (const auto& [wb, flags] : outlines) {
    const DeferredKinds deferred;
    auto written = write_xlsb_with_result(*wb);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
    EXPECT_EQ(deferred.count("format_defaults"), 0U);
    EXPECT_EQ(deferred.count("outline_level"), 0U);
    const std::optional<Bytes> payload = FmtInfo(written.value().bytes);
    ASSERT_TRUE(payload.has_value());
    EXPECT_EQ(DecodeFmt(*payload).flags, flags);
  }
}

/// A two-sheet xlsb whose first sheet is `first_sheet`.
Bytes BookWithFirstSheet(const std::vector<std::uint8_t>& first_sheet) {
  using namespace reader_test_support;
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", first_sheet});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});
  return BuildZip(parts);
}

/// A two-sheet xlsb whose first sheet carries `payload` as its BrtWsFmtInfo.
Bytes BookWithFormatPayload(const std::vector<std::uint8_t>& payload) {
  return BookWithFirstSheet(reader_test_support::SheetBinWorksheetFormat(payload));
}

TEST(XlsbParityFormatDefaults, ReadSkipsUndecodable) {
  using reader_test_support::WorksheetFormatPayload;
  const std::vector<std::uint8_t> valid = WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 0U);
  std::vector<std::vector<std::uint8_t>> skipped = {
      WorksheetFormatPayload(255U * 256U + 1U, 8U, 300U, 0U),  // width past 255
      WorksheetFormatPayload(65536U, 8U, 300U, 0U),
      WorksheetFormatPayload(0xFFFFFFFFU, 256U, 300U, 0U),
      WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 8U << 16U),
      WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 8U << 24U),
      WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 0x10U),  // reserved flag bit
  };
  for (const std::size_t length : {3U, 5U, 7U, 11U}) {
    skipped.emplace_back(valid.data(), valid.data() + length);
  }
  // A skipped record leaves the sheet as one without the record.
  const XlsbReadResult without = ReadXlsb(BookWithFirstSheet(reader_test_support::SheetBinReal(1.0)));
  const SheetFormatDefaults model_defaults = without.workbook.sheet(0).format_defaults();
  for (const std::vector<std::uint8_t>& payload : skipped) {
    SCOPED_TRACE(payload.size());
    const XlsbReadResult read = ReadXlsb(BookWithFormatPayload(payload));
    EXPECT_EQ(read.skipped_feature_count, 1U);
    ExpectSameDefaults(read.workbook.sheet(0).format_defaults(), model_defaults);
  }

  std::vector<std::uint8_t> trailing = WorksheetFormatPayload(0U, 10U, 280U, 0U);
  trailing.push_back(0U);
  const XlsbReadResult trailing_read = ReadXlsb(BookWithFormatPayload(trailing));
  EXPECT_EQ(trailing_read.skipped_feature_count, 0U);
  EXPECT_EQ(trailing_read.workbook.sheet(0).format_defaults().base_col_width, 10.0);
}

TEST(XlsbParityFormatDefaults, ReadClampsDefaultRowHeightAsExcel) {
  using reader_test_support::WorksheetFormatPayload;
  for (const std::uint16_t twips : {std::uint16_t{8191U}, std::uint16_t{8192U}, std::uint16_t{65535U}}) {
    SCOPED_TRACE(twips);
    const XlsbReadResult read = ReadXlsb(BookWithFormatPayload(WorksheetFormatPayload(0xFFFFFFFFU, 8U, twips, 1U)));
    EXPECT_EQ(read.skipped_feature_count, 0U);
    EXPECT_EQ(read.workbook.sheet(0).format_defaults().default_row_height, 409.5);
  }
}

constexpr std::uint16_t kBrtBeginWsView = 137;

/// The BrtBeginWsView records of worksheet part `sheet` (1-based).
std::size_t ViewCount(const Bytes& package, std::size_t sheet) {
  const std::optional<Bytes> part = test::ReadPart(package, "xl/worksheets/sheet" + std::to_string(sheet) + ".bin");
  EXPECT_TRUE(part.has_value());
  std::size_t views = 0U;
  for (const test::IndexedRecord& record : test::SplitPart(part.value_or(Bytes{}))) {
    views += record.type == kBrtBeginWsView ? 1U : 0U;
  }
  return views;
}

// Sources whose xlsb save once succeeded: each saves again, writing a stand-in for a default Excel does not hold
// (reported as format_defaults), outline level 8 as 7 (not reported), and leaving out an extra view of a window
// the workbook lacks (reported as sheet_view); the package validates and reads back.
/// The first worksheet's records of `type`, in stream order.
std::vector<Bytes> RecordsOf(const Bytes& package, std::uint16_t type) {
  std::vector<Bytes> out;
  const std::optional<Bytes> part = test::ReadPart(package, "xl/worksheets/sheet1.bin");
  EXPECT_TRUE(part.has_value());
  if (part) {
    for (const test::IndexedRecord& record : test::SplitPart(*part)) {
      if (record.type == type) {
        out.push_back(record.payload);
      }
    }
  }
  return out;
}

TEST(XlsbParityFormatDefaults, RowAndColumnOutlineLevelsAboveSevenWriteSeven) {
  // BrtRowHdr: iOutLevel is the low three bits of the byte after flags1 (offset 11).
  constexpr std::size_t kRowOutlineOffset = 11U;
  // BrtColInfo: the level sits in bits 8-10 of the flags at offset 16.
  constexpr std::size_t kColFlagsOffset = 16U;
  const Bytes rows = WriteXlsb(ReadXlsx("outline_row_8_src.xlsx"));
  bool row_seen = false;
  for (const Bytes& header : RecordsOf(rows, static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr))) {
    ASSERT_GT(header.size(), kRowOutlineOffset);
    if (header[0] == 2U) {  // row 3
      EXPECT_EQ(header[kRowOutlineOffset] & 0x07U, 7U);
      row_seen = true;
    }
  }
  EXPECT_TRUE(row_seen);
  const XlsbReadResult row_back = ReadXlsb(rows);
  const auto& overrides = row_back.workbook.sheet(0).layout().row_overrides;
  const auto row3 = std::find_if(overrides.begin(), overrides.end(), [](const RowLayout& r) { return r.row == 2U; });
  ASSERT_NE(row3, overrides.end());
  EXPECT_EQ(row3->outline_level, 7U);

  const Bytes cols = WriteXlsb(ReadXlsx("outline_col_8_src.xlsx"));
  const std::vector<Bytes> infos = RecordsOf(cols, static_cast<std::uint16_t>(XlsbRecordType::BrtColInfo));
  ASSERT_EQ(infos.size(), 1U);
  ASSERT_GT(infos[0].size(), kColFlagsOffset + 1U);
  EXPECT_EQ(infos[0][kColFlagsOffset + 1U] & 0x07U, 7U);
  const XlsbReadResult col_back = ReadXlsb(cols);
  ASSERT_EQ(col_back.workbook.sheet(0).layout().columns.size(), 1U);
  EXPECT_EQ(col_back.workbook.sheet(0).layout().columns[0].outline_level, 7U);
}

TEST(XlsbSaveRestoredSources, TenSourcesSaveOpenValidateAndReload) {
  struct Case {
    const char* path;
    std::size_t format_defaults;
    std::size_t sheet_view;
    std::uint32_t dx_g_col;
    std::uint16_t base_col_width;
    std::uint16_t row_height_twips;
    std::uint32_t flags;
  };
  const Case cases[] = {
      {"format/base_256_src.xlsx", 1U, 0U, kAbsentWidth, 8U, 300U, 0U},
      {"format/base_large_src.xlsx", 1U, 0U, kAbsentWidth, 8U, 300U, 0U},
      {"format/width_256_src.xlsx", 1U, 0U, kAbsentWidth, 8U, 300U, 0U},
      {"format/width_large_src.xlsx", 1U, 0U, kAbsentWidth, 8U, 300U, 0U},
      {"format/width_record_max_src.xlsx", 1U, 0U, kAbsentWidth, 8U, 300U, 0U},
      {"format/width_record_max.xlsx", 1U, 0U, kAbsentWidth, 10U, 280U, 0U},
      {"format/outline_col_8_src.xlsx", 0U, 0U, kAbsentWidth, 8U, 300U, 7U << 24U},
      {"format/outline_row_8_src.xlsx", 0U, 0U, kAbsentWidth, 8U, 300U, 7U << 16U},
      {"views/bad_window_id_src.xlsx", 0U, 1U, kAbsentWidth, 8U, 300U, 0U},
      {"views/window_id_past_src.xlsx", 0U, 1U, kAbsentWidth, 8U, 300U, 0U},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.path);
    auto read = read_ooxml(reader_test_support::SpanOf(
        test::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/" + c.path)));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    const DeferredKinds deferred;
    auto written = write_xlsb_with_result(read.value().workbook);
    ASSERT_TRUE(static_cast<bool>(written)) << written.error().message << " | " << written.error().context;
    const Bytes& bytes = written.value().bytes;
    EXPECT_EQ(::formulon::test::ValidateXlsbPackage(reader_test_support::SpanOf(bytes)), std::vector<std::string>{});
    EXPECT_TRUE(static_cast<bool>(read_xlsb(reader_test_support::SpanOf(bytes))));
    EXPECT_EQ(deferred.count("format_defaults"), c.format_defaults);
    EXPECT_EQ(deferred.count("sheet_view"), c.sheet_view);
    EXPECT_EQ(deferred.count("outline_level"), 0U);
    EXPECT_GE(written.value().diagnostics.deferred_feature_count, c.format_defaults + c.sheet_view);
    const std::optional<Bytes> payload = FmtInfo(bytes);
    ASSERT_TRUE(payload.has_value());
    const FmtFields fields = DecodeFmt(*payload);
    EXPECT_EQ(fields.dx_g_col, c.dx_g_col);
    EXPECT_EQ(fields.base_col_width, c.base_col_width);
    EXPECT_EQ(fields.row_height_twips, c.row_height_twips);
    EXPECT_EQ(fields.flags, c.flags);
    // The extra view of window 5 is left out; only the main view is written.
    EXPECT_EQ(ViewCount(bytes, 1U), 1U);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
