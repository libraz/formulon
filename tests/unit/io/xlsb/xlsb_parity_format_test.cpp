// The worksheet default-format record (BrtWsFmtInfo) held against Excel 16.115's own xlsb and xlsx saves of the same
// sources (tests/fixtures/excel/xlsb_parity/format), the xlsx reader's normalisation of what Excel normalises on
// open, the encoder's refusal of what Excel does not hold, and the reader's skip of what it does not open.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "utils/error.h"
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

/// The error of an xlsb save of `workbook`; the save must fail.
Error SaveError(const Workbook& workbook) {
  auto written = write_xlsb_with_result(workbook);
  EXPECT_FALSE(static_cast<bool>(written));
  return written ? Error{} : written.error();
}

TEST(XlsbParityFormatDefaults, EncoderRefusesUnrepresentable) {
  const auto with_defaults = [](void (*edit)(SheetFormatDefaults&)) {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("Sheet1");
    edit(wb.sheet(0).mutable_format_defaults());
    return wb;
  };
  const Workbook refused[] = {
      with_defaults([](SheetFormatDefaults& d) { d.base_col_width = 8.5; }),
      with_defaults([](SheetFormatDefaults& d) { d.base_col_width = 256.0; }),
      with_defaults([](SheetFormatDefaults& d) {
        d.has_default_col_width = true;
        d.default_col_width = 255.5;
      }),
      with_defaults([](SheetFormatDefaults& d) {
        d.has_default_row_height = true;
        d.default_row_height = 409.65;
      }),
  };
  for (const Workbook& wb : refused) {
    const Error error = SaveError(wb);
    EXPECT_EQ(error.code, FormulonErrorCode::kIoFeatureNotWritable);
    EXPECT_EQ(error.context, "context=write_xlsb kind=format_defaults sheet_index=0");
  }

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
  for (const Workbook* wb : {&rows, &columns}) {
    const Error error = SaveError(*wb);
    EXPECT_EQ(error.code, FormulonErrorCode::kIoFeatureNotWritable);
    EXPECT_EQ(error.context, "context=write_xlsb kind=outline_level sheet_index=0");
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

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
