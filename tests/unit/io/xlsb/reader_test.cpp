// In-memory XLSB reader tests grouped by cells, formulas, validation, and passthrough.

#include "io/xlsb/reader.h"

#include <array>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "miniz.h"
#include "reader_test_helpers.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {
using namespace reader_test_support;
TEST(XlsbReader, ReadsTwoSheetsWithRealAndIsstCells) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(123.5)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinIsst(0)});
  parts.push_back({"xl/sharedStrings.bin", SharedStringsBin("hello world")});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Workbook& wb = result.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
  EXPECT_EQ(wb.sheet(1).name(), "Beta");

  // Sheet 1: numeric cell.
  const Cell* c0 = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(c0, nullptr);
  ASSERT_TRUE(c0->cached_value.is_number());
  EXPECT_EQ(c0->cached_value.as_number(), 123.5);

  // Sheet 2: text cell resolved against the SST.
  const Cell* c1 = wb.sheet(1).cell_at(0, 0);
  ASSERT_NE(c1, nullptr);
  ASSERT_TRUE(c1->cached_value.is_text());
  EXPECT_EQ(c1->cached_value.as_text(), "hello world");

  // Audit counter: two literal cells decoded.
  EXPECT_EQ(result.value().cells_read, 2U);
}
TEST(XlsbReader, DecodesWorksheetFormatDefaults) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinWorksheetFormat(WorksheetFormatPayload(3200U, 10U, 375U, 1U))});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  auto result = read_xlsb(SpanOf(BuildZip(parts)));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;
  const SheetFormatDefaults& defaults = result.value().workbook.sheet(0).format_defaults();
  EXPECT_DOUBLE_EQ(defaults.base_col_width, 10.0);
  EXPECT_TRUE(defaults.has_default_col_width);
  EXPECT_DOUBLE_EQ(defaults.default_col_width, 12.5);
  EXPECT_TRUE(defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(defaults.default_row_height, 18.75);
}
TEST(XlsbReader, DecodesAbsentAndExcelCompatibleWorksheetFormatDefaults) {
  const auto read_defaults = [](const std::vector<std::uint8_t>& payload) {
    std::vector<PartFile> parts;
    parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
    parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
    parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
    parts.push_back({"xl/workbook.bin", WorkbookBin()});
    parts.push_back({"xl/worksheets/sheet1.bin", SheetBinWorksheetFormat(payload)});
    parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});
    return read_xlsb(SpanOf(BuildZip(parts)));
  };

  auto absent = read_defaults(WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 0U));
  ASSERT_TRUE(static_cast<bool>(absent)) << absent.error().message << " | " << absent.error().context;
  const SheetFormatDefaults& absent_defaults = absent.value().workbook.sheet(0).format_defaults();
  EXPECT_DOUBLE_EQ(absent_defaults.base_col_width, 8.0);
  EXPECT_FALSE(absent_defaults.has_default_col_width);
  EXPECT_DOUBLE_EQ(absent_defaults.default_col_width, 0.0);
  EXPECT_FALSE(absent_defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(absent_defaults.default_row_height, 0.0);

  auto excel = read_defaults(WorksheetFormatPayload(0xFFFFFFFFU, 10U, 400U, 0U));
  ASSERT_TRUE(static_cast<bool>(excel)) << excel.error().message << " | " << excel.error().context;
  const SheetFormatDefaults& excel_defaults = excel.value().workbook.sheet(0).format_defaults();
  EXPECT_DOUBLE_EQ(excel_defaults.base_col_width, 10.0);
  EXPECT_FALSE(excel_defaults.has_default_col_width);
  EXPECT_TRUE(excel_defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(excel_defaults.default_row_height, 20.0);

  auto zero = read_defaults(WorksheetFormatPayload(0U, 8U, 0U, 2U));
  ASSERT_TRUE(static_cast<bool>(zero)) << zero.error().message << " | " << zero.error().context;
  const SheetFormatDefaults& zero_defaults = zero.value().workbook.sheet(0).format_defaults();
  EXPECT_TRUE(zero_defaults.has_default_col_width);
  EXPECT_DOUBLE_EQ(zero_defaults.default_col_width, 0.0);
  EXPECT_TRUE(zero_defaults.has_default_row_height);
  EXPECT_DOUBLE_EQ(zero_defaults.default_row_height, 0.0);
}
TEST(XlsbReader, RejectsMalformedWorksheetFormatInfoPayloads) {
  const auto read_payload = [](const std::vector<std::uint8_t>& payload) {
    std::vector<PartFile> parts;
    parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
    parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
    parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
    parts.push_back({"xl/workbook.bin", WorkbookBin()});
    parts.push_back({"xl/worksheets/sheet1.bin", SheetBinWorksheetFormat(payload)});
    parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});
    return read_xlsb(SpanOf(BuildZip(parts)));
  };

  const std::vector<std::uint8_t> valid = WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 0U);
  for (const std::size_t length : {3U, 5U, 7U, 11U}) {
    std::vector<std::uint8_t> truncated(valid.data(), valid.data() + length);
    auto result = read_payload(truncated);
    ASSERT_FALSE(static_cast<bool>(result)) << "payload length " << length << " unexpectedly succeeded";
    EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
  }

  auto invalid_dx = read_payload(WorksheetFormatPayload(65536U, 8U, 300U, 0U));
  ASSERT_FALSE(static_cast<bool>(invalid_dx));
  EXPECT_EQ(invalid_dx.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);

  auto invalid_cch = read_payload(WorksheetFormatPayload(0xFFFFFFFFU, 256U, 300U, 0U));
  ASSERT_FALSE(static_cast<bool>(invalid_cch));
  EXPECT_EQ(invalid_cch.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);

  auto invalid_outline = read_payload(WorksheetFormatPayload(0xFFFFFFFFU, 8U, 300U, 8U << 16U));
  ASSERT_FALSE(static_cast<bool>(invalid_outline));
  EXPECT_EQ(invalid_outline.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);

  std::vector<std::uint8_t> trailing = valid;
  trailing.push_back(0U);
  auto trailing_result = read_payload(trailing);
  ASSERT_FALSE(static_cast<bool>(trailing_result));
  EXPECT_EQ(trailing_result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbReader, RowStylePresenceUsesFGhostDirtyFlag) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinRowStyles()});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;
  const auto& rows = result.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 2U);
  const auto find_row = [&rows](std::uint32_t row) -> const RowLayout* {
    for (const RowLayout& candidate : rows) {
      if (candidate.row == row) {
        return &candidate;
      }
    }
    return nullptr;
  };
  const RowLayout* style_zero = find_row(0U);
  const RowLayout* style_one = find_row(1U);
  ASSERT_NE(style_zero, nullptr);
  ASSERT_NE(style_one, nullptr);
  EXPECT_TRUE(style_zero->has_style);
  EXPECT_EQ(style_zero->style_xf, 0U);
  EXPECT_TRUE(style_one->has_style);
  EXPECT_EQ(style_one->style_xf, 1U);
  EXPECT_EQ(find_row(2U), nullptr);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
