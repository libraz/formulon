// In-memory XLSB reader tests grouped by cells, formulas, validation, and passthrough.

#include <array>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
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
TEST(XlsbReader, OutOfRangeCellColumnIsRecordCorrupt) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  // Column index == kMaxCols is one past the last valid column (XFD).
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0, /*row=*/0, /*col=*/Sheet::kMaxCols)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbReader, OutOfRangeRowHeaderIsRecordCorrupt) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  // Row index == kMaxRows is one past the last valid row.
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0, /*row=*/Sheet::kMaxRows, /*col=*/0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbReader, TruncatedPhoneticRunArrayIsRecordTruncated) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinIsst(0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/sharedStrings.bin", TruncatedPhoneticSharedStringsBin()});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordTruncated);
}
TEST(XlsbReader, ReversedArrayFormulaRectIsRecordCorrupt) {
  // rw_last > rw_first but col_last < col_first: reversed on the column
  // axis only. The anchor guard is an OR over the two axes, so without
  // rect validation this anchor would be recorded and the spill size
  // math would underflow-wrap.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};  // PtgInt 5
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin",
                   SheetBinArrFmla(/*rw_first=*/0, /*rw_last=*/1, /*col_first=*/2, /*col_last=*/1, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbReader, InBoundsArrayFormulaRectStillDecodes) {
  // A well-ordered in-bounds 2x1 rect anchored at A1 must keep decoding
  // after the bounds validation: the anchor gets the decoded formula
  // text and the read succeeds.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};  // PtgInt 5
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin",
                   SheetBinArrFmla(/*rw_first=*/0, /*rw_last=*/1, /*col_first=*/0, /*col_last=*/0, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  EXPECT_EQ(c->formula_text, "=5");
}
TEST(XlsbReader, RejectsFullGridArrayFormulaBeforeSpillWalk) {
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};  // PtgInt 5
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back(
      {"xl/worksheets/sheet1.bin", SheetBinArrFmla(/*rw_first=*/0, /*rw_last=*/Sheet::kMaxRows - 1U, /*col_first=*/0,
                                                   /*col_last=*/Sheet::kMaxCols - 1U, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}
TEST(XlsbReader, RejectsCumulativeArrayFormulasOverSheetBudget) {
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};  // PtgInt 5
  const std::vector<std::array<std::uint32_t, 4>> rects = {
      {0U, 599999U, 0U, 0U},
      {0U, 599999U, 1U, 1U},
  };
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinArrFmlas(rects, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
  EXPECT_NE(result.error().context.find("format=xlsb"), std::string::npos);
  EXPECT_NE(result.error().context.find("anchor_row=0 anchor_col=1"), std::string::npos);
  EXPECT_NE(result.error().context.find("used=600000 requested=600000 ceiling=1048576"), std::string::npos);
}
TEST(XlsbReader, MissingContentTypesIsContentTypeInvalid) {
  std::vector<PartFile> parts;
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
