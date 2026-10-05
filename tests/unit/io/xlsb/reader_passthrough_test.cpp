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
TEST(XlsbReader, DefaultTypedPartsAreCapturedAsPassthrough) {
  std::string content_types(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Default Extension=\"png\" ContentType=\"image/png\"/>"
      "<Override PartName=\"/xl/workbook.bin\" "
      "ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.bin\" "
      "ContentType=\"application/vnd.ms-excel.binIndexWs\"/>"
      "<Override PartName=\"/xl/worksheets/sheet2.bin\" "
      "ContentType=\"application/vnd.ms-excel.binIndexWs\"/>"
      "</Types>");

  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(content_types)});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(2.0)});
  // Typed by the `png` Default rather than an Override: the passthrough
  // sweep never sees it.
  parts.push_back({"xl/media/image1.png", StringToBytes("\x89PNG\r\n\x1a\n")});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  EXPECT_EQ(result.value().dropped_part_count, 0U);
  bool saw_png = false;
  for (const PassthroughPart& part : result.value().workbook.passthrough_parts()) {
    if (part.path == "xl/media/image1.png") {
      saw_png = true;
      EXPECT_TRUE(part.content_type.empty());
      EXPECT_EQ(part.bytes, StringToBytes("\x89PNG\r\n\x1a\n"));
    }
  }
  EXPECT_TRUE(saw_png);
  ASSERT_EQ(result.value().workbook.default_content_types().size(), 3U);
}
TEST(XlsbReader, FullyModelledPackageReportsNoDroppedParts) {
  // Every entry is either consumed by the reader or captured as an
  // Override passthrough, so a lossless package must not raise the flag.
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
  EXPECT_EQ(result.value().dropped_part_count, 0U);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
