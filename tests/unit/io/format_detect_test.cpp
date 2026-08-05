//
// Tests for workbook container-format detection. The detector decides
// xlsx vs xlsb from the package's central directory so the byte-only C
// ABI load boundary can route to the right reader.

#include "io/format_detect.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "miniz.h"

namespace formulon {
namespace io {
namespace {

struct PartFile {
  std::string path;
  std::string body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const PartFile& p : parts) {
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, p.path.c_str(), p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE);
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_TRUE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  mz_zip_writer_end(&writer);
  return out;
}

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

TEST(FormatDetect, BinaryWorkbookPartIsXlsb) {
  const std::vector<std::uint8_t> archive = BuildZip({
      {"[Content_Types].xml", "<Types/>"},
      {"xl/workbook.bin", std::string("\x01\x02\x03", 3)},
  });
  EXPECT_EQ(detect_workbook_format(SpanOf(archive)), WorkbookFormat::Xlsb);
}

TEST(FormatDetect, XmlWorkbookPartIsOoxml) {
  const std::vector<std::uint8_t> archive = BuildZip({
      {"[Content_Types].xml", "<Types/>"},
      {"xl/workbook.xml", "<workbook/>"},
  });
  EXPECT_EQ(detect_workbook_format(SpanOf(archive)), WorkbookFormat::Ooxml);
}

TEST(FormatDetect, ContentTypeFallbackResolvesXlsb) {
  // No standard workbook part name; the xlsb content type in
  // [Content_Types].xml is the only signal.
  const std::vector<std::uint8_t> archive = BuildZip({
      {"[Content_Types].xml",
       "<Types><Override PartName=\"/xl/wb.bin\" "
       "ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/></Types>"},
      {"xl/wb.bin", std::string("\x00", 1)},
  });
  EXPECT_EQ(detect_workbook_format(SpanOf(archive)), WorkbookFormat::Xlsb);
}

TEST(FormatDetect, XlsmContentTypeIsNotMisdetectedAsXlsb) {
  // Regression: `.xlsm`'s macro-enabled main-part content type
  // (`.../main+xml`) is a superstring of `kCtWorkbookXlsbAlt`
  // (`.../main`, no `+xml`). A plain substring search used to match the
  // xlsb fallback here, misclassifying a non-standard-named `.xlsm`
  // workbook part as MS-XLSB.
  const std::vector<std::uint8_t> archive = BuildZip({
      {"[Content_Types].xml",
       "<Types><Override PartName=\"/xl/wb.xml\" "
       "ContentType=\"application/vnd.ms-excel.sheet.macroEnabled.main+xml\"/></Types>"},
      {"xl/wb.xml", "<workbook/>"},
  });
  EXPECT_NE(detect_workbook_format(SpanOf(archive)), WorkbookFormat::Xlsb);
}

TEST(FormatDetect, ContentTypeFallbackResolvesXlsbAltSpelling) {
  // The "alt" xlsb content-type spelling still matches when it is the
  // exact (quote-delimited) attribute value.
  const std::vector<std::uint8_t> archive = BuildZip({
      {"[Content_Types].xml",
       "<Types><Override PartName=\"/xl/wb.bin\" "
       "ContentType=\"application/vnd.ms-excel.sheet.macroEnabled.main\"/></Types>"},
      {"xl/wb.bin", std::string("\x00", 1)},
  });
  EXPECT_EQ(detect_workbook_format(SpanOf(archive)), WorkbookFormat::Xlsb);
}

TEST(FormatDetect, NonZipIsUnknown) {
  const std::vector<std::uint8_t> garbage = {0xDE, 0xAD, 0xBE, 0xEF, 0x00};
  EXPECT_EQ(detect_workbook_format(SpanOf(garbage)), WorkbookFormat::Unknown);
}

}  // namespace
}  // namespace io
}  // namespace formulon
