//
// Negative-side test for `ooxml_reader`'s workbook-kind detection: when
// `[Content_Types].xml` declares an Override for `/xl/workbook.xml`
// with a content type the reader does not recognise, the read still
// succeeds, the workbook kind defaults to `kXlsx`, and a structured-log
// warning (`ooxml.reader.unknown_workbook_content_type`) is emitted.
//
// The warning is observed through a log sink rather than by capturing
// stderr. Logging ships off (`StructuredLogLevel::kOff`) so an embedded
// library never writes to the host's stderr unasked, so a test that
// wants records has to opt in the way an embedder does; reading them
// back from the sink also keeps the assertion independent of both that
// default and gtest's internal capture hook.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/workbook_kind.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

/// Collects structured-log records for the duration of one test and
/// restores the shipped configuration (no sink, `kOff`) afterwards, so
/// enabling logging here cannot leak into a sibling test.
class StructuredLogCapture {
 public:
  StructuredLogCapture() {
    set_structured_log_sink(&Append, &records_);
    set_structured_log_min_level(StructuredLogLevel::kDebug);
  }
  ~StructuredLogCapture() {
    set_structured_log_sink(nullptr);
    set_structured_log_min_level(StructuredLogLevel::kOff);
  }

  StructuredLogCapture(const StructuredLogCapture&) = delete;
  StructuredLogCapture& operator=(const StructuredLogCapture&) = delete;

  const std::string& records() const { return records_; }

 private:
  static void Append(std::string_view record, void* user_data) { static_cast<std::string*>(user_data)->append(record); }

  std::string records_;
};

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

// ---------------------------------------------------------------------------
// Synthetic package: workbook part is declared with a bogus content type.
// ---------------------------------------------------------------------------

constexpr std::string_view kPackageRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
    "  <Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
    "Target=\"xl/workbook.xml\"/>\n"
    "</Relationships>\n";

constexpr std::string_view kWorkbookXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
    "  <sheets>\n"
    "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
    "  </sheets>\n"
    "</workbook>\n";

constexpr std::string_view kWorkbookRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
    "  <Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
    "Target=\"worksheets/sheet1.xml\"/>\n"
    "</Relationships>\n";

constexpr std::string_view kSheetXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
    "  <sheetData/>\n"
    "</worksheet>\n";

constexpr std::string_view kBogusContentTypes =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
    "  <Default Extension=\"rels\" "
    "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
    "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
    "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.bogus.foo+xml\"/>\n"
    "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
    "</Types>\n";

TEST(OoxmlUnknownKind, FallsBackToXlsxAndDoesNotFailRead) {
  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", kBogusContentTypes},
      {"_rels/.rels", kPackageRels},
      {"xl/workbook.xml", kWorkbookXml},
      {"xl/_rels/workbook.xml.rels", kWorkbookRels},
      {"xl/worksheets/sheet1.xml", kSheetXml},
  });

  // Structured logging ships off so the library stays silent inside an
  // embedding host; this test is the embedder that opts in.
  StructuredLogCapture log;
  auto result_or = read_ooxml(SpanOf(bytes));

  ASSERT_TRUE(static_cast<bool>(result_or))
      << "read_ooxml unexpectedly failed for unknown workbook content type: " << result_or.error().message;
  EXPECT_EQ(result_or.value().workbook.kind(), WorkbookKind::kXlsx);

  // The structured-log warning is the secondary signal: guessing `kXlsx`
  // silently would leave the caller no way to learn the package declared
  // something else. The full JSON shape is not asserted here because the
  // emitter may gain fields over time, only that the event reached the
  // sink.
  EXPECT_NE(log.records().find("ooxml.reader.unknown_workbook_content_type"), std::string::npos)
      << "expected structured-log warning was not emitted; records were: " << log.records();
}

TEST(OoxmlUnknownKind, MissingWorkbookOverrideStillFails) {
  // Sanity counter-test: a `[Content_Types].xml` with no Override at
  // all targeting `xl/workbook.xml` must continue to fail with
  // `kIoContentTypeInvalid` — the unknown-kind fallback only triggers
  // when an Override exists with an unrecognised content type.
  constexpr std::string_view kNoWorkbookOverride =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "</Types>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", kNoWorkbookOverride},
      {"_rels/.rels", kPackageRels},
      {"xl/workbook.xml", kWorkbookXml},
      {"xl/_rels/workbook.xml.rels", kWorkbookRels},
      {"xl/worksheets/sheet1.xml", kSheetXml},
  });

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

}  // namespace
}  // namespace io
}  // namespace formulon
