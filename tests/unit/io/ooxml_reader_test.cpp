//
// Unit tests for `formulon::io::read_ooxml`. The reader is paired with
// `Workbook::save()` (the empty-workbook writer) so the test corpus is
// generated in-process; no fixture files are needed for this slice.

#include "io/ooxml_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml/package_validator.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed in test setup";
  return save_or.value();
}

TEST(OoxmlReader, RoundTripsSingleSheetName) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;
  OoxmlReadResult& result = result_or.value();
  ASSERT_EQ(result.workbook.sheet_count(), 1U);
  EXPECT_EQ(result.workbook.sheet(0).name(), "Sheet1");
}

TEST(OoxmlReader, RoundTripsThreeSheetsInOrder) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  wb.add_sheet("\xE3\x82\xAC\xE3\x83\xB3\xE3\x83\x9E");  // "ガンマ" in UTF-8

  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;
  OoxmlReadResult& result = result_or.value();
  ASSERT_EQ(result.workbook.sheet_count(), 3U);
  EXPECT_EQ(result.workbook.sheet(0).name(), "Alpha");
  EXPECT_EQ(result.workbook.sheet(1).name(), "Beta");
  EXPECT_EQ(result.workbook.sheet(2).name(), "\xE3\x82\xAC\xE3\x83\xB3\xE3\x83\x9E");
}

// Regression: every text-cell `Value::text` must remain valid for the
// workbook's lifetime, NOT only for the lifetime of the
// `OoxmlReadResult`. Earlier slices stashed the inline-string deque on
// the read result; moving the workbook out of the result and discarding
// the result therefore left every text view dangling. The fix moves
// the deque onto the workbook itself; this test pins that contract.
TEST(OoxmlReader, WorkbookOutlivesReadResult) {
  // Build a workbook with at least one inline-string cell so the
  // text-storage deque is exercised on the read path.
  Workbook src = Workbook::create();
  std::string greeting = "Hello, world!";
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::text(greeting))));
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  // Move the workbook out of the read result and immediately drop the
  // result. The workbook must still report the cell text correctly.
  Workbook wb = Workbook::create_empty();
  {
    auto result_or = read_ooxml(SpanOf(bytes));
    ASSERT_TRUE(static_cast<bool>(result_or));
    wb = std::move(result_or.value().workbook);
  }  // OoxmlReadResult destructed here.

  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "Hello, world!");

  // And `save()` must succeed — the writer walks every cell's
  // `Value::text` view and would crash under the old layout.
  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  EXPECT_FALSE(save_or.value().empty());
}

TEST(OoxmlReader, JapaneseSheetNameSurvivesRoundTrip) {
  Workbook wb = Workbook::create();
  // "日本シート" in UTF-8: E6 97 A5 E6 9C AC E3 82 B7 E3 83 BC E3 83 88
  wb.sheet(0).set_name("\xE6\x97\xA5\xE6\x9C\xAC\xE3\x82\xB7\xE3\x83\xBC\xE3\x83\x88");
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "\xE6\x97\xA5\xE6\x9C\xAC\xE3\x82\xB7\xE3\x83\xBC\xE3\x83\x88");
}

TEST(OoxmlReader, RejectsNonZipBuffer) {
  const std::uint8_t garbage[] = {0u, 0u, 0u, 0u};
  ByteSpan span{garbage, sizeof(garbage)};
  auto result_or = read_ooxml(span);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipCorrupt);
}

/// Builds an in-memory ZIP that omits a specific entry from the writer's
/// output. Used to verify reader error paths.
std::vector<std::uint8_t> RebuildArchiveWithout(const std::vector<std::uint8_t>& src, std::string_view skip_entry) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    if (std::string_view(name_buf) == skip_entry) {
      continue;
    }
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);

  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlReader, RejectsArchiveWithoutWorkbookXml) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithout(ok_bytes, "xl/workbook.xml");

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(OoxmlReader, RejectsArchiveWithoutContentTypes) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithout(ok_bytes, "[Content_Types].xml");

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(OoxmlReader, RejectsArchiveWithoutPackageRels) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithout(ok_bytes, "_rels/.rels");

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoRelationshipBroken);
}

TEST(OoxmlReader, UnknownPartsExcludesSheetAndStylesAndSst) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const std::vector<PassthroughPart>& parts = result_or.value().workbook.passthrough_parts();

  // The reader consumes sheet*.xml and xl/styles.xml. The latter is a
  // best-effort validate-only parse but still counts as consumed so the
  // round-trip pipeline does not surface it as unknown. The empty
  // workbook does not emit a sharedStrings part, so the assertion below
  // is just sanity: it must not appear either.
  auto path_eq = [&parts](std::string_view target) {
    return std::find_if(parts.begin(), parts.end(), [target](const PassthroughPart& p) { return p.path == target; });
  };
  EXPECT_EQ(path_eq("xl/styles.xml"), parts.end()) << "styles.xml unexpectedly still in unknown_parts";
  EXPECT_EQ(path_eq("xl/worksheets/sheet1.xml"), parts.end()) << "sheet1.xml unexpectedly still in unknown_parts";
  EXPECT_EQ(path_eq("xl/sharedStrings.xml"), parts.end()) << "sharedStrings.xml unexpectedly in unknown_parts";
}

// ---------------------------------------------------------------------------
// Path-traversal hardening for `ResolveRelativePath`. A hostile rels file
// with a Target like `../../../etc/passwd` must not be allowed to walk
// out of the package root: even though the resulting path never reaches
// the filesystem (the reader looks it up in the in-memory ZIP catalogue),
// allowing the value to flow through to downstream stages — `has_entry`,
// passthrough emission, link resolution — invites confusion-shaped bugs
// at minimum and a real escape if any caller ever maps the path back to
// disk.
// ---------------------------------------------------------------------------

TEST(OoxmlReader, ResolveRelativePathRefusesEscapingTarget) {
  // base_dir = "xl" + target = "../../etc/passwd": one ".." pops "xl",
  // the second ".." escapes the package root and must trigger
  // kIoZipSlip.
  auto result = internal::ResolveRelativePathForTesting("xl", "../../etc/passwd");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
  // Diagnostic context must echo both arguments so the offending input
  // is recoverable from logs.
  EXPECT_NE(result.error().context.find("base_dir=xl"), std::string::npos);
  EXPECT_NE(result.error().context.find("target=../../etc/passwd"), std::string::npos);
}

TEST(OoxmlReader, ResolveRelativePathRefusesPackageAbsoluteTarget) {
  // A package-absolute Target ("/xl/...") bypasses base-dir-relative
  // accounting entirely; we refuse it rather than silently treat it as
  // root-relative.
  auto result = internal::ResolveRelativePathForTesting("xl/_rels", "/etc/passwd");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(OoxmlReader, ResolveRelativePathAllowsValidParentRefs) {
  // A single ".." that stays inside the package is legitimate:
  // xl/worksheets/_rels/sheet1.xml.rels Target="../../theme/theme1.xml"
  // resolves to "xl/theme/theme1.xml".
  auto result = internal::ResolveRelativePathForTesting("xl/worksheets", "../theme/theme1.xml");
  ASSERT_TRUE(static_cast<bool>(result)) << "rejected legitimate parent-relative target";
  EXPECT_EQ(result.value(), "xl/theme/theme1.xml");
}

TEST(OoxmlReader, ResolveRelativePathRefusesEmptyResult) {
  // Pop every directory off the stack but stop short of escaping; the
  // resolver previously returned an empty string here, which downstream
  // callers would happily store into rid->path maps. We now refuse
  // empty results outright.
  auto result = internal::ResolveRelativePathForTesting("xl", "..");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
}

/// Builds an `xl/_rels/workbook.xml.rels` body whose worksheet
/// relationship Target attempts to escape the package root. Used to
/// drive the end-to-end rejection test below.
std::vector<std::uint8_t> RebuildArchiveWithMaliciousWorkbookRels(const std::vector<std::uint8_t>& src) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  static constexpr std::string_view kMaliciousRels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"../../etc/passwd\"/>"
      "</Relationships>";

  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    const std::string_view name(name_buf);
    if (name == "xl/_rels/workbook.xml.rels") {
      EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, kMaliciousRels.data(), kMaliciousRels.size(),
                                      static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
                MZ_FALSE);
      continue;
    }
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlReader, RejectsArchiveWithEscapingRelsTarget) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithMaliciousWorkbookRels(ok_bytes);

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(OoxmlReader, IsSafePartNameAcceptsCanonicalNames) {
  EXPECT_TRUE(ooxml::is_safe_part_name("xl/workbook.xml"));
  EXPECT_TRUE(ooxml::is_safe_part_name("xl/media/image1.png"));
  EXPECT_TRUE(ooxml::is_safe_part_name("[Content_Types].xml"));
}

TEST(OoxmlReader, IsSafePartNameRejectsTraversalAndAbsoluteShapes) {
  EXPECT_FALSE(ooxml::is_safe_part_name(""));
  EXPECT_FALSE(ooxml::is_safe_part_name("/etc/passwd"));            // package-absolute
  EXPECT_FALSE(ooxml::is_safe_part_name("../../etc/passwd"));       // parent traversal
  EXPECT_FALSE(ooxml::is_safe_part_name("xl/../../../etc/x"));      // embedded traversal
  EXPECT_FALSE(ooxml::is_safe_part_name("xl\\worksheets\\a.xml"));  // backslash separator
  EXPECT_FALSE(ooxml::is_safe_part_name("C:/Windows/system32"));    // drive-letter colon
  EXPECT_FALSE(ooxml::is_safe_part_name("xl//workbook.xml"));       // empty segment
  EXPECT_FALSE(ooxml::is_safe_part_name("xl/./workbook.xml"));      // dot segment
}

/// Rebuilds a valid archive with one extra archive entry whose name
/// escapes the package root (`../evil.bin`). The reader's Default-typed
/// sweep must refuse to carry it through passthrough.
std::vector<std::uint8_t> RebuildArchiveWithMaliciousExtraEntry(const std::vector<std::uint8_t>& src) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);

  static constexpr std::string_view kEvil = "malicious";
  EXPECT_NE(mz_zip_writer_add_mem(&writer, "../evil.bin", kEvil.data(), kEvil.size(),
                                  static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
            MZ_FALSE);

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlReader, RejectsArchiveWithTraversalShapedPartName) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithMaliciousExtraEntry(ok_bytes);

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(OoxmlReader, RejectsEncryptedCdfv2ContainerWithEncryptedDiagnostic) {
  // OLE/CDFV2 compound-file signature that wraps a password-protected
  // .xlsx/.xlsb, followed by arbitrary filler. This must surface an
  // "encrypted" diagnostic rather than a generic "corrupt zip".
  std::vector<std::uint8_t> encrypted = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
  encrypted.resize(512, 0x00);

  auto result_or = read_ooxml(SpanOf(encrypted));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipEncrypted);
}

TEST(OoxmlReader, EmptyWorkbookFactoryProducesZeroSheets) {
  Workbook wb = Workbook::create_empty();
  EXPECT_EQ(wb.sheet_count(), 0U);
  // save() should reject zero-sheet workbooks (Excel does the same), so
  // the input cannot be round-tripped: the test stops at the factory
  // contract.
  auto save_or = wb.save();
  EXPECT_FALSE(static_cast<bool>(save_or));
}

}  // namespace
}  // namespace io
}  // namespace formulon
