// Copyright 2026 libraz. Licensed under the MIT License.
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

#include "gtest/gtest.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) { return ByteSpan{bytes.data(), bytes.size()}; }

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

TEST(OoxmlReader, JapaneseSheetNameSurvivesRoundTrip) {
  Workbook wb = Workbook::create();
  // "日本シート" in UTF-8: E6 97 A5 E6 9C AC E3 82 B7 E3 83 BC E3 83 88
  wb.sheet(0).set_name("\xE6\x97\xA5\xE6\x9C\xAC\xE3\x82\xB7\xE3\x83\xBC\xE3\x83\x88");
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(),
            "\xE6\x97\xA5\xE6\x9C\xAC\xE3\x82\xB7\xE3\x83\xBC\xE3\x83\x88");
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

TEST(OoxmlReader, UnknownPartsContainsStylesButNotSheetXml) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const std::vector<std::string>& parts = result_or.value().unknown_parts;

  // The reader now consumes sheet*.xml as well, so the worksheet part
  // must NOT appear in `unknown_parts` (we parsed it). The stylesheet
  // is still on the deferred list and remains here for future bundles
  // to round-trip.
  EXPECT_NE(std::find(parts.begin(), parts.end(), "xl/styles.xml"), parts.end()) << "styles.xml not in unknown_parts";
  EXPECT_EQ(std::find(parts.begin(), parts.end(), "xl/worksheets/sheet1.xml"), parts.end())
      << "sheet1.xml unexpectedly still in unknown_parts";
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
