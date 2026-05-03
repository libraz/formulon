// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `formulon::io::ZipReader`. Each test produces an
// in-memory `.xlsx` via the existing `Workbook::save()` writer and then
// reads it back through `ZipReader` to assert parity with miniz's raw
// `mz_zip_reader_*` API surface.

#include "io/zip_reader.h"

#include <array>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Produces a minimal in-memory `.xlsx` we can use as a test corpus.
/// Returned vector outlives any `ZipReader` constructed against it.
std::vector<std::uint8_t> MakeMinimalXlsx() {
  Workbook wb = Workbook::create();
  auto save_result = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_result)) << "save() failed in test setup";
  return save_result.value();
}

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

TEST(ZipReader, OpenSucceedsOnWriterOutput) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  auto result = zip.open(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result)) << "open failed: " << result.error().message;
}

TEST(ZipReader, EntryCountMatchesWriterParts) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  // Empty-workbook writer emits exactly six parts: [Content_Types].xml,
  // _rels/.rels, xl/workbook.xml, xl/_rels/workbook.xml.rels,
  // xl/worksheets/sheet1.xml, xl/styles.xml.
  EXPECT_EQ(zip.entry_count(), 6U);
}

TEST(ZipReader, HasEntryFindsKnownParts) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  EXPECT_TRUE(zip.has_entry("[Content_Types].xml"));
  EXPECT_TRUE(zip.has_entry("_rels/.rels"));
  EXPECT_TRUE(zip.has_entry("xl/workbook.xml"));
  EXPECT_TRUE(zip.has_entry("xl/_rels/workbook.xml.rels"));
  EXPECT_TRUE(zip.has_entry("xl/worksheets/sheet1.xml"));
  EXPECT_TRUE(zip.has_entry("xl/styles.xml"));

  EXPECT_FALSE(zip.has_entry("missing.xml"));
  EXPECT_FALSE(zip.has_entry(""));
  // Case-sensitive: ZIP entry names retain their exact byte sequence.
  EXPECT_FALSE(zip.has_entry("XL/WORKBOOK.XML"));
}

TEST(ZipReader, ReadEntryReturnsDecompressedBytes) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  auto wb_or = zip.read_entry("xl/workbook.xml");
  ASSERT_TRUE(static_cast<bool>(wb_or)) << "read_entry failed: " << wb_or.error().message;
  const std::vector<std::uint8_t>& wb = wb_or.value();
  EXPECT_GT(wb.size(), 0U);
  // The body must be a valid XML document (we do not parse here, just
  // sanity-check the prologue bytes).
  ASSERT_GE(wb.size(), 5U);
  EXPECT_EQ(wb[0], '<');
  EXPECT_EQ(wb[1], '?');
  EXPECT_EQ(wb[2], 'x');
  EXPECT_EQ(wb[3], 'm');
  EXPECT_EQ(wb[4], 'l');
}

TEST(ZipReader, ReadEntryReportsMissingPart) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  auto missing = zip.read_entry("nope.xml");
  ASSERT_FALSE(static_cast<bool>(missing));
  EXPECT_EQ(missing.error().code, FormulonErrorCode::kIoFileNotFound);
}

TEST(ZipReader, OpenRejectsNonZipBuffer) {
  const std::array<std::uint8_t, 4> garbage{{0u, 0u, 0u, 0u}};
  const ByteSpan span{garbage.data(), garbage.size()};
  ZipReader zip;
  auto result = zip.open(span);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipCorrupt);
}

TEST(ZipReader, OpenRejectsEmptyBuffer) {
  const ByteSpan span{nullptr, 0};
  ZipReader zip;
  auto result = zip.open(span);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipCorrupt);
}

TEST(ZipReader, ListEntriesReturnsAllNames) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  std::vector<std::string> names = zip.list_entries();
  ASSERT_EQ(names.size(), 6U);

  // The set of names must match what the writer emitted, regardless of
  // order. (We assert as a set so future writer reordering does not
  // break this test.)
  bool saw_ct = false;
  bool saw_rels = false;
  bool saw_workbook = false;
  bool saw_workbook_rels = false;
  bool saw_sheet1 = false;
  bool saw_styles = false;
  for (const std::string& n : names) {
    if (n == "[Content_Types].xml")
      saw_ct = true;
    if (n == "_rels/.rels")
      saw_rels = true;
    if (n == "xl/workbook.xml")
      saw_workbook = true;
    if (n == "xl/_rels/workbook.xml.rels")
      saw_workbook_rels = true;
    if (n == "xl/worksheets/sheet1.xml")
      saw_sheet1 = true;
    if (n == "xl/styles.xml")
      saw_styles = true;
  }
  EXPECT_TRUE(saw_ct);
  EXPECT_TRUE(saw_rels);
  EXPECT_TRUE(saw_workbook);
  EXPECT_TRUE(saw_workbook_rels);
  EXPECT_TRUE(saw_sheet1);
  EXPECT_TRUE(saw_styles);
}

TEST(ZipReader, EntryNameOutOfRangeIsEmpty) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  EXPECT_EQ(zip.entry_name(zip.entry_count()).size(), 0U);
  EXPECT_EQ(zip.entry_name(zip.entry_count() + 100).size(), 0U);
}

TEST(ZipReader, MoveConstructorPreservesOpenState) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader src;
  ASSERT_TRUE(static_cast<bool>(src.open(SpanOf(bytes))));
  ASSERT_EQ(src.entry_count(), 6U);

  ZipReader dst(std::move(src));
  EXPECT_EQ(dst.entry_count(), 6U);
  EXPECT_TRUE(dst.has_entry("xl/workbook.xml"));

  auto body_or = dst.read_entry("xl/workbook.xml");
  ASSERT_TRUE(static_cast<bool>(body_or));
  EXPECT_GT(body_or.value().size(), 0U);
}

TEST(ZipReader, MoveAssignmentPreservesOpenState) {
  const std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  ZipReader src;
  ASSERT_TRUE(static_cast<bool>(src.open(SpanOf(bytes))));

  ZipReader dst;
  dst = std::move(src);
  EXPECT_EQ(dst.entry_count(), 6U);
  EXPECT_TRUE(dst.has_entry("xl/styles.xml"));
}

TEST(ZipReader, IdempotentReopen) {
  const std::vector<std::uint8_t> bytes_a = MakeMinimalXlsx();
  Workbook wb_b = Workbook::create();
  wb_b.add_sheet("Two");
  auto save_b_or = wb_b.save();
  ASSERT_TRUE(static_cast<bool>(save_b_or));
  const std::vector<std::uint8_t> bytes_b = save_b_or.value();

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_a))));
  EXPECT_EQ(zip.entry_count(), 6U);  // single sheet

  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_b))));
  EXPECT_EQ(zip.entry_count(), 7U);  // two-sheet variant: extra sheet2.xml
  EXPECT_TRUE(zip.has_entry("xl/worksheets/sheet2.xml"));
}

}  // namespace
}  // namespace io
}  // namespace formulon
