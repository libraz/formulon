//
// Unit tests for `formulon::io::ZipReader`. Each test produces an
// in-memory `.xlsx` via the existing `Workbook::save()` writer and then
// reads it back through `ZipReader` to assert parity with miniz's raw
// `mz_zip_reader_*` API surface.

#include "io/zip_reader.h"

#include <array>
#include <cstdint>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "miniz.h"
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

/// Marks every local and central header in an otherwise valid ZIP as
/// encrypted. miniz deliberately does not decrypt such an entry, which
/// gives the error-mapping test a small deterministic fixture without a
/// password-capable ZIP writer.
void MarkZipEntriesEncrypted(std::vector<std::uint8_t>& bytes) {
  for (std::size_t i = 0; i + 8U <= bytes.size(); ++i) {
    const bool local = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 3U && bytes[i + 3U] == 4U;
    const bool central = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 1U && bytes[i + 3U] == 2U;
    if (!local && !central) {
      continue;
    }
    const std::size_t flag_offset = i + (local ? 6U : 8U);
    bytes[flag_offset] = static_cast<std::uint8_t>(bytes[flag_offset] | 0x01U);
  }
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

TEST(ZipReader, ReadEntryReportsEncryptedPartDistinctly) {
  std::vector<std::uint8_t> bytes = MakeMinimalXlsx();
  MarkZipEntriesEncrypted(bytes);

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  auto result = zip.read_entry("xl/workbook.xml");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipEncrypted);
  EXPECT_NE(result.error().context.find("entry=xl/workbook.xml"), std::string::npos);
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

/// Builds an in-memory ZIP archive containing a single entry whose
/// uncompressed payload is `payload`. Returned bytes can be fed straight
/// to `ZipReader::open`. Used to construct the > 100 MiB ZIP-bomb-shaped
/// fixture below; kept local to this TU because production code never
/// needs to round-trip raw bytes through miniz's writer surface.
std::vector<std::uint8_t> BuildSingleEntryZip(const std::string& name, const std::vector<std::uint8_t>& payload,
                                              mz_uint compression = static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)) {
  mz_zip_archive archive{};
  std::memset(&archive, 0, sizeof(archive));
  EXPECT_EQ(mz_zip_writer_init_heap(&archive, 0, 0), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_add_mem(&archive, name.c_str(), payload.data(), payload.size(), compression), MZ_TRUE);
  void* heap = nullptr;
  std::size_t heap_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&archive, &heap, &heap_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&archive), MZ_TRUE);
  std::vector<std::uint8_t> out(heap_size);
  std::memcpy(out.data(), heap, heap_size);
  mz_free(heap);
  return out;
}

/// Builds an in-memory ZIP containing `count` empty entries named
/// `entry_<n>.bin`. Used to drive the per-archive entry-count cap check.
std::vector<std::uint8_t> BuildManyEmptyEntriesZip(std::size_t count) {
  mz_zip_archive archive{};
  std::memset(&archive, 0, sizeof(archive));
  EXPECT_EQ(mz_zip_writer_init_heap(&archive, 0, 0), MZ_TRUE);
  const std::array<std::uint8_t, 1> tiny{{0x00u}};
  for (std::size_t i = 0; i < count; ++i) {
    std::string name = "entry_" + std::to_string(i) + ".bin";
    EXPECT_EQ(mz_zip_writer_add_mem(&archive, name.c_str(), tiny.data(), tiny.size(),
                                    static_cast<mz_uint>(MZ_NO_COMPRESSION)),
              MZ_TRUE);
  }
  void* heap = nullptr;
  std::size_t heap_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&archive, &heap, &heap_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&archive), MZ_TRUE);
  std::vector<std::uint8_t> out(heap_size);
  std::memcpy(out.data(), heap, heap_size);
  mz_free(heap);
  return out;
}

/// Builds an in-memory ZIP containing `count` entries that each carry the
/// same `payload`. All entries share the supplied compression level so
/// the per-entry cumulative-cap test can drive the running total without
/// hitting the per-entry size cap.
std::vector<std::uint8_t> BuildMultiEntryZip(std::size_t count, const std::vector<std::uint8_t>& payload,
                                             mz_uint compression = static_cast<mz_uint>(MZ_NO_COMPRESSION)) {
  mz_zip_archive archive{};
  std::memset(&archive, 0, sizeof(archive));
  EXPECT_EQ(mz_zip_writer_init_heap(&archive, 0, 0), MZ_TRUE);
  for (std::size_t i = 0; i < count; ++i) {
    std::string name = "blob_" + std::to_string(i) + ".bin";
    EXPECT_EQ(mz_zip_writer_add_mem(&archive, name.c_str(), payload.data(), payload.size(), compression), MZ_TRUE);
  }
  void* heap = nullptr;
  std::size_t heap_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&archive, &heap, &heap_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&archive), MZ_TRUE);
  std::vector<std::uint8_t> out(heap_size);
  std::memcpy(out.data(), heap, heap_size);
  mz_free(heap);
  return out;
}

TEST(ZipReader, ReadEntryRejectsZipBombSizedEntry) {
  // 101 MiB of zero bytes compresses to a few KB but reports an
  // uncompressed size that exceeds the per-entry cap; the reader must
  // refuse extraction before allocating the full payload.
  constexpr std::size_t kPayloadBytes = 101ULL * 1024ULL * 1024ULL;
  std::vector<std::uint8_t> payload(kPayloadBytes, 0u);
  const std::vector<std::uint8_t> bytes = BuildSingleEntryZip("huge.bin", payload);

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  ASSERT_TRUE(zip.has_entry("huge.bin"));

  auto result = zip.read_entry("huge.bin");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipBomb);
  // Context must mention the offending entry so callers can attribute
  // the rejection in logs.
  EXPECT_NE(result.error().context.find("entry=huge.bin"), std::string::npos);
}

TEST(ZipReader, ReadEntryAcceptsBelowCapEntry) {
  // 1 MiB sits comfortably under the 100 MiB cap; round-trip succeeds.
  constexpr std::size_t kPayloadBytes = 1024ULL * 1024ULL;
  std::vector<std::uint8_t> payload(kPayloadBytes, 0xAAu);
  const std::vector<std::uint8_t> bytes = BuildSingleEntryZip("ok.bin", payload);

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));

  auto result = zip.read_entry("ok.bin");
  ASSERT_TRUE(static_cast<bool>(result)) << "read_entry failed: " << result.error().message;
  EXPECT_EQ(result.value().size(), kPayloadBytes);
}

TEST(ZipReader, OpenRejectsTooManyParts) {
  // One entry past the cap is enough to flip `kIoZipBomb`; the writer
  // produces a structurally-valid archive so the rejection is solely on
  // policy grounds.
  const std::vector<std::uint8_t> bytes = BuildManyEmptyEntriesZip(kMaxParts + 1);
  ZipReader zip;
  auto result = zip.open(SpanOf(bytes));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipBomb);
  EXPECT_NE(result.error().context.find("limit=parts"), std::string::npos);
  // The reader must surface as not-open after a refused open(); subsequent
  // reads return a not-open / not-found error rather than aliasing into a
  // half-initialised miniz handle.
  EXPECT_EQ(zip.entry_count(), 0U);
}

TEST(ZipReader, ReadEntryRejectsHighRatioEntry) {
  // 4 MiB of zero bytes compresses to roughly 4 KiB with miniz at
  // best-compression: ratio comfortably exceeds the 1024 cap. (200 KiB
  // was on the edge — deflate's per-block bookkeeping kept the ratio
  // around ~800 in practice.)
  constexpr std::size_t kPayloadBytes = 4ULL * 1024ULL * 1024ULL;
  std::vector<std::uint8_t> payload(kPayloadBytes, 0u);
  const std::vector<std::uint8_t> bytes =
      BuildSingleEntryZip("zeros.bin", payload, static_cast<mz_uint>(MZ_BEST_COMPRESSION));

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  ASSERT_TRUE(zip.has_entry("zeros.bin"));

  auto result = zip.read_entry("zeros.bin");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipBomb);
  EXPECT_NE(result.error().context.find("limit=ratio"), std::string::npos);
  EXPECT_NE(result.error().context.find("entry=zeros.bin"), std::string::npos);
}

TEST(ZipReader, ReadEntryRejectsCumulativeTotal) {
  // Six 50 MiB entries: each individually under the per-entry 100 MiB
  // cap, but together they cross the 256 MiB cumulative cap on the sixth
  // read. Compress the archive at default level so the on-disk archive
  // bytes stay small (~50 KiB) while the advertised uncompressed sizes
  // still drive the cumulative-total check; the ratio cap (1024) is
  // satisfied by 50 MiB / 50 KiB ~= 1024 only when we use uniform-fill
  // data, so we use a non-uniform pattern that compresses to ~5 MiB
  // (ratio ~10) — comfortably under the ratio cap and still light on
  // build-time memory.
  constexpr std::size_t kPayloadBytes = 50ULL * 1024ULL * 1024ULL;
  std::vector<std::uint8_t> payload(kPayloadBytes, 0u);
  // Fill with a low-entropy but non-uniform pattern: ratio ~10 with the
  // default deflate level, well under the 1024 ratio cap.
  for (std::size_t i = 0; i < payload.size(); ++i) {
    payload[i] = static_cast<std::uint8_t>(i & 0xFFu);
  }
  const std::vector<std::uint8_t> bytes =
      BuildMultiEntryZip(/*count=*/6, payload, /*compression=*/static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  ASSERT_EQ(zip.entry_count(), 6U);

  // First five reads succeed: 5 * 50 MiB == 250 MiB <= 256 MiB cap.
  for (int i = 0; i < 5; ++i) {
    const std::string name = "blob_" + std::to_string(i) + ".bin";
    auto ok = zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(ok)) << "read_entry failed at i=" << i << ": " << ok.error().message;
    EXPECT_EQ(ok.value().size(), kPayloadBytes);
  }

  // Sixth read would push cumulative past 256 MiB: 250 MiB + 50 MiB.
  auto refused = zip.read_entry("blob_5.bin");
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kIoZipBomb);
  EXPECT_NE(refused.error().context.find("limit=total"), std::string::npos);

  // Re-opening the same reader resets the running total so a fresh
  // archive is not penalised by a previous run's history.
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  auto reopened = zip.read_entry("blob_0.bin");
  ASSERT_TRUE(static_cast<bool>(reopened)) << "post-reopen read_entry failed: " << reopened.error().message;
  EXPECT_EQ(reopened.value().size(), kPayloadBytes);
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
