//
// Unit tests for `formulon::io::ZipReader`. Each test produces an
// in-memory `.xlsx` via the existing `Workbook::save()` writer and then
// reads it back through `ZipReader` to assert parity with miniz's raw
// `mz_zip_reader_*` API surface.

#include "io/zip_reader.h"

#include <array>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <string_view>
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

/// Marks only the local and central headers for `name` as encrypted. The
/// payload remains ordinary plaintext; miniz rejects the entry from its
/// flags before attempting to interpret those bytes.
bool MarkZipEntryEncrypted(std::vector<std::uint8_t>& bytes, const std::string& name) {
  bool found = false;
  for (std::size_t i = 0; i + 8U <= bytes.size(); ++i) {
    const bool local = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 3U && bytes[i + 3U] == 4U;
    const bool central = bytes[i] == 'P' && bytes[i + 1U] == 'K' && bytes[i + 2U] == 1U && bytes[i + 3U] == 2U;
    if (!local && !central) {
      continue;
    }
    const std::size_t filename_length_offset = i + (local ? 26U : 28U);
    const std::size_t filename_offset = i + (local ? 30U : 46U);
    if (filename_length_offset + 2U > bytes.size()) {
      return false;
    }
    const std::size_t filename_length = static_cast<std::size_t>(bytes[filename_length_offset]) |
                                        (static_cast<std::size_t>(bytes[filename_length_offset + 1U]) << 8U);
    if (filename_offset + filename_length > bytes.size()) {
      return false;
    }
    if (std::string(reinterpret_cast<const char*>(bytes.data() + filename_offset), filename_length) != name) {
      continue;
    }
    const std::size_t flag_offset = i + (local ? 6U : 8U);
    bytes[flag_offset] = static_cast<std::uint8_t>(bytes[flag_offset] | 0x01U);
    found = true;
  }
  return found;
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

/// Builds a structurally valid ZIP whose central directory contains the same
/// non-directory name twice. The production reader rejects this before any
/// caller can observe miniz's ambiguous first/last-match behaviour.
std::vector<std::uint8_t> BuildDuplicateEntryZip() {
  mz_zip_archive archive{};
  std::memset(&archive, 0, sizeof(archive));
  EXPECT_EQ(mz_zip_writer_init_heap(&archive, 0, 0), MZ_TRUE);
  const std::array<std::uint8_t, 3> first{{'o', 'n', 'e'}};
  const std::array<std::uint8_t, 3> second{{'t', 'w', 'o'}};
  EXPECT_EQ(mz_zip_writer_add_mem(&archive, "duplicate.bin", first.data(), first.size(), MZ_NO_COMPRESSION), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_add_mem(&archive, "duplicate.bin", second.data(), second.size(), MZ_NO_COMPRESSION), MZ_TRUE);
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

/// Builds an archive from entry names and borrowed payloads. This keeps the
/// exact-boundary cumulative-budget fixture from making four independent
/// 64 MiB copies in the test process.
std::vector<std::uint8_t> BuildEntriesZip(
    const std::vector<std::pair<std::string, const std::vector<std::uint8_t>*>>& entries,
    mz_uint compression = static_cast<mz_uint>(MZ_NO_COMPRESSION)) {
  mz_zip_archive archive{};
  std::memset(&archive, 0, sizeof(archive));
  EXPECT_EQ(mz_zip_writer_init_heap(&archive, 0, 0), MZ_TRUE);
  for (const auto& entry : entries) {
    EXPECT_NE(entry.second, nullptr);
    if (entry.second == nullptr) {
      continue;
    }
    EXPECT_EQ(
        mz_zip_writer_add_mem(&archive, entry.first.c_str(), entry.second->data(), entry.second->size(), compression),
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

/// The cumulative-budget tests exercise the boundary through `open()`'s
/// lower-only ceiling parameter rather than the 256 MiB production constant.
/// The shape under test is identical — four chunks reach the ceiling exactly
/// and a trailing one-byte entry proves it is exhausted — but the arithmetic
/// runs on 4 MiB instead of 256 MiB, which keeps a security boundary in the
/// fast tier instead of pushing 768 MiB of decompression through it.
constexpr std::size_t kCumulativeChunkBytes = 1ULL * 1024ULL * 1024ULL;
constexpr std::size_t kCumulativeCeilingBytes = 4ULL * kCumulativeChunkBytes;

/// Renders the `used=/requested=/ceiling=` tail of a cumulative-budget error
/// context, so the expectations stay tied to the constants above.
std::string CumulativeOverflowContext(std::string_view entry, std::size_t used, std::size_t requested) {
  return "limit=total entry=" + std::string(entry) + " used=" + std::to_string(used) +
         " requested=" + std::to_string(requested) + " ceiling=" + std::to_string(kCumulativeCeilingBytes);
}

/// Builds the exact-boundary archive used by cumulative-budget tests. The
/// archive is small on disk because the repeated pattern compresses well, but
/// each central-directory entry still advertises a full chunk of extraction.
std::vector<std::uint8_t> BuildCumulativeBoundaryZip() {
  std::vector<std::uint8_t> payload(kCumulativeChunkBytes, 0u);
  for (std::size_t i = 0; i < payload.size(); ++i) {
    payload[i] = static_cast<std::uint8_t>(i & 0xFFu);
  }
  const std::vector<std::uint8_t> one_byte{0xA5u};
  std::vector<std::pair<std::string, const std::vector<std::uint8_t>*>> entries;
  for (int i = 0; i < 4; ++i) {
    entries.emplace_back("blob_" + std::to_string(i) + ".bin", &payload);
  }
  entries.emplace_back("one.bin", &one_byte);
  return BuildEntriesZip(entries, static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
}

void WriteLittleEndian32(std::vector<std::uint8_t>& bytes, std::size_t offset, std::uint32_t value) {
  ASSERT_LE(offset + 4U, bytes.size());
  bytes[offset] = static_cast<std::uint8_t>(value & 0xFFU);
  bytes[offset + 1U] = static_cast<std::uint8_t>((value >> 8U) & 0xFFU);
  bytes[offset + 2U] = static_cast<std::uint8_t>((value >> 16U) & 0xFFU);
  bytes[offset + 3U] = static_cast<std::uint8_t>((value >> 24U) & 0xFFU);
}

/// Rewrites central-directory size fields for a policy-priority fixture. The
/// reader's policy checks intentionally precede local-header validation and
/// extraction, so tiny payloads can exercise advertised oversized/ratio
/// metadata without allocating another 100 MiB test vector.
bool RewriteCentralDirectorySizes(std::vector<std::uint8_t>& bytes, const std::string& name,
                                  std::uint32_t compressed_size, std::uint32_t uncompressed_size) {
  for (std::size_t i = 0; i + 46U <= bytes.size(); ++i) {
    if (bytes[i] != 'P' || bytes[i + 1U] != 'K' || bytes[i + 2U] != 1U || bytes[i + 3U] != 2U) {
      continue;
    }
    const std::size_t filename_length =
        static_cast<std::size_t>(bytes[i + 28U]) | (static_cast<std::size_t>(bytes[i + 29U]) << 8U);
    const std::size_t filename_offset = i + 46U;
    if (filename_offset + filename_length > bytes.size()) {
      return false;
    }
    if (std::string(reinterpret_cast<const char*>(bytes.data() + filename_offset), filename_length) != name) {
      continue;
    }
    WriteLittleEndian32(bytes, i + 20U, compressed_size);
    WriteLittleEndian32(bytes, i + 24U, uncompressed_size);
    return true;
  }
  return false;
}

/// Changes only the central-directory CRC for `name`, leaving the local
/// header and payload intact. miniz then rejects extraction after the reader
/// has charged the advertised size, which locks the failed-extraction budget
/// contract without making a malformed archive fail during `open()`.
bool CorruptCentralDirectoryCrc(std::vector<std::uint8_t>& bytes, const std::string& name) {
  for (std::size_t i = 0; i + 46U <= bytes.size(); ++i) {
    if (bytes[i] != 'P' || bytes[i + 1U] != 'K' || bytes[i + 2U] != 1U || bytes[i + 3U] != 2U) {
      continue;
    }
    const std::size_t filename_length =
        static_cast<std::size_t>(bytes[i + 28U]) | (static_cast<std::size_t>(bytes[i + 29U]) << 8U);
    const std::size_t filename_offset = i + 46U;
    if (filename_offset + filename_length > bytes.size()) {
      return false;
    }
    if (std::string(reinterpret_cast<const char*>(bytes.data() + filename_offset), filename_length) == name) {
      bytes[i + 16U] ^= 0x01U;
      return true;
    }
  }
  return false;
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

TEST(ZipReader, OpenRejectsDuplicateNonDirectoryNames) {
  const std::vector<std::uint8_t> bytes = BuildDuplicateEntryZip();
  ZipReader zip;
  auto result = zip.open(SpanOf(bytes));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipCorrupt);
  EXPECT_NE(result.error().context.find("duplicate_entry=duplicate.bin"), std::string::npos);
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
  // Four chunks land exactly on the session ceiling. A same-entry reread and a
  // one-byte fifth entry both fail before allocation; the latter makes the +1
  // boundary explicit. Reopening resets the budget.
  const std::vector<std::uint8_t> bytes = BuildCumulativeBoundaryZip();

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes), kCumulativeCeilingBytes)));
  ASSERT_EQ(zip.entry_count(), 5U);

  for (int i = 0; i < 4; ++i) {
    const std::string name = "blob_" + std::to_string(i) + ".bin";
    auto ok = zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(ok)) << "read_entry failed at i=" << i << ": " << ok.error().message;
    EXPECT_EQ(ok.value().size(), kCumulativeChunkBytes);
  }

  auto reread = zip.read_entry("blob_0.bin");
  ASSERT_FALSE(static_cast<bool>(reread));
  EXPECT_EQ(reread.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_EQ(reread.error().context,
            CumulativeOverflowContext("blob_0.bin", kCumulativeCeilingBytes, kCumulativeChunkBytes));

  auto over = zip.read_entry("one.bin");
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_EQ(over.error().context, CumulativeOverflowContext("one.bin", kCumulativeCeilingBytes, 1U));

  // A miniz validation failure after the budget charge must also consume the
  // advertised bytes. The corrupted first entry plus three valid entries
  // reaches the same exact ceiling, so the trailing one-byte entry proves
  // that the failed extraction was not refunded.
  std::vector<std::uint8_t> corrupt_bytes = bytes;
  ASSERT_TRUE(CorruptCentralDirectoryCrc(corrupt_bytes, "blob_0.bin"));
  ZipReader corrupt_zip;
  ASSERT_TRUE(static_cast<bool>(corrupt_zip.open(SpanOf(corrupt_bytes), kCumulativeCeilingBytes)));
  auto failed_extract = corrupt_zip.read_entry("blob_0.bin");
  ASSERT_FALSE(static_cast<bool>(failed_extract));
  EXPECT_EQ(failed_extract.error().code, FormulonErrorCode::kIoZipCorrupt);
  for (int i = 1; i < 4; ++i) {
    const std::string name = "blob_" + std::to_string(i) + ".bin";
    auto ok = corrupt_zip.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(ok)) << "read_entry failed at i=" << i << ": " << ok.error().message;
    EXPECT_EQ(ok.value().size(), kCumulativeChunkBytes);
  }
  auto charged_failure_over = corrupt_zip.read_entry("one.bin");
  ASSERT_FALSE(static_cast<bool>(charged_failure_over));
  EXPECT_EQ(charged_failure_over.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_EQ(charged_failure_over.error().context, CumulativeOverflowContext("one.bin", kCumulativeCeilingBytes, 1U));

  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes), kCumulativeCeilingBytes)));
  auto reopened = zip.read_entry("one.bin");
  ASSERT_TRUE(static_cast<bool>(reopened)) << "post-reopen read_entry failed: " << reopened.error().message;
  ASSERT_EQ(reopened.value().size(), 1U);
  EXPECT_EQ(reopened.value()[0], 0xA5u);
}

TEST(ZipReader, OpenCeilingOnlyTightensTheCumulativeBudget) {
  const std::vector<std::uint8_t> bytes = BuildCumulativeBoundaryZip();

  // A ceiling below the default is honoured: two chunks exhaust it.
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes), 2U * kCumulativeChunkBytes)));
  for (int i = 0; i < 2; ++i) {
    const std::string name = "blob_" + std::to_string(i) + ".bin";
    ASSERT_TRUE(static_cast<bool>(zip.read_entry(name))) << "read_entry failed at i=" << i;
  }
  auto over = zip.read_entry("blob_2.bin");
  ASSERT_FALSE(static_cast<bool>(over));
  EXPECT_EQ(over.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_NE(over.error().context.find("ceiling=" + std::to_string(2U * kCumulativeChunkBytes)), std::string::npos);

  // A request above the default cannot raise it. The archive stays well inside
  // `kMaxTotalExtractedBytes`, so every entry reads: what is pinned here is
  // that the oversized request is accepted and clamped rather than rejected or
  // honoured verbatim. Driving the clamped ceiling to overflow would take a
  // 256 MiB extraction, which does not belong in the fast tier.
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes), std::numeric_limits<std::size_t>::max())));
  for (int i = 0; i < 4; ++i) {
    const std::string name = "blob_" + std::to_string(i) + ".bin";
    ASSERT_TRUE(static_cast<bool>(zip.read_entry(name))) << "post-clamp read failed at i=" << i;
  }
  EXPECT_TRUE(static_cast<bool>(zip.read_entry("one.bin")));
}

TEST(ZipReader, CumulativeBudgetSurvivesMoveConstructionAndAssignment) {
  const std::vector<std::uint8_t> bytes = BuildCumulativeBoundaryZip();

  ZipReader construct_source;
  ASSERT_TRUE(static_cast<bool>(construct_source.open(SpanOf(bytes), kCumulativeCeilingBytes)));
  auto construct_prefix = construct_source.read_entry("blob_0.bin");
  ASSERT_TRUE(static_cast<bool>(construct_prefix));
  EXPECT_EQ(construct_prefix.value().size(), kCumulativeChunkBytes);

  ZipReader move_constructed(std::move(construct_source));
  for (int i = 1; i < 4; ++i) {
    const std::string name = "blob_" + std::to_string(i) + ".bin";
    auto ok = move_constructed.read_entry(name);
    ASSERT_TRUE(static_cast<bool>(ok)) << "move-construction read failed at i=" << i << ": " << ok.error().message;
    EXPECT_EQ(ok.value().size(), kCumulativeChunkBytes);
  }
  auto construct_over = move_constructed.read_entry("one.bin");
  ASSERT_FALSE(static_cast<bool>(construct_over));
  EXPECT_EQ(construct_over.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_EQ(construct_over.error().context, CumulativeOverflowContext("one.bin", kCumulativeCeilingBytes, 1U));

  ZipReader move_assigned;
  move_assigned = std::move(move_constructed);
  auto assignment_over = move_assigned.read_entry("one.bin");
  ASSERT_FALSE(static_cast<bool>(assignment_over));
  EXPECT_EQ(assignment_over.error().code, FormulonErrorCode::kIoFileTooLarge);
  EXPECT_EQ(assignment_over.error().context, CumulativeOverflowContext("one.bin", kCumulativeCeilingBytes, 1U));
}

TEST(ZipReader, PolicyRejectionsDoNotConsumeCumulativeBudget) {
  // Keep the policy-rejected entries tiny and rewrite only their central
  // directory metadata. Their policy checks must return before miniz reaches
  // local-header validation, while the four legitimate entries still prove
  // the exact 256 MiB cumulative boundary in each independent session.
  std::vector<std::uint8_t> marker{0x11u};
  std::vector<std::uint8_t> payload(kCumulativeChunkBytes, 0u);
  for (std::size_t i = 0; i < payload.size(); ++i) {
    payload[i] = static_cast<std::uint8_t>(i & 0xFFu);
  }
  std::vector<std::pair<std::string, const std::vector<std::uint8_t>*>> entries{
      {"encrypted.bin", &marker}, {"huge.bin", &marker}, {"ratio.bin", &marker}};
  for (int i = 0; i < 4; ++i) {
    entries.emplace_back("blob_" + std::to_string(i) + ".bin", &payload);
  }
  entries.emplace_back("one.bin", &marker);
  std::vector<std::uint8_t> bytes = BuildEntriesZip(entries, static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  ASSERT_TRUE(MarkZipEntryEncrypted(bytes, "encrypted.bin"));
  ASSERT_TRUE(RewriteCentralDirectorySizes(bytes, "huge.bin", /*compressed_size=*/1U,
                                           static_cast<std::uint32_t>(kMaxExtractedBytesPerEntry + 1U)));
  ASSERT_TRUE(RewriteCentralDirectorySizes(bytes, "ratio.bin", /*compressed_size=*/1U,
                                           /*uncompressed_size=*/4U * 1024U * 1024U));

  const auto assert_boundary_after_rejection = [&](const std::string& rejected_name, FormulonErrorCode expected_code,
                                                   std::string_view expected_limit) {
    ZipReader zip;
    ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes), kCumulativeCeilingBytes)));
    auto rejected = zip.read_entry(rejected_name);
    ASSERT_FALSE(static_cast<bool>(rejected));
    EXPECT_EQ(rejected.error().code, expected_code);
    EXPECT_NE(rejected.error().context.find(expected_limit), std::string::npos);

    for (int i = 0; i < 4; ++i) {
      const std::string name = "blob_" + std::to_string(i) + ".bin";
      auto ok = zip.read_entry(name);
      ASSERT_TRUE(static_cast<bool>(ok)) << "read after " << rejected_name << " failed at i=" << i << ": "
                                         << ok.error().message;
      EXPECT_EQ(ok.value().size(), kCumulativeChunkBytes);
    }
    auto over = zip.read_entry("one.bin");
    ASSERT_FALSE(static_cast<bool>(over));
    EXPECT_EQ(over.error().code, FormulonErrorCode::kIoFileTooLarge);
    EXPECT_EQ(over.error().context, CumulativeOverflowContext("one.bin", kCumulativeCeilingBytes, 1U));
  };

  assert_boundary_after_rejection("encrypted.bin", FormulonErrorCode::kIoZipEncrypted, "entry=encrypted.bin");
  assert_boundary_after_rejection("huge.bin", FormulonErrorCode::kIoZipBomb, "limit=entry");
  assert_boundary_after_rejection("ratio.bin", FormulonErrorCode::kIoZipBomb, "limit=ratio");
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
