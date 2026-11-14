// The Compatibility Version stamp through xlsx and xlsb: read as stored, written as the model holds it. A loaded
// stamped book keeps its stamp (warnBelowVersion included), a loaded unstamped book stays unstamped, a new book is
// written with Excel's new-book stamp, and a version set through the API is written. In xlsx the stamp is the
// workbook <extLst> ext {D14903EA-33C4-47F7-8F05-3474C54BE107}; in xlsb it is record 5135 (u32 reserved,
// warnBelowVersion, setVersion) in its own FRT block before BrtEndBook, after any slicer-cache block.

#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

using Bytes = std::vector<std::uint8_t>;
using Stamp = std::optional<std::uint32_t>;

constexpr std::uint16_t kCompatRecord = 5135;
constexpr std::uint16_t kBeginSlicerCacheIds = 1070;
constexpr std::uint16_t kTableSlicerCacheIds = 2075;

Bytes ReadFixture(const std::string& rel) {
  std::ifstream in(std::string(FORMULON_FIXTURES_DIR) + "/" + rel, std::ios::binary);
  return Bytes(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

ByteSpan SpanOf(const Bytes& b) {
  return ByteSpan{b.data(), b.size()};
}

std::string Entry(const Bytes& package, const char* name) {
  ZipReader zip;
  if (!zip.open(SpanOf(package)) || !zip.has_entry(name)) {
    return {};
  }
  auto bytes = zip.read_entry(name);
  return bytes ? std::string(bytes.value().begin(), bytes.value().end()) : std::string();
}

std::size_t Count(const std::string& text, const std::string& needle) {
  std::size_t n = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + 1U)) {
    ++n;
  }
  return n;
}

Workbook LoadXlsx(const Bytes& bytes) {
  auto read = read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return std::move(read.value().workbook);
}

Workbook LoadXlsb(const Bytes& bytes) {
  auto read = xlsb::read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return std::move(read.value().workbook);
}

Bytes SaveXlsx(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? saved.value() : Bytes{};
}

Bytes SaveXlsb(const Workbook& wb) {
  auto saved = xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? saved.value() : Bytes{};
}

// The stamp ext of a saved workbook.xml, or "" when there is none.
std::string StampXml(const Bytes& xlsx) {
  const std::string wb = Entry(xlsx, "xl/workbook.xml");
  EXPECT_LE(Count(wb, "{D14903EA-33C4-47F7-8F05-3474C54BE107}"), 1U) << "more than one stamp ext";
  const std::size_t at = wb.find("version ");
  const std::size_t uri = wb.find("{D14903EA-33C4-47F7-8F05-3474C54BE107}");
  if (uri == std::string::npos || at == std::string::npos) {
    return "";
  }
  return wb.substr(at, wb.find("/>", at) - at);
}

// Record types of a saved workbook.bin and the 5135 payload words.
struct BookRecords {
  std::vector<std::uint16_t> types;
  std::vector<std::uint32_t> stamp;  // empty when there is no 5135
};

BookRecords ReadBook(const Bytes& xlsb) {
  const std::string text = Entry(xlsb, "xl/workbook.bin");
  const std::vector<std::uint8_t> bin(text.begin(), text.end());
  BookRecords out;
  for (const xlsb::FramedRecord& rec : xlsb::split_records(bin)) {
    out.types.push_back(rec.type);
    if (rec.type == kCompatRecord && rec.payload.size == 12U) {
      for (std::size_t i = 0; i < 3U; ++i) {
        out.stamp.push_back(xlsb::load_u32(rec.payload.data + 4U * i));
      }
    }
  }
  return out;
}

TEST(CompatibilityVersionIo, XlsxStampedBooksReadAndKeepTheirStamp) {
  const Workbook plain = LoadXlsx(ReadFixture("cellxfs_compaction/cx1_D1_dup_both_used.xlsx"));
  EXPECT_EQ(plain.compatibility_stamp(), Stamp(2U));
  EXPECT_EQ(plain.compatibility_warn_below(), Stamp());
  EXPECT_EQ(StampXml(SaveXlsx(plain)), "version setVersion=\"2\"");

  const Workbook warned = LoadXlsx(ReadFixture("excel/rich_errors/b5_fn.xlsx"));
  EXPECT_EQ(warned.compatibility_stamp(), Stamp(2U));
  EXPECT_EQ(warned.compatibility_warn_below(), Stamp(2U));
  EXPECT_EQ(StampXml(SaveXlsx(warned)), "version warnBelowVersion=\"2\" setVersion=\"2\"");
}

TEST(CompatibilityVersionIo, XlsxUnstampedBookStaysUnstamped) {
  const Workbook wb = LoadXlsx(ReadFixture("excel/blankslot_authored.xlsx"));
  EXPECT_EQ(wb.compatibility_stamp(), Stamp());
  EXPECT_EQ(wb.compatibility_version(), 1U);
  EXPECT_EQ(StampXml(SaveXlsx(wb)), "");
}

TEST(CompatibilityVersionIo, XlsxNewAndSetBooksAreStamped) {
  Workbook wb = Workbook::create();
  const Bytes fresh = SaveXlsx(wb);
  EXPECT_EQ(StampXml(fresh), "version warnBelowVersion=\"2\" setVersion=\"2\"");
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(1U)));
  const Bytes set = SaveXlsx(wb);
  EXPECT_EQ(StampXml(set), "version warnBelowVersion=\"2\" setVersion=\"1\"");
  const Workbook back = LoadXlsx(set);
  EXPECT_EQ(back.compatibility_version(), 1U);
}

TEST(CompatibilityVersionIo, XlsbStampedBooksReadAndKeepTheirStamp) {
  const Workbook plain = LoadXlsb(ReadFixture("cellxfs_compaction/cx1_D1_dup_both_used_resaved.xlsb"));
  EXPECT_EQ(plain.compatibility_stamp(), Stamp(2U));
  EXPECT_EQ(plain.compatibility_warn_below(), Stamp());
  const BookRecords saved = ReadBook(SaveXlsb(plain));
  EXPECT_EQ(saved.stamp, (std::vector<std::uint32_t>{0U, 0U, 2U}));
  ASSERT_GE(saved.types.size(), 4U);
  // FRTBegin, 5135, FRTEnd, BrtEndBook.
  EXPECT_EQ(saved.types[saved.types.size() - 3U], kCompatRecord);

  const Workbook warned = LoadXlsb(ReadFixture("excel/rich_errors_xlsb/b5_fn.xlsb"));
  EXPECT_EQ(warned.compatibility_warn_below(), Stamp(2U));
  EXPECT_EQ(ReadBook(SaveXlsb(warned)).stamp, (std::vector<std::uint32_t>{0U, 2U, 2U}));
}

TEST(CompatibilityVersionIo, XlsbUnstampedBookStaysUnstamped) {
  const Workbook wb = LoadXlsb(ReadFixture("cellxfs_compaction/mg2_B14_merge_items_b.xlsb"));
  EXPECT_EQ(wb.compatibility_stamp(), Stamp());
  EXPECT_EQ(wb.compatibility_version(), 1U);
  EXPECT_TRUE(ReadBook(SaveXlsb(wb)).stamp.empty());
}

TEST(CompatibilityVersionIo, XlsbNewAndSetBooksAreStamped) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(ReadBook(SaveXlsb(wb)).stamp, (std::vector<std::uint32_t>{0U, 2U, 2U}));
  ASSERT_TRUE(static_cast<bool>(wb.set_compatibility_version(1U)));
  const Bytes set = SaveXlsb(wb);
  EXPECT_EQ(ReadBook(set).stamp, (std::vector<std::uint32_t>{0U, 2U, 1U}));
  EXPECT_EQ(LoadXlsb(set).compatibility_version(), 1U);
}

// A stamp a later Excel writes (Excel lists 4 as unsupported and computes it as 2) is read and written back as is.
TEST(CompatibilityVersionIo, StampsAboveTheKnownVersionsAreKept) {
  for (const std::uint32_t stamp : {4U, 4294967295U}) {
    const Workbook wb = LoadXlsx(ReadFixture("excel/compat_version/stamp_" + std::to_string(stamp) + ".xlsx"));
    EXPECT_EQ(wb.compatibility_stamp(), Stamp(stamp));
    EXPECT_EQ(wb.compatibility_version(), stamp);
    EXPECT_EQ(wb.text_index_unit(), TextIndexUnit::kCodePoints);
    EXPECT_EQ(StampXml(SaveXlsx(wb)), "version setVersion=\"" + std::to_string(stamp) + "\"");
    const Bytes xlsb = SaveXlsb(wb);
    EXPECT_EQ(ReadBook(xlsb).stamp, (std::vector<std::uint32_t>{0U, 0U, stamp}));
    EXPECT_EQ(LoadXlsb(xlsb).compatibility_stamp(), Stamp(stamp));
  }
}

// A retained slicer-cache block is spliced in before the stamp block, in Excel's FRT order.
TEST(CompatibilityVersionIo, XlsbSlicerCacheBlockPrecedesTheStamp) {
  Workbook wb = LoadXlsb(ReadFixture("excel/table_slicer.xlsb"));
  wb.restore_compatibility_stamp(2U, 2U);
  const BookRecords saved = ReadBook(SaveXlsb(wb));
  std::size_t slicer = saved.types.size();
  std::size_t stamp = saved.types.size();
  for (std::size_t i = 0; i < saved.types.size(); ++i) {
    if ((saved.types[i] == kBeginSlicerCacheIds || saved.types[i] == kTableSlicerCacheIds) &&
        slicer == saved.types.size()) {
      slicer = i;
    }
    if (saved.types[i] == kCompatRecord) {
      stamp = i;
    }
  }
  ASSERT_LT(slicer, saved.types.size());
  ASSERT_LT(stamp, saved.types.size());
  EXPECT_LT(slicer, stamp);
}

}  // namespace
}  // namespace io
}  // namespace formulon
