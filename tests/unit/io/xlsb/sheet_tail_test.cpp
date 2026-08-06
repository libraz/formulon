//
// Round-trip tests for the `.xlsb` worksheet tail: the records a sheet part
// carries after its cell table (auto-filter, conditional formatting, data
// validation, hyperlinks) and the sheet rels those records address.
//
// A sheet part is consumed whole by the reader, so package-level passthrough
// cannot rescue what the model does not express. These tests build a package
// with a hand-written tail, run it through `read_xlsb` -> `write_xlsb`, and
// assert the tail comes back — in the same order, exactly once, and with its
// relationship targets still reachable.

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// Record ids used by the tail ([MS-XLSB] §2.4.x). None of them are in the
// reader's dispatch set, which is the point: they must survive without the
// engine understanding them.
constexpr std::uint16_t kBrtBeginAFilter = 161;
constexpr std::uint16_t kBrtEndAFilter = 162;
constexpr std::uint16_t kBrtBeginCondFmt = 463;
constexpr std::uint16_t kBrtEndCondFmt = 464;
constexpr std::uint16_t kBrtBeginCFRule = 465;
constexpr std::uint16_t kBrtEndCFRule = 466;
constexpr std::uint16_t kBrtDVal = 64;
constexpr std::uint16_t kBrtBeginDVals = 573;
constexpr std::uint16_t kBrtEndDVals = 574;
constexpr std::uint16_t kBrtHLink = 494;

constexpr std::uint16_t kBrtBeginSheet = 129;
constexpr std::uint16_t kBrtEndSheet = 130;
constexpr std::uint16_t kBrtBeginSheetData = 145;
constexpr std::uint16_t kBrtEndSheetData = 146;
constexpr std::uint16_t kBrtRowHdr = 0;
constexpr std::uint16_t kBrtCellReal = 5;
constexpr std::uint16_t kBrtMergeCell = 176;
constexpr std::uint16_t kBrtBeginMergeCells = 177;
constexpr std::uint16_t kBrtEndMergeCells = 178;

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

// ---------------------------------------------------------------------------
// Record framing (mirrors `io/xlsb/record.cpp`).
// ---------------------------------------------------------------------------

void AppendVarInt(std::vector<std::uint8_t>& out, std::uint32_t value, std::size_t max_bytes) {
  for (std::size_t i = 0; i < max_bytes; ++i) {
    const auto byte = static_cast<std::uint8_t>(value & 0x7FU);
    value >>= 7U;
    if (value == 0U) {
      out.push_back(byte);
      return;
    }
    out.push_back(static_cast<std::uint8_t>(byte | 0x80U));
  }
}

void AppendRecord(std::vector<std::uint8_t>& out, std::uint16_t type, const std::vector<std::uint8_t>& payload) {
  AppendVarInt(out, type, /*max_bytes=*/2);
  AppendVarInt(out, static_cast<std::uint32_t>(payload.size()), /*max_bytes=*/4);
  out.insert(out.end(), payload.begin(), payload.end());
}

void AppendU8(std::vector<std::uint8_t>& out, std::uint8_t value) {
  out.push_back(value);
}

void AppendU32(std::vector<std::uint8_t>& out, std::uint32_t value) {
  out.push_back(static_cast<std::uint8_t>(value & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((value >> 8U) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((value >> 16U) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((value >> 24U) & 0xFFU));
}

void AppendDouble(std::vector<std::uint8_t>& out, double value) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &value, sizeof(value));
  AppendU32(out, static_cast<std::uint32_t>(bits & 0xFFFFFFFFU));
  AppendU32(out, static_cast<std::uint32_t>((bits >> 32U) & 0xFFFFFFFFU));
}

void AppendXLWideString(std::vector<std::uint8_t>& out, std::string_view text) {
  AppendU32(out, static_cast<std::uint32_t>(text.size()));
  for (const char ch : text) {
    out.push_back(static_cast<std::uint8_t>(ch));
    out.push_back(0);
  }
}

void AppendXLNullableWideString(std::vector<std::uint8_t>& out, std::string_view text) {
  if (text.empty()) {
    AppendU32(out, 0xFFFFFFFFU);
    return;
  }
  AppendXLWideString(out, text);
}

/// Payload bytes that let a test tell one opaque record from another.
std::vector<std::uint8_t> Marker(std::uint32_t tag) {
  std::vector<std::uint8_t> payload;
  AppendU32(payload, tag);
  return payload;
}

// ---------------------------------------------------------------------------
// Package parts.
// ---------------------------------------------------------------------------

std::string ContentTypesXml(bool with_drawing_part) {
  std::string out(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/workbook.bin\" "
      "ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.bin\" ContentType=\"application/vnd.ms-excel.binIndexWs\"/>");
  if (with_drawing_part) {
    out.append(
        "<Override PartName=\"/xl/drawings/drawing1.xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>");
  }
  out.append("</Types>");
  return out;
}

std::string PackageRelsXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.bin\"/>"
      "</Relationships>");
}

std::string WorkbookRelsXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rIdSheet1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.bin\"/>"
      "</Relationships>");
}

std::vector<std::uint8_t> WorkbookBin() {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 131, {});  // BrtBeginBook
  AppendRecord(body, 143, {});  // BrtBeginBundleShs
  {
    std::vector<std::uint8_t> payload;
    AppendU32(payload, 0);  // visible
    AppendU32(payload, 1);  // iTabID
    AppendXLNullableWideString(payload, "rIdSheet1");
    AppendXLWideString(payload, "Alpha");
    AppendRecord(body, 156, payload);  // BrtBundleSh
  }
  AppendRecord(body, 144, {});  // BrtEndBundleShs
  AppendRecord(body, 132, {});  // BrtEndBook
  return body;
}

/// Builds `xl/worksheets/sheet1.bin`: one numeric cell, then a tail holding
/// an auto-filter block, a merged-cell block, a conditional-formatting block,
/// a data-validation block and a hyperlink — in the order Excel emits them.
std::vector<std::uint8_t> SheetBinWithTail(bool with_merges) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, kBrtBeginSheet, {});
  AppendRecord(body, kBrtBeginSheetData, {});
  {
    std::vector<std::uint8_t> payload;
    AppendU32(payload, 0);
    AppendRecord(body, kBrtRowHdr, payload);
  }
  {
    std::vector<std::uint8_t> payload;
    AppendU32(payload, 0);  // column
    AppendU8(payload, 0);
    AppendU8(payload, 0);
    AppendU8(payload, 0);
    AppendU8(payload, 0);
    AppendDouble(payload, 42.0);
    AppendRecord(body, kBrtCellReal, payload);
  }
  AppendRecord(body, kBrtEndSheetData, {});

  // AUTOFILTER precedes MERGECELLS in the worksheet grammar.
  AppendRecord(body, kBrtBeginAFilter, Marker(0xA1A1A1A1U));
  AppendRecord(body, kBrtEndAFilter, {});

  if (with_merges) {
    std::vector<std::uint8_t> count;
    AppendU32(count, 1U);
    AppendRecord(body, kBrtBeginMergeCells, count);
    std::vector<std::uint8_t> rect;
    AppendU32(rect, 0U);  // first row
    AppendU32(rect, 1U);  // last row
    AppendU32(rect, 2U);  // first col
    AppendU32(rect, 3U);  // last col
    AppendRecord(body, kBrtMergeCell, rect);
    AppendRecord(body, kBrtEndMergeCells, {});
  }

  AppendRecord(body, kBrtBeginCondFmt, Marker(0xC0C0C0C0U));
  AppendRecord(body, kBrtBeginCFRule, Marker(0xC1C1C1C1U));
  AppendRecord(body, kBrtEndCFRule, {});
  AppendRecord(body, kBrtEndCondFmt, {});

  AppendRecord(body, kBrtBeginDVals, Marker(0xD0D0D0D0U));
  AppendRecord(body, kBrtDVal, Marker(0xD1D1D1D1U));
  AppendRecord(body, kBrtEndDVals, {});

  {
    std::vector<std::uint8_t> payload;
    AppendU32(payload, 0U);  // first row
    AppendU32(payload, 0U);  // last row
    AppendU32(payload, 0U);  // first col
    AppendU32(payload, 0U);  // last col
    AppendXLNullableWideString(payload, "rIdHyperlink");
    AppendRecord(body, kBrtHLink, payload);
  }

  AppendRecord(body, kBrtEndSheet, {});
  return body;
}

std::string SheetRelsXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rIdHyperlink\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" "
      "Target=\"https://example.invalid/a\" TargetMode=\"External\"/>"
      "<Relationship Id=\"rIdDrawing\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
      "Target=\"../drawings/drawing1.xml\"/>"
      "<Relationship Id=\"rIdGhost\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
      "Target=\"../drawings/absent.xml\"/>"
      "</Relationships>");
}

// ---------------------------------------------------------------------------
// ZIP packaging + record scanning.
// ---------------------------------------------------------------------------

struct PartFile {
  std::string path;
  std::vector<std::uint8_t> body;
};

std::vector<std::uint8_t> BytesOf(std::string_view text) {
  return std::vector<std::uint8_t>(text.begin(), text.end());
}

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const PartFile& part : parts) {
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, part.path.c_str(), part.body.data(), part.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "miniz add failed for " << part.path;
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

std::vector<PartFile> SourceParts(bool with_merges, bool with_sheet_rels, bool with_drawing_part) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", BytesOf(ContentTypesXml(with_drawing_part))});
  parts.push_back({"_rels/.rels", BytesOf(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", BytesOf(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinWithTail(with_merges)});
  if (with_sheet_rels) {
    parts.push_back({"xl/worksheets/_rels/sheet1.bin.rels", BytesOf(SheetRelsXml())});
  }
  if (with_drawing_part) {
    parts.push_back({"xl/drawings/drawing1.xml", BytesOf("<xdr:wsDr/>")});
  }
  return parts;
}

/// Every record type in `body`, in stream order.
std::vector<std::uint16_t> RecordTypes(const std::vector<std::uint8_t>& body) {
  std::vector<std::uint16_t> types;
  ByteSpan cursor = SpanOf(body);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      break;
    }
    types.push_back(record_or.value().type);
  }
  return types;
}

std::size_t IndexOfType(const std::vector<std::uint16_t>& types, std::uint16_t wanted) {
  for (std::size_t i = 0; i < types.size(); ++i) {
    if (types[i] == wanted) {
      return i;
    }
  }
  return types.size();
}

std::size_t CountOfType(const std::vector<std::uint16_t>& types, std::uint16_t wanted) {
  std::size_t count = 0;
  for (const std::uint16_t type : types) {
    if (type == wanted) {
      ++count;
    }
  }
  return count;
}

/// Reads `path` out of a written package. Fails the calling test when the
/// package does not open or the part is missing.
std::vector<std::uint8_t> PartOf(const std::vector<std::uint8_t>& archive, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(archive))));
  if (!zip.has_entry(path)) {
    ADD_FAILURE() << "missing part: " << path;
    return {};
  }
  auto bytes_or = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes_or));
  return bytes_or ? bytes_or.value() : std::vector<std::uint8_t>{};
}

/// Reads a package and writes it straight back out.
std::vector<std::uint8_t> ReadThenWrite(const std::vector<std::uint8_t>& archive) {
  auto read_or = read_xlsb(SpanOf(archive));
  EXPECT_TRUE(static_cast<bool>(read_or)) << (read_or ? "" : read_or.error().message);
  if (!read_or) {
    return {};
  }
  auto write_or = write_xlsb(read_or.value().workbook);
  EXPECT_TRUE(static_cast<bool>(write_or)) << (write_or ? "" : write_or.error().message);
  return write_or ? write_or.value() : std::vector<std::uint8_t>{};
}

// ---------------------------------------------------------------------------
// Tests.
// ---------------------------------------------------------------------------

TEST(XlsbSheetTail, ReaderRetainsRecordsAfterTheCellTable) {
  const std::vector<std::uint8_t> archive = BuildZip(SourceParts(
      /*with_merges=*/true, /*with_sheet_rels=*/true, /*with_drawing_part=*/true));
  auto read_or = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  const XlsbSheetTail& tail = read_or.value().workbook.sheet(0).xlsb_tail();
  EXPECT_FALSE(tail.empty());

  // The merged-cell block splits the tail: the auto-filter is captured before
  // it, everything else after. The merge records themselves are not retained
  // — the writer re-emits them from `Sheet::merges()`.
  const std::vector<std::uint16_t> before = RecordTypes(tail.before_merges);
  EXPECT_EQ(before, (std::vector<std::uint16_t>{kBrtBeginAFilter, kBrtEndAFilter}));

  const std::vector<std::uint16_t> after = RecordTypes(tail.after_merges);
  EXPECT_EQ(after, (std::vector<std::uint16_t>{kBrtBeginCondFmt, kBrtBeginCFRule, kBrtEndCFRule, kBrtEndCondFmt,
                                               kBrtBeginDVals, kBrtDVal, kBrtEndDVals, kBrtHLink}));
}

TEST(XlsbSheetTail, RoundTripKeepsEverySheetLevelFeature) {
  const std::vector<std::uint8_t> archive = BuildZip(SourceParts(
      /*with_merges=*/true, /*with_sheet_rels=*/true, /*with_drawing_part=*/true));
  const std::vector<std::uint8_t> rewritten = ReadThenWrite(archive);
  ASSERT_FALSE(rewritten.empty());

  const std::vector<std::uint16_t> types = RecordTypes(PartOf(rewritten, "xl/worksheets/sheet1.bin"));
  for (const std::uint16_t wanted : {kBrtBeginAFilter, kBrtEndAFilter, kBrtBeginCondFmt, kBrtBeginCFRule, kBrtEndCFRule,
                                     kBrtEndCondFmt, kBrtBeginDVals, kBrtDVal, kBrtEndDVals, kBrtHLink}) {
    EXPECT_EQ(CountOfType(types, wanted), 1U) << "record type " << wanted;
  }

  // Order relative to the model-owned merged-cell block is preserved.
  EXPECT_LT(IndexOfType(types, kBrtEndSheetData), IndexOfType(types, kBrtBeginAFilter));
  EXPECT_LT(IndexOfType(types, kBrtBeginAFilter), IndexOfType(types, kBrtBeginMergeCells));
  EXPECT_LT(IndexOfType(types, kBrtEndMergeCells), IndexOfType(types, kBrtBeginCondFmt));
  EXPECT_LT(IndexOfType(types, kBrtBeginCondFmt), IndexOfType(types, kBrtBeginDVals));
  EXPECT_LT(IndexOfType(types, kBrtBeginDVals), IndexOfType(types, kBrtHLink));
  EXPECT_LT(IndexOfType(types, kBrtHLink), IndexOfType(types, kBrtEndSheet));
}

TEST(XlsbSheetTail, MergedCellsComeFromTheModelAndAreNotDuplicated) {
  const std::vector<std::uint8_t> archive = BuildZip(SourceParts(
      /*with_merges=*/true, /*with_sheet_rels=*/false, /*with_drawing_part=*/false));
  const std::vector<std::uint8_t> rewritten = ReadThenWrite(archive);
  ASSERT_FALSE(rewritten.empty());

  const std::vector<std::uint16_t> types = RecordTypes(PartOf(rewritten, "xl/worksheets/sheet1.bin"));
  EXPECT_EQ(CountOfType(types, kBrtBeginMergeCells), 1U);
  EXPECT_EQ(CountOfType(types, kBrtMergeCell), 1U);
  EXPECT_EQ(CountOfType(types, kBrtEndMergeCells), 1U);
}

TEST(XlsbSheetTail, TailSurvivesACellEdit) {
  const std::vector<std::uint8_t> archive = BuildZip(SourceParts(
      /*with_merges=*/false, /*with_sheet_rels=*/false, /*with_drawing_part=*/false));
  auto read_or = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  Workbook wb = std::move(read_or.value().workbook);
  wb.sheet(0).set_cell_value(0U, 0U, Value::number(99.0));

  auto write_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message;
  const std::vector<std::uint16_t> types = RecordTypes(PartOf(write_or.value(), "xl/worksheets/sheet1.bin"));
  EXPECT_EQ(CountOfType(types, kBrtBeginCondFmt), 1U);
  EXPECT_EQ(CountOfType(types, kBrtDVal), 1U);
  EXPECT_EQ(CountOfType(types, kBrtHLink), 1U);

  // Without a merged-cell block in the source the whole tail lands in one
  // buffer, and still precedes the end of the sheet.
  EXPECT_LT(IndexOfType(types, kBrtBeginAFilter), IndexOfType(types, kBrtEndSheet));
}

TEST(XlsbSheetTail, SheetRelsKeepTheirOriginalIdsAndDropAbsentTargets) {
  const std::vector<std::uint8_t> archive = BuildZip(SourceParts(
      /*with_merges=*/true, /*with_sheet_rels=*/true, /*with_drawing_part=*/true));
  const std::vector<std::uint8_t> rewritten = ReadThenWrite(archive);
  ASSERT_FALSE(rewritten.empty());

  const std::vector<std::uint8_t> rels_bytes = PartOf(rewritten, "xl/worksheets/_rels/sheet1.bin.rels");
  const std::string rels(rels_bytes.begin(), rels_bytes.end());

  // The hyperlink id is what the retained BrtHLink record addresses, so it
  // must not be renumbered.
  EXPECT_NE(rels.find("Id=\"rIdHyperlink\""), std::string::npos) << rels;
  EXPECT_NE(rels.find("Target=\"https://example.invalid/a\""), std::string::npos) << rels;
  EXPECT_NE(rels.find("TargetMode=\"External\""), std::string::npos) << rels;

  // An internal target that travelled through passthrough stays connected...
  EXPECT_NE(rels.find("Id=\"rIdDrawing\""), std::string::npos) << rels;
  EXPECT_NE(rels.find("Target=\"../drawings/drawing1.xml\""), std::string::npos) << rels;

  // ...while one whose part is not in the package is dropped rather than
  // emitted as a dangling relationship.
  EXPECT_EQ(rels.find("rIdGhost"), std::string::npos) << rels;
}

TEST(XlsbSheetTail, SheetWithoutRelationshipsEmitsNoRelsPart) {
  const std::vector<std::uint8_t> archive = BuildZip(SourceParts(
      /*with_merges=*/false, /*with_sheet_rels=*/false, /*with_drawing_part=*/false));
  const std::vector<std::uint8_t> rewritten = ReadThenWrite(archive);
  ASSERT_FALSE(rewritten.empty());

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(rewritten))));
  EXPECT_FALSE(zip.has_entry("xl/worksheets/_rels/sheet1.bin.rels"));
}

TEST(XlsbSheetTail, ModelBuiltWorkbookCarriesNoTail) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));
  EXPECT_TRUE(wb.sheet(0).xlsb_tail().empty());

  auto write_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message;
  const std::vector<std::uint16_t> types = RecordTypes(PartOf(write_or.value(), "xl/worksheets/sheet1.bin"));
  EXPECT_EQ(CountOfType(types, kBrtBeginCondFmt), 0U);
  EXPECT_EQ(CountOfType(types, kBrtHLink), 0U);
  EXPECT_EQ(types.back(), kBrtEndSheet);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
