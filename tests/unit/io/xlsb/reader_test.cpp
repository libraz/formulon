//
// Unit tests for the MS-XLSB read pipeline skeleton. The tests
// hand-craft a minimal `.xlsb` package in memory (no committed binary
// fixtures) and assert the reader extracts sheet names + cell values
// correctly. A full corpus arrives in a later bundle.

#include "io/xlsb/reader.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "miniz.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return io::ByteSpan{v.data(), v.size()};
}

// ---------------------------------------------------------------------------
// Tiny XLSB record-stream builders. Mirrors the framing in
// `io/xlsb/record.{h,cpp}` (1-/2-byte type, 1..4-byte size, payload) so
// the tests can construct synthetic parts without dragging in the
// writer (which lands in Bundle 4.2).
// ---------------------------------------------------------------------------

void AppendVarInt(std::vector<std::uint8_t>& out, std::uint32_t v, std::size_t max_bytes) {
  for (std::size_t i = 0; i < max_bytes; ++i) {
    const std::uint8_t byte = static_cast<std::uint8_t>(v & 0x7F);
    v >>= 7;
    if (v == 0) {
      out.push_back(byte);
      return;
    }
    out.push_back(static_cast<std::uint8_t>(byte | 0x80));
  }
  // If we're here the caller asked for too few bytes; fall through.
}

void AppendRecord(std::vector<std::uint8_t>& out, std::uint16_t type, const std::vector<std::uint8_t>& payload) {
  AppendVarInt(out, type, /*max_bytes=*/2);
  AppendVarInt(out, static_cast<std::uint32_t>(payload.size()), /*max_bytes=*/4);
  out.insert(out.end(), payload.begin(), payload.end());
}

void AppendU8(std::vector<std::uint8_t>& out, std::uint8_t v) {
  out.push_back(v);
}

void AppendU32(std::vector<std::uint8_t>& out, std::uint32_t v) {
  out.push_back(static_cast<std::uint8_t>(v & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 8) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 16) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 24) & 0xFFU));
}

void AppendDouble(std::vector<std::uint8_t>& out, double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(v));
  AppendU32(out, static_cast<std::uint32_t>(bits & 0xFFFFFFFFU));
  AppendU32(out, static_cast<std::uint32_t>((bits >> 32) & 0xFFFFFFFFU));
}

/// Encodes an ASCII string as XLWideString (cch + UCS-2 little-endian).
void AppendXLWideString(std::vector<std::uint8_t>& out, std::string_view s) {
  AppendU32(out, static_cast<std::uint32_t>(s.size()));
  for (char c : s) {
    out.push_back(static_cast<std::uint8_t>(c));
    out.push_back(0);
  }
}

/// XLNullableWideString. Empty / null both encode as the 0xFFFFFFFF sentinel.
void AppendXLNullableWideString(std::vector<std::uint8_t>& out, std::string_view s) {
  if (s.empty()) {
    AppendU32(out, 0xFFFFFFFFU);
    return;
  }
  AppendXLWideString(out, s);
}

// ---------------------------------------------------------------------------
// Synthetic .xlsb part builders.
// ---------------------------------------------------------------------------

std::string ContentTypesXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/workbook.bin\" "
      "ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.bin\" "
      "ContentType=\"application/vnd.ms-excel.binIndexWs\"/>"
      "<Override PartName=\"/xl/worksheets/sheet2.bin\" "
      "ContentType=\"application/vnd.ms-excel.binIndexWs\"/>"
      "<Override PartName=\"/xl/sharedStrings.bin\" "
      "ContentType=\"application/vnd.ms-excel.sharedStrings\"/>"
      "</Types>");
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
      "<Relationship Id=\"rIdSheet2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet2.bin\"/>"
      "<Relationship Id=\"rIdSst\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" "
      "Target=\"sharedStrings.bin\"/>"
      "</Relationships>");
}

/// Builds one `BrtName` record body carrying `name` at workbook scope
/// with `rgce` as its formula. Only the prefix the reader consumes is
/// emitted (flags + 3 reserved + itab + cch + UTF-16LE name + cce +
/// rgce); the trailing fields a real Excel record carries after the
/// formula body are never read, so omitting them keeps the builder
/// honest about what the decoder depends on.
std::vector<std::uint8_t> NameRecord(std::string_view name, const std::vector<std::uint8_t>& rgce) {
  std::vector<std::uint8_t> p;
  p.push_back(0);  // flags[0] — fHidden clear.
  p.push_back(0);  // flags[1]
  p.push_back(0);  // reserved
  p.push_back(0);
  p.push_back(0);
  AppendU32(p, 0xFFFFFFFFU);  // itab == -1: workbook scope.
  AppendU32(p, static_cast<std::uint32_t>(name.size()));
  for (char c : name) {
    p.push_back(static_cast<std::uint8_t>(c));
    p.push_back(0);
  }
  AppendU32(p, static_cast<std::uint32_t>(rgce.size()));
  p.insert(p.end(), rgce.begin(), rgce.end());
  return p;
}

/// Builds `xl/workbook.bin` containing two BrtBundleSh entries
/// pointing at rIdSheet1 / rIdSheet2. `name_records` are emitted as
/// `BrtName` records after the sheet bundle, which is where Excel puts
/// them and where both name passes expect to find them.
std::vector<std::uint8_t> WorkbookBin(const std::vector<std::vector<std::uint8_t>>& name_records = {}) {
  std::vector<std::uint8_t> body;

  // BrtBeginBook (131): empty payload.
  AppendRecord(body, 131, {});

  // BrtBeginBundleShs (143).
  AppendRecord(body, 143, {});

  // BrtBundleSh (156) #1 = "Alpha"
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);  // hsState (visible)
    AppendU32(p, 1);  // iTabID
    AppendXLNullableWideString(p, "rIdSheet1");
    AppendXLWideString(p, "Alpha");
    AppendRecord(body, 156, p);
  }
  // BrtBundleSh (156) #2 = "Beta"
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendU32(p, 2);
    AppendXLNullableWideString(p, "rIdSheet2");
    AppendXLWideString(p, "Beta");
    AppendRecord(body, 156, p);
  }

  AppendRecord(body, 144, {});  // BrtEndBundleShs

  for (const std::vector<std::uint8_t>& rec : name_records) {
    AppendRecord(body, 39, rec);  // BrtName
  }

  AppendRecord(body, 132, {});  // BrtEndBook
  return body;
}

/// Builds a sheet with a single BrtCellReal at (`row`, `col`) with value
/// `cell_value`. `row` / `col` default to A1 but are overridable so the
/// bounds-validation tests can inject out-of-range indices.
std::vector<std::uint8_t> SheetBinReal(double cell_value, std::uint32_t row = 0, std::uint32_t col = 0) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData

  // BrtRowHdr (0): just the row index, plus 18 bytes of reserved
  // metadata that the skeleton skips. We supply only the first u32 —
  // the reader reads `current_row` from byte 0..3 and stops there
  // because the rest of the BrtRowHdr payload is not consumed.
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, row);  // row index
    AppendRecord(body, 0, p);
  }

  // BrtCellReal (5): cell-header (col, style3, ph1) + 8-byte double.
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, col);  // column
    AppendU8(p, 0);     // style[0]
    AppendU8(p, 0);     // style[1]
    AppendU8(p, 0);     // style[2]
    AppendU8(p, 0);     // fPhShow
    AppendDouble(p, cell_value);
    AppendRecord(body, 5, p);
  }

  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

/// Builds a sheet with a single BrtArrFmla whose RfX rect is
/// (`rw_first`..`rw_last`, `col_first`..`col_last`) and whose
/// CellParsedFormula rgce is `rgce` (cb = 0).
std::vector<std::uint8_t> SheetBinArrFmla(std::uint32_t rw_first, std::uint32_t rw_last, std::uint32_t col_first,
                                          std::uint32_t col_last, const std::vector<std::uint8_t>& rgce) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData

  {
    std::vector<std::uint8_t> p;
    AppendU32(p, rw_first);
    AppendRecord(body, 0, p);  // BrtRowHdr
  }

  // BrtArrFmla (426): RfX (4 x u32) + reserved byte + cce (u32) + rgce
  // + cb (u32).
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, rw_first);
    AppendU32(p, rw_last);
    AppendU32(p, col_first);
    AppendU32(p, col_last);
    AppendU8(p, 0);  // reserved/flag
    AppendU32(p, static_cast<std::uint32_t>(rgce.size()));
    p.insert(p.end(), rgce.begin(), rgce.end());
    AppendU32(p, 0);  // cb
    AppendRecord(body, 426, p);
  }

  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

/// Builds a sheet with a single BrtCellIsst at row 0, col 0 referencing
/// SST entry `sst_index`.
std::vector<std::uint8_t> SheetBinIsst(std::uint32_t sst_index) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});
  AppendRecord(body, 145, {});

  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendRecord(body, 0, p);  // BrtRowHdr
  }

  // BrtCellIsst (7): cell-header + u32 sst index.
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendU8(p, 0);
    AppendU8(p, 0);
    AppendU8(p, 0);
    AppendU8(p, 0);
    AppendU32(p, sst_index);
    AppendRecord(body, 7, p);
  }

  AppendRecord(body, 146, {});
  AppendRecord(body, 130, {});
  return body;
}

/// Builds a sheet with a single BrtFmlaNum at row 0, col 0: cached
/// numeric result `cached`, and a `CellParsedFormula` whose rgce is
/// `rgce`.
std::vector<std::uint8_t> SheetBinFmlaNum(double cached, const std::vector<std::uint8_t>& rgce) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData

  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendRecord(body, 0, p);  // BrtRowHdr
  }

  // BrtFmlaNum (9): cell-header (8) + double (8) + grbitFlags (u16) +
  // cce (u32) + rgce + cb (u32).
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);  // column
    AppendU8(p, 0);   // style[0]
    AppendU8(p, 0);   // style[1]
    AppendU8(p, 0);   // style[2]
    AppendU8(p, 0);   // fPhShow
    AppendDouble(p, cached);
    p.push_back(0);  // grbitFlags lo
    p.push_back(0);  // grbitFlags hi
    AppendU32(p, static_cast<std::uint32_t>(rgce.size()));
    p.insert(p.end(), rgce.begin(), rgce.end());
    AppendU32(p, 0);  // cb
    AppendRecord(body, 9, p);
  }

  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

/// Builds `xl/sharedStrings.bin` with one BrtSSTItem entry (the
/// payload is a single BrtSSTItem record carrying a 1-byte flags
/// prefix + XLWideString).
std::vector<std::uint8_t> SharedStringsBin(std::string_view item) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 159, {});  // BrtBeginSst
  // BrtSSTItem (19): u8 flags + XLWideString.
  {
    std::vector<std::uint8_t> p;
    AppendU8(p, 0);
    AppendXLWideString(p, item);
    AppendRecord(body, 19, p);
  }
  AppendRecord(body, 160, {});  // BrtEndSst
  return body;
}

// ---------------------------------------------------------------------------
// ZIP packaging via miniz.
// ---------------------------------------------------------------------------

struct PartFile {
  std::string path;
  std::vector<std::uint8_t> body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const PartFile& p : parts) {
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, p.path.c_str(), p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "miniz add failed for " << p.path;
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

std::vector<std::uint8_t> StringToBytes(std::string_view s) {
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

// ---------------------------------------------------------------------------
// Tests.
// ---------------------------------------------------------------------------

TEST(XlsbReader, ReadsTwoSheetsWithRealAndIsstCells) {
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

  const Workbook& wb = result.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
  EXPECT_EQ(wb.sheet(1).name(), "Beta");

  // Sheet 1: numeric cell.
  const Cell* c0 = wb.sheet(0).cell_at(0, 0);
  ASSERT_NE(c0, nullptr);
  ASSERT_TRUE(c0->cached_value.is_number());
  EXPECT_EQ(c0->cached_value.as_number(), 123.5);

  // Sheet 2: text cell resolved against the SST.
  const Cell* c1 = wb.sheet(1).cell_at(0, 0);
  ASSERT_NE(c1, nullptr);
  ASSERT_TRUE(c1->cached_value.is_text());
  EXPECT_EQ(c1->cached_value.as_text(), "hello world");

  // Audit counter: two literal cells decoded.
  EXPECT_EQ(result.value().cells_read, 2U);
}

TEST(XlsbReader, DecodesFormulaAndPreservesCachedValue) {
  // rgce = `PtgInt 5` (`=5`): 0x1E 0x05 0x00.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinFmlaNum(5.0, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  // The real formula text is decoded from the Ptg stream.
  EXPECT_EQ(c->formula_text, "=5");
  // The cached value is preserved.
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 5.0);
}

TEST(XlsbReader, UndecodableFormulaPreservesCachedValueWithoutFakeFormula) {
  // rgce = `PtgName` (0x23) + a 4-byte name index. PtgName is not in the
  // supported common-token set, so the decode fails and the reader must
  // preserve the cached value and store NO formula (never a fake one
  // that would recalc to #NAME?).
  const std::vector<std::uint8_t> rgce = {0x23, 0x01, 0x00, 0x00, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinFmlaNum(42.0, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  // No fabricated formula.
  EXPECT_TRUE(c->formula_text.empty());
  // Cached value preserved so the cell still shows the right number.
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 42.0);
  EXPECT_EQ(result.value().undecoded_formula_count, 1U);
  EXPECT_EQ(result.value().undecoded_defined_name_count, 0U);
}

TEST(XlsbReader, ExternalBookNameInCellFormulaKeepsCachedValueAndCounts) {
  // rgce = `PtgNameX` (0x39) + ixti (u16) + name index (u32): a defined
  // name owned by another workbook. Resolving it needs the supporting-
  // book and external-name tables, which this reader does not decode, so
  // the token stays explicitly unsupported. Reusing the index against
  // this workbook's own name table would silently retarget the formula
  // at a different name, so the cell must instead fall back to the same
  // contract an undecodable stream gets: cached value kept, no formula,
  // one diagnostic counted.
  const std::vector<std::uint8_t> rgce = {0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinFmlaNum(7.25, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  EXPECT_TRUE(c->formula_text.empty());
  ASSERT_TRUE(c->cached_value.is_number());
  EXPECT_EQ(c->cached_value.as_number(), 7.25);
  EXPECT_EQ(result.value().undecoded_formula_count, 1U);
  EXPECT_EQ(result.value().undecoded_defined_name_count, 0U);
}

TEST(XlsbReader, DecodableDefinedNameRegistersAndCountsNothing) {
  // Positive control for the record builder the next test relies on: a
  // `BrtName` whose rgce is `PtgInt 5` must reach the workbook intact,
  // so a zero counter there means "nothing was undecodable" rather than
  // "no name record was ever parsed".
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin({NameRecord("Rate", rgce)})});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const std::vector<io::DefinedName>& names = result.value().workbook.defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].name, "Rate");
  EXPECT_EQ(names[0].formula, "5");
  EXPECT_EQ(names[0].local_sheet_id, -1);
  EXPECT_EQ(result.value().undecoded_defined_name_count, 0U);
}

TEST(XlsbReader, ExternalBookNameInDefinedNameIsSkippedAndCounted) {
  // Same unsupported token, this time as a defined name's own body. The
  // name is dropped rather than registered with a fabricated formula,
  // and the read still succeeds so one external reference cannot cost
  // the caller the whole workbook.
  const std::vector<std::uint8_t> rgce = {0x39, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00};
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin({NameRecord("ExtRate", rgce)})});
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  EXPECT_TRUE(result.value().workbook.defined_names().empty());
  EXPECT_EQ(result.value().undecoded_defined_name_count, 1U);
  EXPECT_EQ(result.value().undecoded_formula_count, 0U);
}

TEST(XlsbReader, OutOfRangeCellColumnIsRecordCorrupt) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  // Column index == kMaxCols is one past the last valid column (XFD).
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0, /*row=*/0, /*col=*/Sheet::kMaxCols)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbReader, OutOfRangeRowHeaderIsRecordCorrupt) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  // Row index == kMaxRows is one past the last valid row.
  parts.push_back({"xl/worksheets/sheet1.bin", SheetBinReal(1.0, /*row=*/Sheet::kMaxRows, /*col=*/0)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbReader, ReversedArrayFormulaRectIsRecordCorrupt) {
  // rw_last > rw_first but col_last < col_first: reversed on the column
  // axis only. The anchor guard is an OR over the two axes, so without
  // rect validation this anchor would be recorded and the spill size
  // math would underflow-wrap.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};  // PtgInt 5
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin",
                   SheetBinArrFmla(/*rw_first=*/0, /*rw_last=*/1, /*col_first=*/2, /*col_last=*/1, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoXlsbRecordCorrupt);
}

TEST(XlsbReader, InBoundsArrayFormulaRectStillDecodes) {
  // A well-ordered in-bounds 2x1 rect anchored at A1 must keep decoding
  // after the bounds validation: the anchor gets the decoded formula
  // text and the read succeeds.
  const std::vector<std::uint8_t> rgce = {0x1E, 0x05, 0x00};  // PtgInt 5
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});
  parts.push_back({"xl/worksheets/sheet1.bin",
                   SheetBinArrFmla(/*rw_first=*/0, /*rw_last=*/1, /*col_first=*/0, /*col_last=*/0, rgce)});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinReal(1.0)});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message << " | " << result.error().context;

  const Cell* c = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(c, nullptr);
  EXPECT_EQ(c->formula_text, "=5");
}

TEST(XlsbReader, MissingContentTypesIsContentTypeInvalid) {
  std::vector<PartFile> parts;
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin()});

  const std::vector<std::uint8_t> archive = BuildZip(parts);
  auto result = read_xlsb(SpanOf(archive));
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
