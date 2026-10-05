#pragma once

#include <array>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "miniz.h"

namespace formulon::io::xlsb::reader_test_support {

inline io::ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return io::ByteSpan{v.data(), v.size()};
}

// ---------------------------------------------------------------------------
// Tiny XLSB record-stream builders. Mirrors the framing in
// `io/xlsb/record.{h,cpp}` (1-/2-byte type, 1..4-byte size, payload) so
// the tests can construct synthetic parts without dragging in the
// writer (which lands in Bundle 4.2).
// ---------------------------------------------------------------------------

inline void AppendVarInt(std::vector<std::uint8_t>& out, std::uint32_t v, std::size_t max_bytes) {
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

inline void AppendRecord(std::vector<std::uint8_t>& out, std::uint16_t type, const std::vector<std::uint8_t>& payload) {
  AppendVarInt(out, type, /*max_bytes=*/2);
  AppendVarInt(out, static_cast<std::uint32_t>(payload.size()), /*max_bytes=*/4);
  out.insert(out.end(), payload.begin(), payload.end());
}

inline void AppendU8(std::vector<std::uint8_t>& out, std::uint8_t v) {
  out.push_back(v);
}

inline void AppendU32(std::vector<std::uint8_t>& out, std::uint32_t v) {
  out.push_back(static_cast<std::uint8_t>(v & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 8) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 16) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 24) & 0xFFU));
}

inline void AppendDouble(std::vector<std::uint8_t>& out, double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(v));
  AppendU32(out, static_cast<std::uint32_t>(bits & 0xFFFFFFFFU));
  AppendU32(out, static_cast<std::uint32_t>((bits >> 32) & 0xFFFFFFFFU));
}

/// Encodes an ASCII string as XLWideString (cch + UCS-2 little-endian).
inline void AppendXLWideString(std::vector<std::uint8_t>& out, std::string_view s) {
  AppendU32(out, static_cast<std::uint32_t>(s.size()));
  for (char c : s) {
    out.push_back(static_cast<std::uint8_t>(c));
    out.push_back(0);
  }
}

/// XLNullableWideString. Empty / null both encode as the 0xFFFFFFFF sentinel.
inline void AppendXLNullableWideString(std::vector<std::uint8_t>& out, std::string_view s) {
  if (s.empty()) {
    AppendU32(out, 0xFFFFFFFFU);
    return;
  }
  AppendXLWideString(out, s);
}

// ---------------------------------------------------------------------------
// Synthetic .xlsb part builders.
// ---------------------------------------------------------------------------

inline std::string ContentTypesXml() {
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

inline std::string PackageRelsXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.bin\"/>"
      "</Relationships>");
}

inline std::string WorkbookRelsXml() {
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
/// with `rgce` as its formula and `comment` as the trailing Name
/// Manager comment string (empty encodes the real-Excel null
/// `XLNullableWideString` sentinel). Emits exactly the byte shape a real
/// Excel-365-produced `BrtName` for a plain (non-placeholder) defined
/// name carries: flags + 3 reserved + itab + cch + UTF-16LE name + cce +
/// rgce + cb (always 0 here -- no array-constant `rgcb`) + the one
/// trailing comment string. Verified against `xlsb_fidelity_base.xlsb`'s
/// own "Rate" `BrtName` record byte-for-byte.
inline std::vector<std::uint8_t> NameRecord(std::string_view name, const std::vector<std::uint8_t>& rgce,
                                            std::string_view comment = {}) {
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
  AppendU32(p, 0U);  // cb: no rgcb.
  AppendXLNullableWideString(p, comment);
  return p;
}

/// Builds `xl/workbook.bin` containing two BrtBundleSh entries
/// pointing at rIdSheet1 / rIdSheet2. `name_records` are emitted as
/// `BrtName` records after the sheet bundle, which is where Excel puts
/// them and where both name passes expect to find them.
inline std::vector<std::uint8_t> WorkbookBin(const std::vector<std::vector<std::uint8_t>>& name_records = {}) {
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
inline std::vector<std::uint8_t> SheetBinReal(double cell_value, std::uint32_t row = 0, std::uint32_t col = 0) {
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

inline std::vector<std::uint8_t> WorksheetFormatPayload(std::uint32_t dx_g_col, std::uint16_t cch_def_col_width,
                                                        std::uint16_t miy_def_rw_height, std::uint32_t flags) {
  std::vector<std::uint8_t> payload;
  AppendU32(payload, dx_g_col);
  payload.push_back(static_cast<std::uint8_t>(cch_def_col_width & 0xFFU));
  payload.push_back(static_cast<std::uint8_t>((cch_def_col_width >> 8U) & 0xFFU));
  payload.push_back(static_cast<std::uint8_t>(miy_def_rw_height & 0xFFU));
  payload.push_back(static_cast<std::uint8_t>((miy_def_rw_height >> 8U) & 0xFFU));
  AppendU32(payload, flags);
  return payload;
}

inline std::vector<std::uint8_t> SheetBinWorksheetFormat(const std::vector<std::uint8_t>& format_payload) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});              // BrtBeginSheet
  AppendRecord(body, 485, format_payload);  // BrtWsFmtInfo
  AppendRecord(body, 145, {});              // BrtBeginSheetData
  AppendRecord(body, 146, {});              // BrtEndSheetData
  AppendRecord(body, 130, {});              // BrtEndSheet
  return body;
}

/// Builds a layout-only sheet carrying three BrtRowHdr records: explicit
/// style 0, explicit style 1, and ixfe 1 without fGhostDirty. The last one
/// is the negative control proving that ixfe is ignored unless the flag says
/// the row style is effective.
inline std::vector<std::uint8_t> SheetBinRowStyles() {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData
  const auto append_row = [&body](std::uint32_t row, std::uint32_t style_xf, std::uint8_t flags2) {
    std::vector<std::uint8_t> p;
    AppendU32(p, row);
    AppendU32(p, style_xf);
    p.push_back(0);
    p.push_back(0);  // miyRw
    AppendU8(p, 0);  // flags1
    AppendU8(p, flags2);
    AppendU8(p, 0);            // fPhShow
    AppendU32(p, 0);           // ccolspan
    AppendRecord(body, 0, p);  // BrtRowHdr
  };
  append_row(0U, 0U, 0x40U);
  append_row(1U, 1U, 0x40U);
  append_row(2U, 1U, 0U);
  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

/// Builds a sheet with a single BrtArrFmla whose RfX rect is
/// (`rw_first`..`rw_last`, `col_first`..`col_last`) and whose
/// CellParsedFormula rgce is `rgce` (cb = 0).
inline std::vector<std::uint8_t> SheetBinArrFmla(std::uint32_t rw_first, std::uint32_t rw_last, std::uint32_t col_first,
                                                 std::uint32_t col_last, const std::vector<std::uint8_t>& rgce) {
  const std::vector<std::array<std::uint32_t, 4>> rects = {{{rw_first, rw_last, col_first, col_last}}};
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData

  {
    std::vector<std::uint8_t> p;
    AppendU32(p, rw_first);
    AppendRecord(body, 0, p);  // BrtRowHdr
  }

  for (const std::array<std::uint32_t, 4>& rect : rects) {
    std::vector<std::uint8_t> p;
    AppendU32(p, rect[0]);
    AppendU32(p, rect[1]);
    AppendU32(p, rect[2]);
    AppendU32(p, rect[3]);
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

inline std::vector<std::uint8_t> SheetBinArrFmlas(const std::vector<std::array<std::uint32_t, 4>>& rects,
                                                  const std::vector<std::uint8_t>& rgce) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData

  {
    std::vector<std::uint8_t> p;
    AppendU32(p, rects.front()[0]);
    AppendRecord(body, 0, p);  // BrtRowHdr
  }

  for (const std::array<std::uint32_t, 4>& rect : rects) {
    // BrtArrFmla (426): RfX (4 x u32) + reserved byte + cce (u32) + rgce
    // + cb (u32).
    std::vector<std::uint8_t> p;
    AppendU32(p, rect[0]);
    AppendU32(p, rect[1]);
    AppendU32(p, rect[2]);
    AppendU32(p, rect[3]);
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
inline std::vector<std::uint8_t> SheetBinIsst(std::uint32_t sst_index) {
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
inline std::vector<std::uint8_t> SheetBinFmlaNum(double cached, const std::vector<std::uint8_t>& rgce) {
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
inline std::vector<std::uint8_t> SharedStringsBin(std::string_view item) {
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

/// Builds `xl/sharedStrings.bin` with one BrtSSTItem entry whose
/// phonetic tail (`kRichStrPhonetic` set, no rich-text runs) declares
/// one run but supplies zero bytes for it -- `DecodePhoneticTail`
/// (sst_reader.cpp) bounds the run count against the remaining payload
/// before reading, so this exercises the `kIoXlsbRecordTruncated` path
/// rather than an out-of-bounds read.
inline std::vector<std::uint8_t> TruncatedPhoneticSharedStringsBin() {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 159, {});  // BrtBeginSst
  {
    std::vector<std::uint8_t> p;
    AppendU8(p, 0x02);          // flags: fPhonetic set, fRichSt clear
    AppendXLWideString(p, "");  // the string itself
    AppendXLWideString(p, "");  // phonetic text
    AppendU32(p, 1U);           // run count -- no run bytes follow
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

inline std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
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

inline std::vector<std::uint8_t> StringToBytes(std::string_view s) {
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

// ---------------------------------------------------------------------------
// Tests.
// ---------------------------------------------------------------------------

}  // namespace formulon::io::xlsb::reader_test_support
