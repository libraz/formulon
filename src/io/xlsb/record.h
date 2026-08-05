//
// MS-XLSB record framing primitives.
//
// XLSB binary parts are sequences of records, each laid out as:
//
//   * record-type     — 1 or 2 bytes. The MSB of each byte signals
//                       continuation. Up to two bytes total, encoding a
//                       14-bit type id (per [MS-XLSB] §2.1.4).
//   * payload-size    — 1 to 4 bytes, MSB-continuation as well, encoding
//                       a 28-bit unsigned size.
//   * payload         — `size` raw bytes of body, type-specific.
//
// The Reader walks the byte stream record-by-record, dispatches on
// `type`, and forwards the `payload` slice to a per-record decoder.
// This module keeps the parsing primitives ([MS-XLSB] §2.5 dispatches —
// "varint type", "varint size", "RkNumber", "XLWideString",
// "XLNullableWideString") in one place so the per-record decoders don't
// each re-derive them.
//
// Design references:
//   * [MS-XLSB] §2.1 (Record framing) and §2.5.121 (RkNumber)

#ifndef FORMULON_IO_XLSB_RECORD_H_
#define FORMULON_IO_XLSB_RECORD_H_

#include <cstdint>
#include <string>

#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// One framed XLSB record.
///
///   * `type`     — decoded record-type id (`BrtCellReal`, etc.). The
///                  enum `XlsbRecordType` lists every record the
///                  reader dispatches on; unknown ids are surfaced as
///                  the raw integer.
///   * `payload`  — non-owning view into the record body. Lifetime is
///                  the same as the underlying `ByteSpan` passed to
///                  `read_record`; the reader will make a copy if it
///                  needs to outlive the parse loop.
struct XlsbRecord {
  std::uint16_t type;
  ByteSpan payload;
};

/// The subset of record types the reader consumes. Numeric ids come
/// from [MS-XLSB] §2.4.x; record types not in this enum still flow
/// through `read_record` (they're simply not dispatched on).
///
/// Keeping the enum minimal is intentional: the rest of the binary
/// surface lands as it becomes Reader-relevant. Anything unrecognised
/// is a no-op for the dispatch loop today; the per-part decoder
/// signals "skip" by reading the payload length and stepping over.
enum class XlsbRecordType : std::uint16_t {
  // Cell records.
  BrtRowHdr = 0,       ///< Row metadata; precedes a run of cell records.
  BrtCellBlank = 1,    ///< Blank cell (still carries column index).
  BrtCellRk = 2,       ///< RK-encoded numeric cell.
  BrtCellError = 3,    ///< Error-literal cell.
  BrtCellBool = 4,     ///< Boolean-literal cell.
  BrtCellReal = 5,     ///< IEEE 754 double-literal cell.
  BrtCellSt = 6,       ///< Inline-string cell.
  BrtCellIsst = 7,     ///< Cell referencing the SST.
  BrtFmlaString = 8,   ///< Formula cell with text result.
  BrtFmlaNum = 9,      ///< Formula cell with numeric result.
  BrtFmlaBool = 10,    ///< Formula cell with boolean result.
  BrtFmlaError = 11,   ///< Formula cell with error result.
  BrtSSTItem = 19,     ///< One entry of the shared-string table.
  BrtColInfo = 60,     ///< Column width / visibility / outline span.
  BrtMergeCell = 176,  ///< One merged-cell rectangle.

  // Container records (begin/end markers).
  BrtBeginSheet = 129,       ///< Start of a worksheet stream.
  BrtEndSheet = 130,         ///< End of a worksheet stream.
  BrtBeginBook = 131,        ///< Start of the workbook stream.
  BrtEndBook = 132,          ///< End of the workbook stream.
  BrtWbProp = 153,           ///< Workbook properties, including the date system.
  BrtBeginBundleShs = 143,   ///< Start of the sheet-bundle list.
  BrtEndBundleShs = 144,     ///< End of the sheet-bundle list.
  BrtBeginSheetData = 145,   ///< Start of `<sheetData>` equivalent.
  BrtEndSheetData = 146,     ///< End of `<sheetData>` equivalent.
  BrtBundleSh = 156,         ///< One entry in the sheet-bundle list.
  BrtBeginSst = 159,         ///< Start of the shared-string table.
  BrtEndSst = 160,           ///< End of the shared-string table.
  BrtBeginMergeCells = 177,  ///< Start of the merged-cell collection.
  BrtEndMergeCells = 178,    ///< End of the merged-cell collection.
  BrtBeginColInfos = 390,    ///< Start of the column-layout record collection.
  BrtEndColInfos = 391,      ///< End of the column-layout record collection.

  // Defined-name / future-function-name table (workbook globals).
  BrtName = 39,  ///< One `<definedName>`-equivalent, incl. hidden
                 ///< `_xlfn.*` / `_xlpm.*` future-function and LET/LAMBDA
                 ///< parameter name entries referenced by `PtgName`.

  // ExternSheet table (workbook globals): resolves a `PtgRef3d` /
  // `PtgArea3d` `ixti` to a `(itabFirst, itabLast)` sheet-index range.
  // `itabFirst == itabLast` is a single-sheet qualified reference;
  // `itabFirst != itabLast` is a genuine 3-D range (e.g. `Sheet1:Sheet3`).
  BrtExternSheet = 362,

  // Array-formula body: the real Ptg tokens for a CSE / dynamic-array
  // formula. The formula-result "shell" record at the array's anchor
  // cell carries a `PtgExp` placeholder instead of the real tokens; this
  // record supplies them, keyed by the anchor's `(row, col)` via its
  // leading `RfX` (row/col range).
  BrtArrFmla = 426,

  // Styles (`xl/styles.bin`): custom number formats and the font / fill
  // / border / xf tables `iStyleRef` (in the per-cell header) indexes
  // into.
  BrtFmt = 44,                 ///< Custom `<numFmt>` (id + format string).
  BrtFont = 43,                ///< One `<font>` entry.
  BrtFill = 45,                ///< One `<fill>` entry.
  BrtBorder = 46,              ///< One `<border>` entry.
  BrtXF = 47,                  ///< One `<xf>` entry (cellXfs or cellStyleXfs).
  BrtBeginCellStyleXFs = 626,  ///< Start of the `<cellStyleXfs>` table.
  BrtEndCellStyleXFs = 627,    ///< End of the `<cellStyleXfs>` table.
  BrtBeginCellXFs = 617,       ///< Start of the `<cellXfs>` table.
  BrtEndCellXFs = 618,         ///< End of the `<cellXfs>` table.
};

/// Reads one framed record from `cursor`. On success `cursor` advances
/// past the record (header + payload); on failure `cursor` is left in
/// an unspecified state.
///
/// Errors:
///   * `kIoXlsbRecordTruncated` — the record header or payload would
///                                 overrun the underlying buffer.
Expected<XlsbRecord, Error> read_record(ByteSpan& cursor);

/// Reads a raw 8-bit unsigned integer; advances `cursor`. Returns
/// `kIoXlsbRecordTruncated` when the buffer is empty.
Expected<std::uint8_t, Error> read_u8(ByteSpan& cursor);

/// Reads a little-endian 16-bit unsigned integer; advances `cursor`.
Expected<std::uint16_t, Error> read_u16(ByteSpan& cursor);

/// Reads a little-endian 32-bit unsigned integer; advances `cursor`.
Expected<std::uint32_t, Error> read_u32(ByteSpan& cursor);

/// Reads an `XLWideString` ([MS-XLSB] §2.5.166): a 32-bit length-prefix
/// followed by `length` UCS-2 little-endian code units. The returned
/// string is UTF-8 encoded. Lone surrogates are passed through as if
/// they were proper BMP code points (best-effort decode); this matches
/// what Excel emits in practice on round-trips of malformed inputs.
Expected<std::string, Error> read_xlwidestring(ByteSpan& cursor);

/// Reads an `XLNullableWideString` ([MS-XLSB] §2.5.167): identical to
/// `XLWideString` except the special length value `0xFFFFFFFF` denotes
/// a null string and decodes to the empty string.
Expected<std::string, Error> read_xlnullablewidestring(ByteSpan& cursor);

/// Decodes an MS-XLSB `RkNumber` ([MS-XLSB] §2.5.121).
///
/// The 32-bit value carries:
///   * bit 0       — `fX100`: divide by 100 if set.
///   * bit 1       — `fInt`: the upper 30 bits represent a signed
///                   integer when set, otherwise the upper bits of an
///                   IEEE 754 double (lower 34 bits zeroed).
///   * bits 2..31  — the 30-bit payload (signed when `fInt` is set).
///
/// The function does *not* advance any cursor — callers pass in the
/// raw 32-bit RK value already extracted via `read_u32`.
double decode_rk_number(std::uint32_t rk);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_RECORD_H_
