//
// MS-XLSB worksheet-part decoding: `xl/worksheets/sheet*.bin` into the
// `Sheet` model. Every record is decoded, retained verbatim in the sheet's
// `XlsbSheetTail`, left to the writer to re-derive, or counted as dropped.

#ifndef FORMULON_IO_XLSB_SHEET_READER_H_
#define FORMULON_IO_XLSB_SHEET_READER_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <tuple>
#include <vector>

#include "io/array_anchor_budget.h"
#include "io/xlsb/ptg_reader.h"
#include "phonetic.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace io {
namespace xlsb {

/// Per-sheet decode state. The reader walks records in order and
/// updates `current_row` whenever it sees a `BrtRowHdr`. Cell records
/// then resolve `(current_row, col)` to the absolute cell address.
struct SheetDecodeState {
  std::uint32_t current_row = 0;
  bool row_seen = false;
  std::uint32_t cells_decoded = 0;
  /// Dynamic-array anchors recorded while walking `BrtArrFmla` records.
  /// Registered as spill regions in a second pass after the whole sheet
  /// has been decoded (see `RegisterArraySpills`'s doc comment for why
  /// this cannot happen inline).
  std::vector<ArrayAnchor> array_anchors;
  /// Worksheet-tail records retained verbatim (see `XlsbSheetTail`).
  XlsbSheetTail tail;
  /// True once `BrtEndSheetData` has been seen: everything from there to
  /// `BrtEndSheet` is tail.
  bool in_tail = false;
  /// True once the merged-cell block, or any record of a later slot, has
  /// been passed. This is the fallback grammar phase for tail records not
  /// covered by an explicit slot id.
  bool merges_seen = false;
  /// True once the first raw BrtHLink, or any record of the post-hyperlink
  /// slot, has been encountered. Raw hyperlink records are model-owned and
  /// are never retained, but this marker keeps unrelated records after them
  /// in the correct post-hyperlink buffer.
  bool hyperlinks_seen = false;
  /// Source records this sheet could not carry whole (see
  /// `RecordDisposition::kAccounted`).
  std::uint32_t dropped_records = 0;
  /// Cell-metadata index of the `BrtCellMeta` just read, which applies to
  /// the next cell record; 0 when none is pending.
  std::uint32_t pending_cell_meta = 0;
  /// `(row, col, ifmd)` of every formula cell a `BrtCellMeta` preceded. The
  /// caller keeps as dynamic-array formulas those whose `ifmd` is the
  /// workbook's dynamic-array entry (`apply_loaded_dynamic_array_marks`).
  std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>> cell_metadata;
};

/// Decodes one sheet binary part. Cells (literal + formula) flow into
/// `wb.sheet(sheet_index)`. SST indices are resolved against
/// `sst_entries`; out-of-range indices are returned as
/// `kIoXlsbCorrupt`.
Expected<SheetDecodeState, Error> DecodeSheetBin(
    const std::vector<std::uint8_t>& body, std::size_t sheet_index, Workbook& wb,
    const std::vector<std::string_view>& sst_entries, const std::vector<std::vector<PhoneticRun>>& sst_phonetic,
    const std::vector<PhoneticProperties>& sst_phonetic_props, std::deque<std::string>& text_storage,
    const std::vector<std::string>& sheet_names, const std::vector<XlsbName>& name_table,
    const std::vector<XlsbSheetRange>& sheet_ranges, const XlsbExternalBooks& external_books,
    std::uint32_t* undecoded_formula_count);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_SHEET_READER_H_
