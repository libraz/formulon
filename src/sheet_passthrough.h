//
// Worksheet content the calculation model does not interpret, kept verbatim
// so a save does not drop it: unmodelled `<worksheet>` children with their
// schema slot, and the retained record tail of an `.xlsb` sheet part.

#ifndef FORMULON_SHEET_PASSTHROUGH_H_
#define FORMULON_SHEET_PASSTHROUGH_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace formulon {

/// The ECMA-376 §18.3.1.99 `CT_Worksheet` child sequence, and the
/// position lookup a serializer needs to place an element in it.
///
/// The sequence is ordered, and a worksheet whose children appear out of
/// that order is invalid, so "where does this element go" has exactly one
/// answer and it is this table.
namespace worksheet_child {

/// Every `<worksheet>` child name, in schema order. Index into this array
/// is the slot number `WorksheetRawChild::slot` records.
inline constexpr std::string_view kOrder[] = {
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
};

/// Number of names in `kOrder`, and the value `slot_of` returns for a
/// name outside the sequence.
inline constexpr std::size_t kCount = sizeof(kOrder) / sizeof(kOrder[0]);

/// Slot of `name` in `kOrder`, or `kCount` when the schema has no such
/// child. A caller must decide what an unplaceable element means for it;
/// this function does not guess.
inline std::size_t slot_of(std::string_view name) noexcept {
  for (std::size_t i = 0; i < kCount; ++i) {
    if (kOrder[i] == name) {
      return i;
    }
  }
  return kCount;
}

}  // namespace worksheet_child

/// One `<worksheet>` child the calculation model does not interpret,
/// captured verbatim together with the position it must be written back
/// at.
///
/// `slot` indexes `worksheet_child::kOrder`. An element whose name is not
/// in that table takes the slot of the child that preceded it in the
/// source document, so it lands between the same two siblings on the way
/// out without the table having to be exhaustive.
struct WorksheetRawChild {
  std::uint32_t slot = 0;
  std::string xml;
};

/// Every `<worksheet>` child the calculation model does not fold in,
/// ordered by `slot`.
///
/// The list is a sweep of what the reader did not consume, not an
/// allowlist of anticipated names: an element nobody has modelled yet is
/// preserved by default rather than dropped, which is what keeps a save
/// from silently damaging an Excel-authored sheet.
using WorksheetRawExtensions = std::vector<WorksheetRawChild>;

/// One entry of a source `.xlsb` workbook's `BrtExternSheet` table, by
/// sheet name so it survives sheet reordering.
struct XlsbExternSheetEntry {
  /// Sheet names the entry spans; both empty for the sheetless entry a
  /// book-scope `PtgNameX` resolves through.
  std::string first;
  std::string last;
  /// True for an entry naming no sheet of this workbook: another
  /// workbook's sheets, or a sheet index outside the source workbook.
  bool unresolved = false;
  /// For another workbook's entry, the `ExternalLinkRecord::index` of its
  /// link, with `first` / `last` naming that book's sheets. 0 for this
  /// workbook, and for an entry whose book or sheets could not be bound.
  std::uint32_t external_book = 0;
};

/// Worksheet records captured verbatim from an `.xlsb` sheet part that the
/// calculation model does not express: conditional formatting, data
/// validation, hyperlinks, auto-filter, print setup, manual breaks, and the
/// drawing / table part references.
///
/// The XML path keeps such content as unknown *parts* or as raw element
/// strings (`WorksheetRawExtensions`), but a binary sheet part is a single
/// record stream the reader consumes whole, so neither mechanism reaches it.
/// Retaining the framed record bytes is the binary equivalent.
///
/// The three buffers represent the grammar slots before merges, after merges
/// and before hyperlinks, and after hyperlinks. The merge and hyperlink
/// blocks are the only tail constructs the model owns and re-emits itself;
/// keeping these slots preserves source order even when a source omits one of
/// those blocks.
///
/// Retention is byte-verbatim. A workbook row/column edit remaps the x14 CF
/// ranges and formulas and each sparkline's range and source; coordinates
/// in any other retained record keep their pre-edit values. A retained formula
/// blob's sheet-qualified references are `ixti` indices into the source
/// `BrtExternSheet` table, which `extern_sheets` records so the writer can
/// keep those indices.
struct XlsbSheetTail {
  /// Records between the end of the cell table and the merged-cell block.
  std::vector<std::uint8_t> before_merges;
  /// Records after the merged-cell block and before the model-owned
  /// `BrtHLink` block. Raw `BrtHLink` records are decoded into
  /// `Sheet::hyperlinks()` and are never retained here.
  std::vector<std::uint8_t> after_merges_before_hyperlinks;
  /// Records after the model-owned `BrtHLink` block, up to `BrtEndSheet`.
  std::vector<std::uint8_t> after_hyperlinks;
  /// The source workbook's `BrtExternSheet` table, in `ixti` order.
  std::vector<XlsbExternSheetEntry> extern_sheets;

  bool empty() const noexcept {
    return before_merges.empty() && after_merges_before_hyperlinks.empty() && after_hyperlinks.empty();
  }
};

}  // namespace formulon

#endif  // FORMULON_SHEET_PASSTHROUGH_H_
