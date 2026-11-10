// Retention of worksheet-tail records that have no model representation:
// grammar-slot classification and the fallback that keeps, regenerates or
// counts each unmodelled record.

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/xlsb/cf_records.h"
#include "io/xlsb/dv_records.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/sheet_reader_internal.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// True for a record the sheet writer produces from the model, so keeping
/// the source copy would emit it twice. Every entry is a record
/// `sheet_writer.cpp`'s `emit_sheet` writes; adding one here is an
/// assertion that the writer covers it, and is the only way to make a
/// record vanish without the load saying so.
bool IsWriterRegeneratedSheetRecord(XlsbRecordType type) {
  switch (type) {
    case XlsbRecordType::BrtBeginSheet:
    case XlsbRecordType::BrtEndSheet:
    case XlsbRecordType::BrtWsDim:
    case XlsbRecordType::BrtBeginWsViews:
    case XlsbRecordType::BrtEndWsView:
    case XlsbRecordType::BrtEndWsViews:
    case XlsbRecordType::BrtBeginColInfos:
    case XlsbRecordType::BrtEndColInfos:
    case XlsbRecordType::BrtBeginSheetData:
    case XlsbRecordType::BrtEndSheetData:
    case XlsbRecordType::BrtBeginMergeCells:
    case XlsbRecordType::BrtEndMergeCells:
    // Re-derived for every spilled dynamic array from the sheet's spill
    // regions, keyed to the workbook's own XLDAPR metadata entry. The
    // entry travels with the workbook, so a package that carried these
    // records carries the metadata part that makes them meaningful. The
    // value-metadata record likewise names a rich value of the regenerated part.
    case XlsbRecordType::BrtCellMeta:
    case XlsbRecordType::BrtValueMeta:
      return true;
    default:
      return false;
  }
}

/// Tail records in this slot occur after the model-owned merge block and
/// before the model-owned hyperlink block, even when the source stream omits
/// BrtBeginMergeCells/BrtEndMergeCells entirely.  The worksheet grammar uses
/// these ids for phonetic metadata, conditional formatting and data
/// validations; relying only on `merges_seen` would put them before newly
/// added model hyperlinks on a no-merge sheet.
bool IsAfterMergesBeforeHyperlinks(XlsbRecordType type) {
  switch (static_cast<std::uint16_t>(type)) {
    case kBrtPhoneticInfo:
    case kBrtBeginConditionalFormatting:
    case kBrtEndConditionalFormatting:
    case kBrtBeginCFRule:
    case kBrtEndCFRule:
    case kBrtBeginIconSet:
    case kBrtEndIconSet:
    case kBrtBeginDataBar:
    case kBrtEndDataBar:
    case kBrtBeginColorScale:
    case kBrtEndColorScale:
    case kBrtCFVO:
    case kBrtColor:
    case kBrtBeginDVals:
    case kBrtDVal:
    case kBrtEndDVals:
      return true;
    default:
      return false;
  }
}
}  // namespace

namespace detail {

/// Appends the framed bytes of one worksheet-tail record to the grammar slot
/// represented by `XlsbSheetTail`. `framed` spans the record header *and*
/// payload, so re-emission is a plain byte copy rather than a re-encode.
/// Slots only move forward: a record classified into a later slot advances
/// the phase, so an unlisted record after it (an x14 FRT block after
/// BrtMargins on a sheet with no merges or hyperlinks) is not hoisted ahead.
void RetainTailRecord(SheetDecodeState& state, XlsbRecordType type, const std::uint8_t* framed, std::size_t size) {
  std::vector<std::uint8_t>* dst = &state.tail.before_merges;
  if (state.hyperlinks_seen || is_after_hyperlinks(static_cast<std::uint16_t>(type))) {
    dst = &state.tail.after_hyperlinks;
    state.hyperlinks_seen = true;
  } else if (state.merges_seen || IsAfterMergesBeforeHyperlinks(type)) {
    dst = &state.tail.after_merges_before_hyperlinks;
    state.merges_seen = true;
  }
  dst->insert(dst->end(), framed, framed + size);
}

/// Decides what happens to a record the dispatch did not decode. This is
/// the whole of the fallback: retention where the writer has a slot for
/// the bytes, silence only where the writer produces the record itself,
/// and a counted warning otherwise.
///
/// Verbatim retention is available for the worksheet tail alone, because
/// `emit_sheet` re-derives the entire prefix (properties, dimensions,
/// views, column layout) and offers no position to splice foreign bytes
/// into. A record before `BrtEndSheetData` therefore has to be decoded to
/// survive; one that is not lands in the counted bucket rather than
/// disappearing.
RecordDisposition ResolveUnmodelledRecord(SheetDecodeState& state, XlsbRecordType type, const std::uint8_t* framed,
                                          std::size_t size) {
  if (IsWriterRegeneratedSheetRecord(type)) {
    return RecordDisposition::kRegenerated;
  }
  if (state.in_tail) {
    RetainTailRecord(state, type, framed, size);
    return RecordDisposition::kRetained;
  }
  StructuredLog("xlsb.record.dropped")
      .field("record_type", static_cast<std::int64_t>(type))
      .field("bytes", static_cast<std::int64_t>(size))
      .warn();
  return RecordDisposition::kAccounted;
}
}  // namespace detail

/// Tail records in this slot are emitted after the model-owned hyperlink
/// block.  This explicit grammar classification is needed when the source
/// has no BrtHLink at all but the in-memory caller adds one before writing.
bool is_after_hyperlinks(std::uint16_t type) {
  switch (type) {
    case kBrtPrintOptions:
    case kBrtMargins:
    case kBrtBeginHeaderFooter:
    case kBrtEndHeaderFooter:
    case kBrtPageSetup:
    case kBrtBeginRwBrk:
    case kBrtBrk:
    case kBrtEndRwBrk:
    case kBrtBeginColBrk:
    case kBrtEndColBrk:
    case kBrtBigName:
    case kBrtBeginCellWatches:
    case kBrtCellWatch:
    case kBrtEndCellWatches:
    case kBrtBeginCellIgnoreECs:
    case kBrtCellIgnoreEC:
    case kBrtEndCellIgnoreECs:
    case kBrtBeginSmartTags:
    case kBrtBeginCellSmartTags:
    case kBrtBeginCellSmartTag:
    case kBrtCellSmartTagProperty:
    case kBrtEndCellSmartTag:
    case kBrtEndCellSmartTags:
    case kBrtEndSmartTags:
    case kBrtDrawing:
    case kBrtLegacyDrawing:
    case kBrtLegacyDrawingHF:
    case kBrtBkhim:
    case kBrtBeginOleObjects:
    case kBrtOleObject:
    case kBrtEndOleObjects:
    case kBrtBeginActiveXControls:
    case kBrtActiveX:
    case kBrtEndActiveXControls:
    case kBrtBeginWebPubItems:
    case kBrtBeginWebPubItem:
    case kBrtEndWebPubItem:
    case kBrtEndWebPubItems:
    case kBrtBeginTableParts:
    case kBrtTablePart:
    case kBrtEndTableParts:
      return true;
    default:
      return false;
  }
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
