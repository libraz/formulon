#include "io/xlsb/sheet_reader.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <deque>
#include <iterator>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/array_anchor_budget.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/dv_records.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "phonetic.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/structured_log.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// True when `ptgs` is the placeholder Excel stores in a cell whose real
/// tokens live in a separate record: a lone `PtgExp` (`0x01`) naming the
/// owning anchor. Both the anchor and the phantoms of an array or
/// dynamic-array formula carry it, and `BrtArrFmla` supplies the tokens
/// out-of-band, so the shell is the expected encoding rather than a Ptg
/// the decoder failed on.
///
/// `PtgExp` is only ever the whole stream — it takes no operands and
/// cannot combine with another token — so testing the first byte
/// classifies the record.
bool IsAnchorShellFormula(ByteSpan ptgs) {
  return ptgs.size != 0U && ptgs.data[0] == 0x01U;
}

/// Attempts to decode the `rgce` Ptg byte stream into an Excel formula
/// text (with a leading `=`). Returns the formula text on success, or an
/// empty string when the stream uses a token outside the supported set
/// (a structured-log diagnostic records the reason). On the empty-string
/// path the caller PRESERVES the cell's cached value and stores no
/// formula, so the cell still displays the correct cached result instead
/// of a fabricated formula that would recalc to `#NAME?`.
///
/// An anchor shell returns empty as well, but neither logs nor counts:
/// `undecoded_formula_count` states that the reader could not recover a
/// formula the source carried, and for a shell there is no formula at
/// that position to recover. The record that does carry the tokens is
/// decoded on its own. A shell whose owning record the reader does not
/// model instead surfaces through that record's own disposition, so the
/// loss is still reported — just not as a decode failure here.
std::string DecodeFormulaText(ByteSpan ptg_bytes, ByteSpan rgcb, const std::vector<std::string>& sheet_names,
                              const std::vector<XlsbName>& name_table, const std::vector<XlsbSheetRange>& sheet_ranges,
                              const XlsbExternalBooks& external_books, std::size_t sheet_index, std::uint32_t row,
                              std::uint32_t col, std::uint32_t* undecoded_formula_count) {
  if (IsAnchorShellFormula(ptg_bytes)) {
    return {};
  }
  Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
  auto ast_or = decode_ptgs(ptg_bytes, rgcb, arena, sheet_names, name_table, sheet_ranges, external_books,
                            static_cast<std::int32_t>(sheet_index));
  if (!ast_or) {
    StructuredLog("xlsb.formula.not_decoded")
        .field("sheet_index", static_cast<std::int64_t>(sheet_index))
        .field("row", static_cast<std::int64_t>(row))
        .field("col", static_cast<std::int64_t>(col))
        .field("ptg_bytes", static_cast<std::int64_t>(ptg_bytes.size))
        .field("reason", ast_or.error().message)
        .warn();
    if (undecoded_formula_count != nullptr) {
      ++*undecoded_formula_count;
    }
    return {};
  }
  std::string out("=");
  out.append(parser::format_formula(*ast_or.value()));
  return out;
}

/// What became of one source record. The four values are exhaustive: a
/// record is decoded into the model, kept verbatim for re-emission,
/// left to the writer to re-derive, or reported as content this load
/// could not carry. There is no fifth outcome, and in particular no
/// outcome that discards a record without saying so.
///
/// `[[nodiscard]]` is deliberate. The defect this classification exists
/// to prevent is a record reaching the end of the dispatch with nothing
/// done to it, so the dispatch's answer must not be droppable either.
enum class [[nodiscard]] RecordDisposition {
  /// Decoded into the `Workbook` / `Sheet` model. Also covers a record
  /// whose content the reader took responsibility for and whose partial
  /// loss it reported through a more specific counter -- a formula whose
  /// Ptg stream fell outside the supported set is counted once, as an
  /// undecoded formula, not twice.
  kModelled,
  /// Kept verbatim in `XlsbSheetTail` and re-emitted byte-for-byte.
  kRetained,
  /// Not kept, because the writer emits an equivalent record of its own
  /// from the model. The membership test is
  /// `IsWriterRegeneratedSheetRecord`.
  kRegenerated,
  /// Neither decoded nor preserved: the content is reported through
  /// `XlsbReadResult::dropped_record_count` and a structured warning.
  /// A record decoded only in part resolves here too, so the residue is
  /// visible rather than implied.
  kAccounted,
};

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
    // records carries the metadata part that makes them meaningful.
    case XlsbRecordType::BrtCellMeta:
      return true;
    default:
      return false;
  }
}

constexpr std::uint32_t kMaxHyperlinkRelIdUnits = 32767U;
constexpr std::uint32_t kMaxHyperlinkLocationUnits = 2083U;
constexpr std::uint32_t kMaxHyperlinkTooltipUnits = 255U;
constexpr std::uint32_t kMaxHyperlinkDisplayUnits = 32767U;

Expected<std::string, Error> ReadHyperlinkWideString(ByteSpan& cursor, const char* field, std::uint32_t max_units) {
  if (cursor.size < 4U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                      std::string("xlsb BrtHLink ") + field + " length truncated", "context=xlsb_reader");
  }
  const std::uint32_t cch =
      static_cast<std::uint32_t>(cursor.data[0]) | (static_cast<std::uint32_t>(cursor.data[1]) << 8U) |
      (static_cast<std::uint32_t>(cursor.data[2]) << 16U) | (static_cast<std::uint32_t>(cursor.data[3]) << 24U);
  if (cch > max_units) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                      std::string("xlsb BrtHLink ") + field + " exceeds string limit",
                      "context=xlsb_reader cch=" + std::to_string(cch));
  }
  return read_xlwidestring(cursor);
}

/// The `BrtWsProp` fields `sheet_writer.cpp`'s `EmitWorksheetProperties`
/// writes for a sheet with no source record of its own. A source record
/// matching them carries nothing beyond what the writer will re-derive;
/// one that differs carries flags (dialog-sheet, fit-to-page, outline
/// direction, filter mode, sync anchors) the model has no field for.
constexpr std::uint16_t kDefaultWsPropFlags = 0x04C9U;
constexpr std::uint8_t kDefaultWsPropFlags2 = 0x02U;
constexpr std::uint32_t kWsPropSyncUnused = 0xFFFFFFFFU;

/// Appends `<tabColor .../>` for the decoded `BrtColor`, or nothing when
/// the colour is automatic -- Excel writes no `<tabColor>` for the
/// default tab, and an explicit `auto="1"` would make the two containers
/// disagree on an otherwise identical sheet.
void AppendTabColor(std::string& out, const ColorSpec& spec, std::uint32_t argb) {
  char buf[32];
  switch (spec.kind) {
    case ColorSpec::Kind::kRgb:
      std::snprintf(buf, sizeof(buf), "<tabColor rgb=\"%08X\"/>", argb);
      out.append(buf);
      break;
    case ColorSpec::Kind::kTheme:
      std::snprintf(buf, sizeof(buf), "<tabColor theme=\"%u\"", spec.theme);
      out.append(buf);
      if (spec.tint != 0.0) {
        // Shortest round-trip spelling, matching the styles writer: the
        // same tint must not be spelled two ways depending on whether the
        // sheet came in as XLSB or XLSX.
        out.append(" tint=\"");
        append_xml_number(out, spec.tint);
        out.push_back('"');
      }
      out.append("/>");
      break;
    case ColorSpec::Kind::kIndexed:
      std::snprintf(buf, sizeof(buf), "<tabColor indexed=\"%u\"/>", spec.indexed);
      out.append(buf);
      break;
    case ColorSpec::Kind::kAuto:
    case ColorSpec::Kind::kNone:
      break;
  }
}

/// Decodes `BrtWsProp` ([MS-XLSB] §2.4.858) into the sheet's raw
/// `<sheetPr>` fragment. Layout, verified against a real
/// Excel-365-produced `xl/worksheets/sheetN.bin` and symmetric with
/// `sheet_writer.cpp`: three flag bytes, an 8-byte `BrtColor` tab colour,
/// `rwSync` and `colSync`, then the VBA `CodeName` as an XLWideString.
///
/// The record reaches the model as XML rather than as typed fields
/// because `<sheetPr>` is what `SheetPrintSettings` stores and what the
/// introspection API hands back: routing the XLSB form through the same
/// string makes a `.xlsb`-loaded sheet answer exactly as its `.xlsx` twin
/// does. A sheet with neither a code name nor a tab colour produces no
/// fragment at all, which is also what the OOXML reader records for a
/// worksheet with no `<sheetPr>` element.
///
/// Returns `false` when the record carried flag or sync fields the model
/// has no room for, so the caller can report the residue.
Expected<bool, Error> DecodeWorksheetProperties(const XlsbRecord& rec, Sheet& sheet, std::size_t sheet_index) {
  ByteSpan p = rec.payload;
  auto flags_or = read_u16(p);
  auto flags2_or = read_u8(p);
  auto color_flags_or = read_u8(p);
  auto color_index_or = read_u8(p);
  auto color_tint_or = read_u16(p);
  auto red_or = read_u8(p);
  auto green_or = read_u8(p);
  auto blue_or = read_u8(p);
  auto alpha_or = read_u8(p);
  auto rw_sync_or = read_u32(p);
  auto col_sync_or = read_u32(p);
  if (!flags_or || !flags2_or || !color_flags_or || !color_index_or || !color_tint_or || !red_or || !green_or ||
      !blue_or || !alpha_or || !rw_sync_or || !col_sync_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtWsProp fields truncated",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index));
  }
  auto code_name_or = read_xlwidestring(p);
  if (!code_name_or) {
    return code_name_or.error();
  }

  const std::uint32_t argb =
      (static_cast<std::uint32_t>(alpha_or.value()) << 24U) | (static_cast<std::uint32_t>(red_or.value()) << 16U) |
      (static_cast<std::uint32_t>(green_or.value()) << 8U) | static_cast<std::uint32_t>(blue_or.value());
  ColorSpec tab;
  switch (static_cast<std::uint32_t>(color_flags_or.value() >> 1U)) {
    case 1U:
      tab.kind = ColorSpec::Kind::kIndexed;
      tab.indexed = color_index_or.value();
      break;
    case 2U:
      tab.kind = ColorSpec::Kind::kRgb;
      tab.rgb = argb;
      break;
    case 3U:
      tab.kind = ColorSpec::Kind::kTheme;
      tab.theme = color_index_or.value();
      tab.tint = static_cast<double>(static_cast<std::int16_t>(color_tint_or.value())) / 32767.0;
      break;
    default:
      tab.kind = ColorSpec::Kind::kAuto;
      break;
  }
  // Excel marks the default tab with the automatic palette slot rather
  // than the automatic colour type, so both spellings mean "no tab
  // colour" and neither produces a `<tabColor>` element.
  constexpr std::uint8_t kAutomaticPaletteIndex = 0x40U;
  if (tab.kind == ColorSpec::Kind::kIndexed && color_index_or.value() == kAutomaticPaletteIndex) {
    tab.kind = ColorSpec::Kind::kAuto;
  }

  std::string sheet_pr;
  const std::string& code_name = code_name_or.value();
  std::string body;
  AppendTabColor(body, tab, argb);
  if (!code_name.empty() || !body.empty()) {
    sheet_pr.append("<sheetPr");
    if (!code_name.empty()) {
      sheet_pr.append(" codeName=\"");
      AppendXmlAttrEscaped(sheet_pr, code_name);
      sheet_pr.push_back('"');
    }
    if (body.empty()) {
      sheet_pr.append("/>");
    } else {
      sheet_pr.push_back('>');
      sheet_pr.append(body);
      sheet_pr.append("</sheetPr>");
    }
  }
  if (!sheet_pr.empty()) {
    sheet.mutable_print_settings().sheet_pr_xml = std::move(sheet_pr);
  }

  return flags_or.value() == kDefaultWsPropFlags && flags2_or.value() == kDefaultWsPropFlags2 &&
         rw_sync_or.value() == kWsPropSyncUnused && col_sync_or.value() == kWsPropSyncUnused;
}

/// Decodes the fixed-width worksheet default-format record. The model only
/// carries the default column/row metrics and base column width; the thick
/// border and outline-level metadata has no corresponding model state, so it
/// is surfaced as a structured warning while the representable fields still
/// round-trip.
Expected<void, Error> DecodeWorksheetFormatInfo(const XlsbRecord& rec, Sheet& sheet, std::size_t sheet_index) {
  constexpr std::uint32_t kAbsentDefaultColumnWidth = 0xFFFFFFFFU;
  constexpr std::uint16_t kCanonicalDefaultRowHeightTwips = 300U;
  if (rec.payload.size < 12U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtWsFmtInfo payload truncated",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index));
  }

  ByteSpan p = rec.payload;
  auto dx_g_col_or = read_u32(p);
  auto cch_def_col_width_or = read_u16(p);
  auto miy_def_rw_height_or = read_u16(p);
  auto flags_or = read_u32(p);
  if (!dx_g_col_or || !cch_def_col_width_or || !miy_def_rw_height_or || !flags_or) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtWsFmtInfo fields truncated",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index));
  }
  if (p.size != 0U) {
    return make_error(
        FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo has trailing bytes",
        "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) + " trailing=" + std::to_string(p.size));
  }

  const std::uint32_t dx_g_col = dx_g_col_or.value();
  const std::uint16_t cch_def_col_width = cch_def_col_width_or.value();
  const std::uint16_t miy_def_rw_height = miy_def_rw_height_or.value();
  const std::uint32_t flags = flags_or.value();
  if (dx_g_col != kAbsentDefaultColumnWidth && dx_g_col > 65535U) {
    return make_error(
        FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo default column width out of range",
        "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) + " dxGCol=" + std::to_string(dx_g_col));
  }
  if (cch_def_col_width > 255U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo base column width out of range",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) +
                          " cchDefColWidth=" + std::to_string(cch_def_col_width));
  }

  const bool thick_top = (flags & 0x00000004U) != 0U;
  const bool thick_bottom = (flags & 0x00000008U) != 0U;
  const std::uint8_t row_outline_max = static_cast<std::uint8_t>((flags >> 16U) & 0xFFU);
  const std::uint8_t col_outline_max = static_cast<std::uint8_t>((flags >> 24U) & 0xFFU);
  if (row_outline_max > 7U || col_outline_max > 7U) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtWsFmtInfo outline level out of range",
                      "context=xlsb_reader sheet_index=" + std::to_string(sheet_index) + " row_outline_max=" +
                          std::to_string(row_outline_max) + " col_outline_max=" + std::to_string(col_outline_max));
  }

  SheetFormatDefaults& defaults = sheet.mutable_format_defaults();
  defaults.base_col_width = static_cast<double>(cch_def_col_width);
  if (dx_g_col == kAbsentDefaultColumnWidth) {
    defaults.has_default_col_width = false;
    defaults.default_col_width = 0.0;
  } else {
    defaults.has_default_col_width = true;
    defaults.default_col_width = static_cast<double>(dx_g_col) / 256.0;
  }

  const bool f_unsynced = (flags & 0x00000001U) != 0U;
  const bool f_dy_zero = (flags & 0x00000002U) != 0U;
  if (f_dy_zero) {
    defaults.has_default_row_height = true;
    defaults.default_row_height = 0.0;
  } else if (f_unsynced) {
    defaults.has_default_row_height = true;
    defaults.default_row_height = static_cast<double>(miy_def_rw_height) / 20.0;
  } else if (miy_def_rw_height == kCanonicalDefaultRowHeightTwips) {
    defaults.has_default_row_height = false;
    defaults.default_row_height = 0.0;
  } else {
    // Excel commonly writes FFFFFFFF/10/400/0 for an OOXML sheet whose
    // visible default row height is 20pt. Preserve that effective value for
    // compatibility even without fUnsynced.
    defaults.has_default_row_height = true;
    defaults.default_row_height = static_cast<double>(miy_def_rw_height) / 20.0;
  }

  if (thick_top || thick_bottom || row_outline_max != 0U || col_outline_max != 0U) {
    StructuredLog("xlsb.reader.unsupported_ws_format_metadata")
        .field("sheet_index", static_cast<std::int64_t>(sheet_index))
        .field("thick_top", thick_top)
        .field("thick_bottom", thick_bottom)
        .field("row_outline_max", static_cast<std::int64_t>(row_outline_max))
        .field("col_outline_max", static_cast<std::int64_t>(col_outline_max))
        .warn();
  }
  return Expected<void, Error>::Ok();
}

/// Tail records in this slot occur after the model-owned merge block and
/// before the model-owned hyperlink block, even when the source stream omits
/// BrtBeginMergeCells/BrtEndMergeCells entirely.  The worksheet grammar uses
/// these ids for phonetic metadata, conditional formatting and data
/// validations; relying only on `merges_seen` would put them before newly
/// added model hyperlinks on a no-merge sheet.
bool IsAfterMergesBeforeHyperlinks(XlsbRecordType type) {
  switch (static_cast<std::uint16_t>(type)) {
    case 537:  // BrtPhoneticInfo
    case 461:  // BrtBeginConditionalFormatting
    case 462:  // BrtEndConditionalFormatting
    case 463:  // BrtBeginCFRule
    case 464:  // BrtEndCFRule
    case 465:  // BrtBeginIconSet
    case 466:  // BrtEndIconSet
    case 467:  // BrtBeginDataBar
    case 468:  // BrtEndDataBar
    case 469:  // BrtBeginColorScale
    case 470:  // BrtEndColorScale
    case 471:  // BrtCFVO
    case 564:  // BrtColor (conditional-formatting color)
    case 573:  // BrtBeginDVals
    case 64:   // BrtDVal
    case 574:  // BrtEndDVals
      return true;
    default:
      return false;
  }
}

/// Tail records in this slot are emitted after the model-owned hyperlink
/// block.  This explicit grammar classification is needed when the source
/// has no BrtHLink at all but the in-memory caller adds one before writing.
bool IsAfterHyperlinks(XlsbRecordType type) {
  switch (static_cast<std::uint16_t>(type)) {
    case 477:  // BrtPrintOptions
    case 476:  // BrtMargins
    case 479:  // BrtBeginHeaderFooter
    case 480:  // BrtEndHeaderFooter
    case 478:  // BrtPageSetup
    case 392:  // BrtBeginRwBrk
    case 396:  // BrtBrk
    case 393:  // BrtEndRwBrk
    case 394:  // BrtBeginColBrk
    case 395:  // BrtEndColBrk
    case 625:  // BrtBigName
    case 605:  // BrtBeginCellWatches
    case 607:  // BrtCellWatch
    case 606:  // BrtEndCellWatches
    case 648:  // BrtBeginCellIgnoreECs
    case 649:  // BrtCellIgnoreEC
    case 650:  // BrtEndCellIgnoreECs
    case 594:  // BrtBeginSmartTags
    case 592:  // BrtBeginCellSmartTags
    case 590:  // BrtBeginCellSmartTag
    case 589:  // BrtCellSmartTagProperty
    case 591:  // BrtEndCellSmartTag
    case 593:  // BrtEndCellSmartTags
    case 595:  // BrtEndSmartTags
    case 550:  // BrtDrawing
    case 551:  // BrtLegacyDrawing
    case 552:  // BrtLegacyDrawingHF
    case 562:  // BrtBkhim
    case 638:  // BrtBeginOleObjects
    case 639:  // BrtOleObject
    case 640:  // BrtEndOleObjects
    case 643:  // BrtBeginActiveXControls
    case 644:  // BrtActiveX
    case 645:  // BrtEndActiveXControls
    case 554:  // BrtBeginWebPubItems
    case 556:  // BrtBeginWebPubItem
    case 557:  // BrtEndWebPubItem
    case 555:  // BrtEndWebPubItems
    case 660:  // BrtBeginTableParts
    case 661:  // BrtTablePart
    case 662:  // BrtEndTableParts
      return true;
    default:
      return false;
  }
}

/// Appends the framed bytes of one worksheet-tail record to the grammar slot
/// represented by `XlsbSheetTail`. `framed` spans the record header *and*
/// payload, so re-emission is a plain byte copy rather than a re-encode.
/// Slots only move forward: a record classified into a later slot advances
/// the phase, so an unlisted record after it (an x14 FRT block after
/// BrtMargins on a sheet with no merges or hyperlinks) is not hoisted ahead.
void RetainTailRecord(SheetDecodeState& state, XlsbRecordType type, const std::uint8_t* framed, std::size_t size) {
  std::vector<std::uint8_t>* dst = &state.tail.before_merges;
  if (state.hyperlinks_seen || IsAfterHyperlinks(type)) {
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

/// Decodes the model-owned tail features -- sheet protection, conditional
/// formatting blocks and the data-validation container -- which the writer
/// re-emits from `Sheet`. `rec` has just been read from `cursor`; a block's
/// remaining records are consumed here. A record or block holding content
/// outside the measured set is retained verbatim instead. Returns false
/// for any other record.
Expected<bool, Error> DecodeTailFeature(ByteSpan& cursor, const std::uint8_t* framed, const XlsbRecord& rec,
                                        SheetDecodeState& state, Sheet& sheet, const FeatureFormulaReadContext& ctx) {
  const auto type = static_cast<XlsbRecordType>(rec.type);
  if (rec.type == kBrtSheetProtection || rec.type == kBrtSheetProtectionIso) {
    SheetProtection& protection = sheet.mutable_protection();
    const bool ok = rec.type == kBrtSheetProtection ? decode_sheet_protection(rec.payload, protection)
                                                    : decode_sheet_protection_iso(rec.payload, protection);
    if (!ok) {
      StructuredLog("xlsb.protection.retained_raw").field("record_type", static_cast<std::int64_t>(rec.type)).warn();
      RetainTailRecord(state, type, framed, static_cast<std::size_t>(cursor.data - framed));
    }
    return true;
  }
  const bool is_cf = rec.type == kBrtBeginConditionalFormatting;
  if (!is_cf && rec.type != kBrtBeginDVals) {
    return false;
  }
  const std::uint16_t end_type = is_cf ? kBrtEndConditionalFormatting : kBrtEndDVals;
  for (;;) {
    auto next = read_record(cursor);
    if (!next) {
      return next.error();
    }
    if (next.value().type == end_type) {
      break;
    }
  }
  const ByteSpan block{framed, static_cast<std::size_t>(cursor.data - framed)};
  // Both block kinds sit after the merged cells; later records follow them.
  state.merges_seen = true;
  bool decoded = false;
  if (is_cf) {
    if (auto format = decode_cf_block(block, ctx)) {
      sheet.mutable_conditional_formats().push_back(std::move(*format));
      decoded = true;
    }
  } else if (auto validations = decode_dv_block(block, ctx)) {
    std::vector<DataValidation>& dst = sheet.mutable_validations();
    dst.insert(dst.end(), std::make_move_iterator(validations->begin()), std::make_move_iterator(validations->end()));
    decoded = true;
  }
  if (!decoded) {
    StructuredLog("xlsb.feature.retained_raw").field("record_type", static_cast<std::int64_t>(rec.type)).warn();
    RetainTailRecord(state, type, block.data, block.size);
  }
  return true;
}

/// Column + style-table index decoded by `ReadCellHeader`.
struct CellHeaderInfo {
  std::uint32_t col = 0;
  /// 0-based index into `StylesTable::cell_xfs`. `0` is the default xf
  /// and is never stored on a cell (mirrors the OOXML reader's
  /// `xf_index != 0` guard before calling `set_cell_xf_index`).
  std::uint32_t xf_index = 0;
};

/// Reads the eight-byte cell header common to every cell record:
///   * column   : u32 (zero-based)
///   * iStyleRef: u24 (index into `StylesTable::cell_xfs`)
///   * fPhShow  : u8  (Phonetic-text flag; ignored)
///
/// Returns the column index and style-xf index, and advances the
/// cursor past the header.
Expected<CellHeaderInfo, Error> ReadCellHeader(ByteSpan& cursor) {
  auto col_or = read_u32(cursor);
  if (!col_or) {
    return col_or.error();
  }
  if (col_or.value() >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb cell header column out of range",
                      "context=xlsb_reader");
  }
  // iStyleRef (3 bytes, little-endian) + fPhShow (1 byte) = 4 bytes.
  if (cursor.size < 4) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb cell header truncated (style/phonetic)",
                      "context=xlsb_reader");
  }
  const std::uint32_t xf_index = static_cast<std::uint32_t>(cursor.data[0]) |
                                 (static_cast<std::uint32_t>(cursor.data[1]) << 8) |
                                 (static_cast<std::uint32_t>(cursor.data[2]) << 16);
  cursor.data += 4;
  cursor.size -= 4;
  return CellHeaderInfo{col_or.value(), xf_index};
}

/// Stores `xf_index` on `(sheet_index, row, col)` when it differs from
/// the default xf (`0`), mirroring the OOXML sheet reader's
/// `xf_index != 0` guard before calling `Workbook::set_cell_xf_index`.
Expected<void, Error> ApplyXfIndex(Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                   std::uint32_t xf_index) {
  if (xf_index == 0) {
    return Expected<void, Error>::Ok();
  }
  return wb.set_cell_xf_index(sheet_index, row, col, xf_index);
}

/// Registers each recorded dynamic-array anchor as a spill region so the
/// footprint's non-anchor cells' raw literal payload (Excel writes a
/// plain `BrtCellRk` / `BrtCellReal` / ... for the cached spill targets
/// of e.g. `=SEQUENCE(3)` in F6, spilling into F7:F8) does not read back
/// as independent literals that block the anchor's own re-spill on
/// recalc. Mirrors the OOXML reader's `RegisterArraySpills`
/// (`io/sheet_reader.cpp`) byte-for-byte in intent: capture the
/// footprint's current cached values, blank the non-anchor cells, then
/// commit the spill region. Must run only after the entire sheet has
/// been decoded -- see the `BrtArrFmla` case's comment in
/// `DecodeSheetBin` for why registering inline (while later rows in the
/// footprint have not been decoded yet) does not work.
Expected<void, Error> RegisterArraySpills(Workbook& wb, std::size_t sheet_index,
                                          const std::vector<ArrayAnchor>& anchors) {
  // Validate and charge every footprint before the first reserve or cell
  // walk. Keep the budget local to this decoded sheet so independent sheets
  // cannot consume one another's dynamic-array allowance.
  ResourceBudget budget(kMaxDynamicArrayCells, FormulonErrorCode::kIoXlsbRecordCorrupt);
  for (const ArrayAnchor& a : anchors) {
    auto cells_or =
        checked_array_anchor_cells(a.row, a.col, a.last_row, a.last_col, FormulonErrorCode::kIoXlsbRecordCorrupt,
                                   "context=xlsb_reader array_anchor");
    if (!cells_or) {
      return cells_or.error();
    }
    std::string context("context=xlsb_reader format=xlsb anchor_row=");
    context.append(std::to_string(a.row));
    context.append(" anchor_col=");
    context.append(std::to_string(a.col));
    context.append(" last_row=");
    context.append(std::to_string(a.last_row));
    context.append(" last_col=");
    context.append(std::to_string(a.last_col));
    auto charged = consume_array_anchor_budget(budget, cells_or.value(), std::move(context));
    if (!charged) {
      return charged.error();
    }
  }

  for (const ArrayAnchor& a : anchors) {
    const std::uint32_t rows = a.last_row - a.row + 1U;
    const std::uint32_t cols = a.last_col - a.col + 1U;
    const std::uint64_t cell_count = static_cast<std::uint64_t>(rows) * cols;
    std::vector<Value> values;
    values.reserve(static_cast<std::size_t>(cell_count));
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        const Cell* cell = wb.sheet(sheet_index).cell_at(r, c);
        values.push_back(cell != nullptr ? cell->cached_value : Value::blank());
      }
    }
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        if (r == a.row && c == a.col) {
          continue;
        }
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(r, c, Value::blank());
      }
    }
    wb.sheet(sheet_index).commit_spill(a.row, a.col, rows, cols, std::move(values));
  }
  return Expected<void, Error>::Ok();
}

/// Classifies and decodes one worksheet record.
///
/// The switch below has no fall-through exit: every arm ends in a
/// `RecordDisposition`, and the `default` arm hands the record to
/// `ResolveUnmodelledRecord`. Because the function returns by value on
/// every path, an arm that ends without producing a disposition does not
/// compile, which is what keeps "decoded, retained, regenerated or
/// counted" exhaustive as records are added.
Expected<RecordDisposition, Error> DispatchSheetRecord(
    const XlsbRecord& rec, XlsbRecordType type, const std::uint8_t* framed, std::size_t framed_size,
    SheetDecodeState& state, std::size_t sheet_index, Workbook& wb, const std::vector<std::string_view>& sst_entries,
    const std::vector<std::vector<PhoneticRun>>& sst_phonetic,
    const std::vector<PhoneticProperties>& sst_phonetic_props, std::deque<std::string>& text_storage,
    const std::vector<std::string>& sheet_names, const std::vector<XlsbName>& name_table,
    const std::vector<XlsbSheetRange>& sheet_ranges, const XlsbExternalBooks& external_books,
    std::uint32_t* undecoded_formula_count) {
  switch (type) {
    case XlsbRecordType::BrtBeginWsView: {
      // BrtBeginWsView ([MS-XLSB] §2.4.141) stores the SheetView fields
      // modelled by Formulon in its flags word and wScale.  Other viewport
      // state is intentionally not represented by SheetView yet.
      ByteSpan p = rec.payload;
      auto flags_or = read_u16(p);
      auto skip_view = read_u32(p);
      auto skip_top = read_u32(p);
      auto skip_left = read_u32(p);
      auto skip_color = read_u8(p);
      auto skip_reserved8 = read_u8(p);
      auto skip_reserved16 = read_u16(p);
      auto zoom_or = read_u16(p);
      if (!flags_or || !skip_view || !skip_top || !skip_left || !skip_color || !skip_reserved8 || !skip_reserved16 ||
          !zoom_or) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtBeginWsView truncated",
                          "context=xlsb_reader");
      }
      SheetView& view = wb.sheet(sheet_index).mutable_view();
      const std::uint16_t flags = flags_or.value();
      view.show_grid_lines = (flags & 0x0004U) != 0U;
      view.show_row_col_headers = (flags & 0x0008U) != 0U;
      view.show_zeros = (flags & 0x0010U) != 0U;
      view.right_to_left = (flags & 0x0020U) != 0U;
      view.tab_selected = (flags & 0x0040U) != 0U;
      if (zoom_or.value() >= 10U && zoom_or.value() <= 400U) {
        view.zoom_scale = zoom_or.value();
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtWsProp: {
      auto complete_or = DecodeWorksheetProperties(rec, wb.sheet(sheet_index), sheet_index);
      if (!complete_or) {
        return complete_or.error();
      }
      if (!complete_or.value()) {
        StructuredLog("xlsb.record.dropped")
            .field("record_type", static_cast<std::int64_t>(type))
            .field("bytes", static_cast<std::int64_t>(framed_size))
            .field("reason", std::string_view("sheet properties outside <sheetPr>"))
            .warn();
        return RecordDisposition::kAccounted;
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtWsFmtInfo: {
      auto format_status = DecodeWorksheetFormatInfo(rec, wb.sheet(sheet_index), sheet_index);
      if (!format_status) {
        return format_status.error();
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtPane: {
      // BrtPane ([MS-XLSB] §2.4.723): two Xnum split positions, then
      // the lower-right pane's top-left cell and the frozen-state flags.
      // In a frozen pane the Xnum fields are integral row/column counts.
      ByteSpan p = rec.payload;
      if (p.size < 29U) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtPane truncated", "context=xlsb_reader");
      }
      double rows = 0.0;
      double cols = 0.0;
      std::memcpy(&rows, p.data, sizeof(rows));
      std::memcpy(&cols, p.data + sizeof(rows), sizeof(cols));
      p.data += 16U;
      p.size -= 16U;
      auto top_row_or = read_u32(p);
      auto left_col_or = read_u32(p);
      auto active_pane_or = read_u32(p);
      auto flags_or = read_u8(p);
      if (!top_row_or || !left_col_or || !active_pane_or || !flags_or) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtPane fields truncated",
                          "context=xlsb_reader");
      }
      const std::uint8_t frozen_flags = static_cast<std::uint8_t>(flags_or.value() & 0x03U);
      if (frozen_flags == 0U) {
        return RecordDisposition::kAccounted;
      }
      if (frozen_flags == 0x03U || !std::isfinite(rows) || !std::isfinite(cols) || rows < 0.0 || cols < 0.0 ||
          std::trunc(rows) != rows || std::trunc(cols) != cols || rows >= Sheet::kMaxRows || cols >= Sheet::kMaxCols) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtPane frozen counts invalid",
                          "context=xlsb_reader");
      }
      SheetView& view = wb.sheet(sheet_index).mutable_view();
      view.freeze_rows = static_cast<std::uint32_t>(rows);
      view.freeze_cols = static_cast<std::uint32_t>(cols);
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtRowHdr: {
      // BrtRowHdr ([MS-XLSB] §2.4.770): rw, ixfe, miyRw, two flag
      // bytes, fPhShow, then ccolspan. flags2 bit 0x40 (fGhostDirty) is
      // the sole row-style presence signal; ixfe is ignored without it.
      // Older minimal producers may provide only rw, which remains enough
      // for cell decoding.
      ByteSpan p = rec.payload;
      auto row_or = read_u32(p);
      if (!row_or) {
        return row_or.error();
      }
      if (row_or.value() >= Sheet::kMaxRows) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtRowHdr row out of range",
                          "context=xlsb_reader");
      }
      state.current_row = row_or.value();
      state.row_seen = true;
      if (p.size >= 9U) {
        auto row_style_xf_or = read_u32(p);
        auto height_or = read_u16(p);
        auto flags1_or = read_u8(p);
        auto flags2_or = read_u8(p);
        auto skip_phonetic_show = read_u8(p);
        if (!row_style_xf_or || !height_or || !flags1_or || !flags2_or || !skip_phonetic_show) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtRowHdr layout truncated",
                            "context=xlsb_reader");
        }
        const std::uint8_t flags = flags2_or.value();
        const bool has_style = (flags & 0x40U) != 0U;
        // flags2 bit 0x20 (fUnsynced) is the sole custom-height signal.
        // Excel stores the rendered height in `miyRw` for every row,
        // including rows left at the sheet default, so a non-zero `miyRw`
        // says nothing about whether the source made the height custom.
        // Mirrors the OOXML reader, where the override follows `ht`
        // presence rather than the `customHeight` attribute.
        const bool has_height = (flags & 0x20U) != 0U;
        if (has_height || (flags & 0x10U) != 0U || (flags & 0x07U) != 0U || has_style) {
          RowLayout layout;
          layout.row = state.current_row;
          if (has_height) {
            layout.height = static_cast<double>(height_or.value()) / 20.0;
            layout.has_height = true;
            layout.custom_height = true;
          }
          layout.hidden = (flags & 0x10U) != 0U;
          layout.outline_level = static_cast<std::uint8_t>(flags & 0x07U);
          layout.has_style = has_style;
          layout.style_xf = has_style ? row_style_xf_or.value() : 0U;
          wb.sheet(sheet_index).mutable_layout().row_overrides.push_back(std::move(layout));
        }
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtColInfo: {
      // BrtColInfo ([MS-XLSB] §2.4.336): colFirst, colLast, coldx,
      // ixfe, flags(u16). coldx is in 1/256 standard digits.
      ByteSpan p = rec.payload;
      auto first_or = read_u32(p);
      auto last_or = read_u32(p);
      auto width_or = read_u32(p);
      auto style_xf_or = read_u32(p);
      auto flags_or = read_u16(p);
      if (!first_or || !last_or || !width_or || !style_xf_or || !flags_or) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtColInfo truncated",
                          "context=xlsb_reader");
      }
      if (first_or.value() >= Sheet::kMaxCols || last_or.value() >= Sheet::kMaxCols ||
          first_or.value() > last_or.value()) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtColInfo range out of bounds",
                          "context=xlsb_reader");
      }
      ColumnLayout layout;
      layout.first = first_or.value();
      layout.last = last_or.value();
      layout.width = static_cast<double>(width_or.value()) / 256.0;
      layout.has_width = width_or.value() != 0U || (flags_or.value() & 0x0002U) != 0U;
      // BrtColInfo always carries ixfe; unlike OOXML there is no separate
      // style-presence bit. Treat the mandatory value as an effective style
      // (including ixfe=0), so an XLSB round-trip may canonicalize a
      // width/hidden-only span to explicit style 0 without changing its
      // rendering semantics.
      layout.has_style = true;
      layout.style_xf = style_xf_or.value();
      layout.hidden = (flags_or.value() & 0x0001U) != 0U;
      layout.outline_level = static_cast<std::uint8_t>((flags_or.value() >> 8U) & 0x07U);
      wb.sheet(sheet_index).mutable_layout().columns.push_back(std::move(layout));
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtMergeCell: {
      // BrtMergeCell ([MS-XLSB] §2.4.713): RfX = first/last row and
      // first/last column, all u32.  Containers are structural only; a
      // record may appear after sheet data and therefore must not depend on
      // the active BrtRowHdr state.
      ByteSpan p = rec.payload;
      auto first_row_or = read_u32(p);
      auto last_row_or = read_u32(p);
      auto first_col_or = read_u32(p);
      auto last_col_or = read_u32(p);
      if (!first_row_or || !last_row_or || !first_col_or || !last_col_or) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtMergeCell truncated",
                          "context=xlsb_reader");
      }
      if (first_row_or.value() >= Sheet::kMaxRows || last_row_or.value() >= Sheet::kMaxRows ||
          first_col_or.value() >= Sheet::kMaxCols || last_col_or.value() >= Sheet::kMaxCols ||
          first_row_or.value() > last_row_or.value() || first_col_or.value() > last_col_or.value()) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtMergeCell range out of bounds",
                          "context=xlsb_reader");
      }
      wb.sheet(sheet_index)
          .mutable_merges()
          .push_back(MergeRange{first_row_or.value(), first_col_or.value(), last_row_or.value(), last_col_or.value()});
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtHLink: {
      // BrtHLink ([MS-XLSB] §2.4.494): RfX followed by a non-null
      // relationship id and three XLWideStrings (location, tooltip,
      // display). A present-but-empty relationship id is the internal
      // hyperlink form; external links must resolve that id through the
      // sheet relationship part after the record stream is decoded.
      ByteSpan p = rec.payload;
      auto first_row_or = read_u32(p);
      auto last_row_or = read_u32(p);
      auto first_col_or = read_u32(p);
      auto last_col_or = read_u32(p);
      if (!first_row_or || !last_row_or || !first_col_or || !last_col_or) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtHLink range truncated",
                          "context=xlsb_reader");
      }
      if (!Sheet::rect_in_grid(first_row_or.value(), first_col_or.value(), last_row_or.value(), last_col_or.value())) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink range out of bounds",
                          "context=xlsb_reader");
      }
      // `read_xlnullablewidestring` intentionally maps both a null value
      // and a present empty value to `""`; inspect the sentinel first so a
      // null RelID cannot be mistaken for the required internal form.
      if (p.size < 4U) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtHLink relationship id truncated",
                          "context=xlsb_reader");
      }
      const std::uint32_t rel_len =
          static_cast<std::uint32_t>(p.data[0]) | (static_cast<std::uint32_t>(p.data[1]) << 8U) |
          (static_cast<std::uint32_t>(p.data[2]) << 16U) | (static_cast<std::uint32_t>(p.data[3]) << 24U);
      if (rel_len == 0xFFFFFFFFU) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink relationship id is null",
                          "context=xlsb_reader");
      }
      if (rel_len > kMaxHyperlinkRelIdUnits) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink relationship id exceeds string limit",
                          "context=xlsb_reader cch=" + std::to_string(rel_len));
      }
      auto rid_or = read_xlnullablewidestring(p);
      if (!rid_or) {
        return rid_or.error();
      }
      auto location_or = ReadHyperlinkWideString(p, "location", kMaxHyperlinkLocationUnits);
      if (!location_or) {
        return location_or.error();
      }
      auto tooltip_or = ReadHyperlinkWideString(p, "tooltip", kMaxHyperlinkTooltipUnits);
      if (!tooltip_or) {
        return tooltip_or.error();
      }
      auto display_or = ReadHyperlinkWideString(p, "display", kMaxHyperlinkDisplayUnits);
      if (!display_or) {
        return display_or.error();
      }
      if (p.size != 0U) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink has trailing bytes",
                          "context=xlsb_reader trailing=" + std::to_string(p.size));
      }
      Hyperlink hyperlink;
      hyperlink.row = first_row_or.value();
      hyperlink.col = first_col_or.value();
      hyperlink.last_row = last_row_or.value();
      hyperlink.last_col = last_col_or.value();
      hyperlink.rid = std::move(rid_or.value());
      hyperlink.location = std::move(location_or.value());
      hyperlink.tooltip = std::move(tooltip_or.value());
      hyperlink.display = std::move(display_or.value());
      wb.sheet(sheet_index).mutable_hyperlinks().push_back(std::move(hyperlink));
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellBlank: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::blank());
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellRk: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      auto rk_or = read_u32(p);
      if (!rk_or) {
        return rk_or.error();
      }
      const double v = decode_rk_number(rk_or.value());
      wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::number(v));
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellReal: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      if (p.size < 8) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtCellReal payload truncated",
                          "context=xlsb_reader");
      }
      double v;
      std::memcpy(&v, p.data, sizeof(v));
      wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::number(v));
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellBool: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      auto b_or = read_u8(p);
      if (!b_or) {
        return b_or.error();
      }
      wb.sheet(sheet_index)
          .set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::boolean(b_or.value() != 0));
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellError: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      auto code_or = read_u8(p);
      if (!code_or) {
        return code_or.error();
      }
      // Map the OOXML wire code to `ErrorCode` via the single
      // `kErrorTable`-backed lookup shared with the writer, so this
      // path stays symmetric with `ooxml_code()` (see
      // `error_from_ooxml_code` in `value.h`).
      const ErrorCode ec = error_from_ooxml_code(static_cast<std::int32_t>(code_or.value()));
      wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::error(ec));
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellSt: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      auto str_or = read_xlwidestring(p);
      if (!str_or) {
        return str_or.error();
      }
      text_storage.push_back(std::move(str_or.value()));
      wb.sheet(sheet_index)
          .set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::text(text_storage.back()));
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellIsst: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      auto idx_or = read_u32(p);
      if (!idx_or) {
        return idx_or.error();
      }
      if (idx_or.value() >= sst_entries.size()) {
        std::string ctx("context=xlsb_reader sheet_index=");
        ctx.append(std::to_string(sheet_index));
        ctx.append(" row=").append(std::to_string(state.current_row));
        ctx.append(" col=").append(std::to_string(col_or.value().col));
        ctx.append(" sst_index=").append(std::to_string(idx_or.value()));
        ctx.append(" sst_size=").append(std::to_string(sst_entries.size()));
        return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb sst index out of range", std::move(ctx));
      }
      wb.sheet(sheet_index)
          .set_cell_cached_value_borrowed(state.current_row, col_or.value().col,
                                          Value::text(sst_entries[idx_or.value()]));
      // Attached after the value, mirroring the OOXML reader: every
      // value-mutating setter clears the annotation, so the order is
      // load-bearing. Skipped when the entry carries no guide so an
      // unannotated cell keeps its default-constructed run vector.
      if (idx_or.value() < sst_phonetic.size() && !sst_phonetic[idx_or.value()].empty()) {
        wb.sheet(sheet_index)
            .set_cell_phonetic_runs(state.current_row, col_or.value().col, sst_phonetic[idx_or.value()]);
        if (idx_or.value() < sst_phonetic_props.size()) {
          wb.sheet(sheet_index)
              .set_cell_phonetic_props(state.current_row, col_or.value().col, sst_phonetic_props[idx_or.value()]);
        }
      }
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtFmlaNum:
    case XlsbRecordType::BrtFmlaString:
    case XlsbRecordType::BrtFmlaBool:
    case XlsbRecordType::BrtFmlaError: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      // All four formula records share the same prefix:
      //   cell-header (8 bytes)
      //   value       (8 bytes for Num, XLWideString for String, 2
      //                bytes for Bool/Error: u8 result + u8 flags)
      //   flags       (u16 grbitFlags)
      //   ptgs        (CellParsedFormula = u32 cce + cce bytes of
      //                 Rgce + ... we just slice the remainder as
      //                 opaque bytes for now)
      ByteSpan p = rec.payload;
      auto col_or = ReadCellHeader(p);
      if (!col_or) {
        return col_or.error();
      }
      if (const std::uint32_t ifmd = std::exchange(state.pending_cell_meta, 0U); ifmd != 0U) {
        state.cell_metadata.emplace_back(state.current_row, col_or.value().col, ifmd);
      }
      // Decode the formula's cached result so we can PRESERVE it on the
      // cell even when the Ptg stream cannot be decoded to a formula.
      Value cached = Value::blank();
      switch (type) {
        case XlsbRecordType::BrtFmlaNum: {
          if (p.size < 8) {
            return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFmlaNum value truncated",
                              "context=xlsb_reader");
          }
          double v;
          std::memcpy(&v, p.data, sizeof(v));
          cached = Value::number(v);
          p.data += 8;
          p.size -= 8;
          break;
        }
        case XlsbRecordType::BrtFmlaString: {
          auto s = read_xlwidestring(p);
          if (!s) {
            return s.error();
          }
          text_storage.push_back(std::move(s.value()));
          cached = Value::text(text_storage.back());
          break;
        }
        case XlsbRecordType::BrtFmlaBool: {
          auto b = read_u8(p);
          if (!b) {
            return b.error();
          }
          cached = Value::boolean(b.value() != 0);
          break;
        }
        case XlsbRecordType::BrtFmlaError: {
          auto b = read_u8(p);
          if (!b) {
            return b.error();
          }
          // Same `kErrorTable`-backed lookup the literal path uses.
          cached = Value::error(error_from_ooxml_code(static_cast<std::int32_t>(b.value())));
          break;
        }
        default:
          break;
      }
      // grbitFlags (u16).
      if (p.size < 2) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula flags truncated",
                          "context=xlsb_reader");
      }
      p.data += 2;
      p.size -= 2;
      // CellParsedFormula: u32 cce (rgce byte length) + cce bytes of
      // Ptg stream + u32 cb + cb bytes of rgcb (the array-constant
      // extra-data area `PtgArray` consumes; empty for formulas with
      // no array literals).
      auto cce_or = read_u32(p);
      if (!cce_or) {
        return cce_or.error();
      }
      const std::uint32_t cce = cce_or.value();
      if (cce > p.size) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula rgce length exceeds payload",
                          "context=xlsb_reader");
      }
      ByteSpan rgce{p.data, cce};
      p.data += cce;
      p.size -= cce;
      ByteSpan rgcb{};
      if (p.size >= 4) {
        auto cb_or = read_u32(p);
        if (!cb_or) {
          return cb_or.error();
        }
        const std::uint32_t cb = cb_or.value();
        if (cb > p.size) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula rgcb length exceeds payload",
                            "context=xlsb_reader");
        }
        rgcb = ByteSpan{p.data, cb};
      }
      const std::string formula_text =
          DecodeFormulaText(rgce, rgcb, sheet_names, name_table, sheet_ranges, external_books, sheet_index,
                            state.current_row, col_or.value().col, undecoded_formula_count);
      if (!formula_text.empty()) {
        // Register the real formula via the workbook-level entry so the
        // dep graph tracks it (matching the OOXML reader). The cached
        // value is preserved separately below.
        auto wf = wb.set_cell_formula(sheet_index, state.current_row, col_or.value().col, formula_text);
        if (!wf) {
          return wf.error();
        }
      }
      // Always preserve the cached value (for undecodable formulas this
      // is the only correct datum the cell carries; for decoded ones it
      // matches Excel's stored result until the next recalc).
      if (!cached.is_blank()) {
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, cached);
      }
      if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index); !r) {
        return r.error();
      }
      ++state.cells_decoded;
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellMeta: {
      // The cell-metadata index (the dynamic-array entry) of the next cell;
      // the record itself is re-derived by the writer.
      ByteSpan p = rec.payload;
      if (auto ifmd = read_u32(p); ifmd) {
        state.pending_cell_meta = ifmd.value();
      }
      return RecordDisposition::kRegenerated;
    }
    case XlsbRecordType::BrtArrFmla: {
      if (!state.row_seen) {
        return RecordDisposition::kAccounted;
      }
      // BrtArrFmla ([MS-XLSB], verified against a real Excel-produced
      // `xl/worksheets/sheetN.bin`): RfX (4 x u32: rwFirst, rwLast,
      // colFirst, colLast) + 1 reserved/flag byte + CellParsedFormula
      // (u32 cce + cce bytes rgce + u32 cb + cb bytes rgcb). This
      // record supplies the REAL Ptg tokens for a CSE / dynamic-array
      // formula whose anchor cell's own formula-result "shell" record
      // (BrtFmlaNum/String/Bool/Error, processed above) carries only a
      // `PtgExp` placeholder. Only the anchor cell (`rwFirst`,
      // `colFirst`) gets a formula string; the rest of the array range
      // (if any) has no formula of its own — matching the OOXML
      // reader's treatment of `t="array"` / dynamic-array spill
      // formulas, where only the anchor cell stores `<f>`.
      ByteSpan p = rec.payload;
      auto rw_first_or = read_u32(p);
      if (!rw_first_or) {
        return rw_first_or.error();
      }
      auto rw_last_or = read_u32(p);
      if (!rw_last_or) {
        return rw_last_or.error();
      }
      auto col_first_or = read_u32(p);
      if (!col_first_or) {
        return col_first_or.error();
      }
      auto col_last_or = read_u32(p);
      if (!col_last_or) {
        return col_last_or.error();
      }
      // The RfX rect must lie inside the grid and be well-ordered on
      // BOTH axes before any of it is used: the anchor guard below is
      // an OR, so a rect reversed on only one axis would otherwise
      // still be recorded and wrap the size math in
      // `RegisterArraySpills`.
      if (rw_first_or.value() >= Sheet::kMaxRows || rw_last_or.value() >= Sheet::kMaxRows ||
          col_first_or.value() >= Sheet::kMaxCols || col_last_or.value() >= Sheet::kMaxCols ||
          rw_last_or.value() < rw_first_or.value() || col_last_or.value() < col_first_or.value()) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtArrFmla range out of bounds",
                          "context=xlsb_reader");
      }
      if (p.size < 1) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla flag truncated",
                          "context=xlsb_reader");
      }
      p.data += 1;
      p.size -= 1;
      auto cce_or = read_u32(p);
      if (!cce_or) {
        return cce_or.error();
      }
      const std::uint32_t cce = cce_or.value();
      if (cce > p.size) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla rgce length exceeds payload",
                          "context=xlsb_reader");
      }
      ByteSpan rgce{p.data, cce};
      p.data += cce;
      p.size -= cce;
      ByteSpan rgcb{};
      if (p.size >= 4) {
        auto cb_or = read_u32(p);
        if (!cb_or) {
          return cb_or.error();
        }
        const std::uint32_t cb = cb_or.value();
        if (cb > p.size) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla rgcb length exceeds payload",
                            "context=xlsb_reader");
        }
        rgcb = ByteSpan{p.data, cb};
      }
      const std::string formula_text =
          DecodeFormulaText(rgce, rgcb, sheet_names, name_table, sheet_ranges, external_books, sheet_index,
                            rw_first_or.value(), col_first_or.value(), undecoded_formula_count);
      if (formula_text.empty()) {
        return RecordDisposition::kModelled;
      }
      // The anchor's shell record, read before this one, carries its cached value.
      const Cell* shell = wb.sheet(sheet_index).cell_at(rw_first_or.value(), col_first_or.value());
      Value cached = shell != nullptr ? shell->cached_value : Value::blank();
      const std::string cached_text = cached.is_text() ? std::string(cached.as_text()) : std::string();
      auto wf = wb.set_cell_formula(sheet_index, rw_first_or.value(), col_first_or.value(), formula_text);
      if (!wf) {
        return wf.error();
      }
      if (cached.is_text()) {
        cached = Value::text(cached_text);  // the cell's own copy went with the shell
      }
      wb.sheet(sheet_index).set_cell_cached_value(rw_first_or.value(), col_first_or.value(), cached);
      // Record the footprint for a second pass after the whole sheet
      // has been decoded (see `RegisterArraySpills`, called at the end
      // of this function). `BrtArrFmla` for the anchor `(rwFirst,
      // colFirst)` appears within the anchor's OWN row group in the
      // record stream, i.e. BEFORE the non-anchor rows' own cell
      // records (`rwLast > rwFirst` spans into rows not yet decoded).
      // Registering the spill immediately here would have those later
      // records' `set_cell_cached_value` calls silently repopulate the
      // phantom cells with their raw literal payload, which a
      // subsequent recalc's spill-commit would then see as "already
      // occupied" and surface `#SPILL!` instead of the real result.
      state.array_anchors.push_back(
          ArrayAnchor{rw_first_or.value(), col_first_or.value(), rw_last_or.value(), col_last_or.value()});
      return RecordDisposition::kModelled;
    }
    default:
      return ResolveUnmodelledRecord(state, type, framed, framed_size);
  }
}
}  // namespace

Expected<SheetDecodeState, Error> DecodeSheetBin(
    const std::vector<std::uint8_t>& body, std::size_t sheet_index, Workbook& wb,
    const std::vector<std::string_view>& sst_entries, const std::vector<std::vector<PhoneticRun>>& sst_phonetic,
    const std::vector<PhoneticProperties>& sst_phonetic_props, std::deque<std::string>& text_storage,
    const std::vector<std::string>& sheet_names, const std::vector<XlsbName>& name_table,
    const std::vector<XlsbSheetRange>& sheet_ranges, const XlsbExternalBooks& external_books,
    std::uint32_t* undecoded_formula_count) {
  SheetDecodeState state;
  const FeatureFormulaReadContext feature_ctx{sheet_names, name_table, sheet_ranges, external_books, sheet_index};
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    const std::uint8_t* const framed = cursor.data;
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    const auto type = static_cast<XlsbRecordType>(rec.type);
    const auto framed_size = static_cast<std::size_t>(cursor.data - framed);
    if (type == XlsbRecordType::BrtHLink) {
      state.hyperlinks_seen = true;
    }
    if (state.in_tail) {
      auto feature = DecodeTailFeature(cursor, framed, rec, state, wb.sheet(sheet_index), feature_ctx);
      if (!feature) {
        return feature.error();
      }
      if (feature.value()) {
        continue;
      }
    }
    // Every record resolves to a disposition; the result is consumed rather
    // than discarded so that no record can pass through unclassified.
    auto disposition_or = DispatchSheetRecord(rec, type, framed, framed_size, state, sheet_index, wb, sst_entries,
                                              sst_phonetic, sst_phonetic_props, text_storage, sheet_names, name_table,
                                              sheet_ranges, external_books, undecoded_formula_count);
    if (!disposition_or) {
      return disposition_or.error();
    }
    if (disposition_or.value() == RecordDisposition::kAccounted) {
      ++state.dropped_records;
    }
    // Grammar phase advances after the record has been dispatched, so the
    // marker itself belongs to the phase it closes rather than the one it
    // opens.
    if (type == XlsbRecordType::BrtEndSheetData) {
      state.in_tail = true;
    } else if (type == XlsbRecordType::BrtEndMergeCells) {
      state.merges_seen = true;
    }
  }
  auto spills = RegisterArraySpills(wb, sheet_index, state.array_anchors);
  if (!spills) {
    return spills.error();
  }
  if (!state.tail.empty()) {
    std::vector<cf::ConditionalFormat>& formats = wb.sheet(sheet_index).mutable_conditional_formats();
    for (const std::vector<std::uint8_t>* slot :
         {&state.tail.before_merges, &state.tail.after_merges_before_hyperlinks, &state.tail.after_hyperlinks}) {
      apply_x14_data_bar_overlays(ByteSpan{slot->data(), slot->size()}, formats);
    }
    for (const XlsbSheetRange& range : sheet_ranges) {
      XlsbExternSheetEntry entry;
      const bool sheetless = range.itab_first == -2 && range.itab_last == -2;
      const auto in_book = [&sheet_names](std::int32_t itab) {
        return itab >= 0 && static_cast<std::size_t>(itab) < sheet_names.size();
      };
      entry.unresolved =
          range.external_book != 0U || (!sheetless && !(in_book(range.itab_first) && in_book(range.itab_last)));
      if (!sheetless && !entry.unresolved) {
        entry.first = sheet_names[static_cast<std::size_t>(range.itab_first)];
        entry.last = sheet_names[static_cast<std::size_t>(range.itab_last)];
      }
      state.tail.extern_sheets.push_back(std::move(entry));
    }
    wb.sheet(sheet_index).set_xlsb_tail(state.tail);
  }
  return state;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
