#include "io/xlsb/sheet_reader.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <iterator>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/array_anchor_budget.h"
#include "io/protection_xml.h"
#include "io/rich_value_errors.h"
#include "io/sheet_reader.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/dv_records.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/hyperlink_records.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/sheet_properties_records.h"
#include "io/xlsb/sheet_reader_internal.h"
#include "io/xlsb/sheet_record_flags.h"
#include "io/xlsb/sparkline_records.h"
#include "io/xlsb/view_records.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "phonetic.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using detail::CellHeaderInfo;
using detail::ChargeLoad;
using detail::CommitCell;
using detail::DecodeFormulaText;
using detail::DecodeLiteralCell;
using detail::ReadCellHeader;
using detail::ReadCellParsedFormula;
using detail::RecordDisposition;
using detail::ResolveUnmodelledRecord;
using detail::RetainTailRecord;

/// Decodes a `BrtTable` payload ([MS-XLSB] 2.4.846, record 428): RfX, rwInput1 /
/// colInput1, rwInput2 / colInput2 (u32 each), then fRw, fTbl2, fDeleted1,
/// fDeleted2, fAlwaysCalc in the low bits of one byte. An input the flags
/// say to ignore is not validated. False when the record is short or names
/// cells off the grid.
bool DecodeDataTableRecord(ByteSpan p, DataTable* out) {
  auto rfx = read_rfx(p);
  if (!rfx) {
    return false;
  }
  std::uint32_t cells[4] = {};
  for (std::uint32_t& cell : cells) {
    auto v = read_u32(p);
    if (!v) {
      return false;
    }
    cell = v.value();
  }
  auto flags = read_u8(p);
  if (!flags) {
    return false;
  }
  const MergeRange ref = rfx.value();
  if (ref.last_row < ref.first_row || ref.last_col < ref.first_col || ref.first_row == 0U || ref.first_col == 0U ||
      !Sheet::coord_in_grid(ref.last_row, ref.last_col)) {
    return false;
  }
  DataTable table;
  table.ref = ref;
  table.row_oriented = (flags.value() & kTableRowInput) != 0U;
  table.two_dimensional = (flags.value() & kTableTwoDimensional) != 0U;
  table.always_calc = (flags.value() & kTableAlwaysCalc) != 0U;
  const auto input = [&](bool present, bool deleted, std::uint32_t row, std::uint32_t col, DataTableInput* in) {
    in->present = present;
    in->deleted = present && deleted;
    if (!present || in->deleted) {
      return true;
    }
    in->row = row;
    in->col = col;
    return Sheet::coord_in_grid(row, col);
  };
  if (!input(true, (flags.value() & kTableDeletedInput1) != 0U, cells[0], cells[1], &table.r1) ||
      !input(table.two_dimensional, (flags.value() & kTableDeletedInput2) != 0U, cells[2], cells[3], &table.r2)) {
    return false;
  }
  // Spelled out on an xlsx save, as for a table built here; `aca` has no XLSB bit.
  table.explicit_attrs = kDataTableExplicitDt2D | kDataTableExplicitDtr | kDataTableExplicitDel1 |
                         kDataTableExplicitDel2 | kDataTableExplicitCa;
  *out = table;
  return true;
}

/// Decodes the model-owned tail features -- sheet and range protection,
/// conditional formatting blocks and the data-validation container -- which the writer
/// re-emits from `Sheet`. `rec` has just been read from `cursor`; a block's
/// remaining records are consumed here. A record or block holding content
/// outside the supported set is retained verbatim instead. Returns false
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
  if (rec.type == kBrtRangeProtection || rec.type == kBrtRangeProtectionIso) {
    // An Iso record is decoded together with the BrtRangeProtection after it.
    ByteSpan iso{};
    ByteSpan payload = rec.payload;
    bool paired = rec.type == kBrtRangeProtection;
    if (!paired) {
      ByteSpan peek = cursor;
      auto next = read_record(peek);
      if (next && next.value().type == kBrtRangeProtection) {
        iso = rec.payload;
        payload = next.value().payload;
        cursor = peek;
        paired = true;
      }
    }
    ProtectedRange range;
    if (paired && decode_range_protection(payload, iso, range)) {
      state.protected_ranges.push_back(std::move(range));
    } else {
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
      return std::move(next.error());
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

/// Registers each recorded dynamic-array anchor as a spill region so the
/// footprint's non-anchor cells' raw literal payload (Excel writes a
/// plain `BrtCellRk` / `BrtCellReal` / ... for the cached spill targets
/// of e.g. `=SEQUENCE(3)` in F6, spilling into F7:F8) does not read back
/// as independent literals that block the anchor's own re-spill on
/// recalc. Shares `register_array_spills` with the OOXML reader. Must
/// run only after the entire sheet has
/// been decoded -- see the `BrtArrFmla` case's comment in
/// `DecodeSheetBin` for why registering inline (while later rows in the
/// footprint have not been decoded yet) does not work.
Expected<void, Error> RegisterArraySpills(Workbook& wb, std::size_t sheet_index,
                                          const std::vector<ArrayAnchor>& anchors) {
  return register_array_spills(wb.sheet(sheet_index), anchors, FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb_reader",
                               "xlsb");
}

/// Materialises row visibility after all worksheet records have been read.
///
/// A complete BrtRowHdr is an explicit visible-row record when BrtWsFmtInfo's
/// fDyZero says rows are hidden by default. When that default is clear, a
/// header with no row flags is only the structural prefix for its cells and
/// must not become a model override. Deferring this decision also handles
/// producers that place BrtWsFmtInfo after the row stream. For duplicate row
/// headers, the first effective header wins, matching insert_row_override's
/// reader contract; a structural header only wins when the sheet default
/// makes its visibility explicit.
void ApplyDecodedRowHeaders(Sheet& sheet, const std::vector<SheetDecodeState::DecodedRowHeader>& headers) {
  const bool default_hidden = sheet.format_defaults().zero_height;
  std::unordered_set<std::uint32_t> seen_rows;
  seen_rows.reserve(headers.size());
  for (const SheetDecodeState::DecodedRowHeader& header : headers) {
    if (!header.has_override && !default_hidden) {
      continue;
    }
    if (!seen_rows.insert(header.layout.row).second) {
      continue;
    }
    if (header.has_override) {
      insert_row_override(sheet.mutable_layout(), header.layout);
      continue;
    }
    RowLayout visible;
    visible.row = header.layout.row;
    insert_row_override(sheet.mutable_layout(), visible);
  }
}

/// Workbook-level tables a worksheet record decodes against.
struct SheetDecodeInputs {
  const std::vector<std::string_view>& sst_entries;
  const std::vector<std::vector<PhoneticRun>>& sst_phonetic;
  const std::vector<PhoneticProperties>& sst_phonetic_props;
  std::deque<std::string>& text_storage;
  const std::vector<std::string>& sheet_names;
  const std::vector<XlsbName>& name_table;
  const std::vector<XlsbSheetRange>& sheet_ranges;
  const XlsbExternalBooks& external_books;
  std::uint32_t* undecoded_formula_count;
};

/// Classifies and decodes one worksheet record.
///
/// The switch below has no fall-through exit: every arm ends in a
/// `RecordDisposition`, and the `default` arm hands the record to
/// `ResolveUnmodelledRecord`. Because the function returns by value on
/// every path, an arm that ends without producing a disposition does not
/// compile, which is what keeps "decoded, retained, regenerated or
/// counted" exhaustive as records are added.
Expected<RecordDisposition, Error> DispatchSheetRecord(const XlsbRecord& rec, XlsbRecordType type,
                                                       const std::uint8_t* framed, std::size_t framed_size,
                                                       SheetDecodeState& state, std::size_t sheet_index, Workbook& wb,
                                                       const SheetDecodeInputs& in) {
  switch (type) {
    case XlsbRecordType::BrtBeginWsView: {
      // The first view is the modelled one; each later one is another
      // workbook window's, kept as `extra_views_xml`.
      SheetView* target = &wb.sheet(sheet_index).mutable_view();
      if (state.main_ws_view_seen) {
        state.extra_view = SheetView{};
        target = &state.extra_view;
        state.in_extra_ws_view = true;
      } else {
        state.main_ws_view_seen = true;
        state.in_main_ws_view = true;
      }
      std::uint32_t workbook_view = 0U;
      if (auto view_status = decode_ws_view(rec, *target, &workbook_view); !view_status) {
        return std::move(view_status.error());
      }
      if (state.in_extra_ws_view) {
        state.extra_workbook_view = workbook_view;
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtWsProp: {
      auto complete_or = decode_ws_prop(rec, wb.sheet(sheet_index), sheet_index);
      if (!complete_or) {
        return std::move(complete_or.error());
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
      auto format_status = decode_ws_fmt_info(rec, wb.sheet(sheet_index), sheet_index);
      if (!format_status) {
        return std::move(format_status.error());
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtPane: {
      SheetView* view = state.open_view(wb.sheet(sheet_index));
      if (view == nullptr) {
        return ResolveUnmodelledRecord(state, type, framed, framed_size);
      }
      auto pane_or = decode_pane(rec, *view);
      if (!pane_or) {
        return std::move(pane_or.error());
      }
      return pane_or.value() ? RecordDisposition::kModelled : ResolveUnmodelledRecord(state, type, framed, framed_size);
    }
    case XlsbRecordType::BrtSel: {
      SheetView* view = state.open_view(wb.sheet(sheet_index));
      if (view == nullptr) {
        return ResolveUnmodelledRecord(state, type, framed, framed_size);
      }
      return decode_selection(rec, *view) ? RecordDisposition::kModelled
                                          : ResolveUnmodelledRecord(state, type, framed, framed_size);
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
        return std::move(row_or.error());
      }
      if (row_or.value() >= Sheet::kMaxRows) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtRowHdr row out of range", kReadXlsbContext);
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
                            kReadXlsbContext);
        }
        const std::uint8_t flags = flags2_or.value();
        const bool has_style = (flags & kRowGhostDirty) != 0U;
        // flags2 bit fUnsynced is the sole custom-height signal.
        // Excel stores the rendered height in `miyRw` for every row,
        // including rows left at the sheet default, so a non-zero `miyRw`
        // says nothing about whether the source made the height custom.
        // Mirrors the OOXML reader, where the override follows `ht`
        // presence rather than the `customHeight` attribute.
        const bool has_height = (flags & kRowUnsynced) != 0U;
        RowLayout layout;
        layout.row = state.current_row;
        if (has_height) {
          layout.height = static_cast<double>(height_or.value()) / 20.0;
          layout.has_height = true;
          layout.custom_height = true;
        }
        layout.hidden = (flags & kRowZeroHeight) != 0U;
        layout.outline_level = static_cast<std::uint8_t>(flags & kRowOutlineLevelMask);
        layout.collapsed = (flags & kRowCollapsed) != 0U;
        layout.has_style = has_style;
        layout.style_xf = has_style ? row_style_xf_or.value() : 0U;
        // Buffered until the record loop ends, so every header is charged.
        RETURN_IF_ERROR(ChargeLoad(state, sheet_index, sizeof(SheetDecodeState::DecodedRowHeader), 0U));
        state.row_headers.push_back(SheetDecodeState::DecodedRowHeader{
            layout, has_height || layout.hidden || layout.outline_level != 0U || layout.collapsed || has_style});
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
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtColInfo truncated", kReadXlsbContext);
      }
      if (first_or.value() >= Sheet::kMaxCols || last_or.value() >= Sheet::kMaxCols ||
          first_or.value() > last_or.value()) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtColInfo range out of bounds",
                          kReadXlsbContext);
      }
      ColumnLayout layout;
      layout.first = first_or.value();
      layout.last = last_or.value();
      layout.width = static_cast<double>(width_or.value()) / 256.0;
      layout.has_width = width_or.value() != 0U || (flags_or.value() & kColUserSet) != 0U;
      // BrtColInfo always carries ixfe; unlike OOXML there is no separate
      // style-presence bit. Treat the mandatory value as an effective style
      // (including ixfe=0), so an XLSB round-trip may canonicalize a
      // width/hidden-only span to explicit style 0 without changing its
      // rendering semantics.
      layout.has_style = true;
      layout.style_xf = style_xf_or.value();
      layout.hidden = (flags_or.value() & kColHidden) != 0U;
      layout.outline_level = static_cast<std::uint8_t>((flags_or.value() >> kColOutlineShift) & kColOutlineLevelMask);
      layout.collapsed = (flags_or.value() & kColCollapsed) != 0U;
      wb.sheet(sheet_index).mutable_layout().columns.push_back(std::move(layout));
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtEndColInfos:
      // Cells decode after this record, so the spans must be canonical before
      // any of them consults the column formats. The writer regenerates it.
      normalize_column_spans(wb.sheet(sheet_index).mutable_layout());
      return RecordDisposition::kRegenerated;
    case XlsbRecordType::BrtMergeCell: {
      // BrtMergeCell ([MS-XLSB] §2.4.713): RfX = first/last row and
      // first/last column, all u32.  Containers are structural only; a
      // record may appear after sheet data and therefore must not depend on
      // the active BrtRowHdr state.
      ByteSpan p = rec.payload;
      auto rfx_or = read_rfx(p);
      if (!rfx_or) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtMergeCell truncated", kReadXlsbContext);
      }
      const MergeRange& rfx = rfx_or.value();
      if (!in_grid(rfx)) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtMergeCell range out of bounds",
                          kReadXlsbContext);
      }
      wb.sheet(sheet_index).mutable_merges().push_back(rfx);
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtHLink: {
      auto link_status = decode_hyperlink(rec, wb.sheet(sheet_index));
      if (!link_status) {
        return std::move(link_status.error());
      }
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtCellBlank:
    case XlsbRecordType::BrtCellRk:
    case XlsbRecordType::BrtCellReal:
    case XlsbRecordType::BrtCellBool:
    case XlsbRecordType::BrtCellError:
    case XlsbRecordType::BrtCellSt:
    case XlsbRecordType::BrtCellIsst:
      return DecodeLiteralCell(rec, type, state, sheet_index, wb, in.sst_entries, in.sst_phonetic,
                               in.sst_phonetic_props, in.text_storage);
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
      ASSIGN_OR_RETURN(auto col, ReadCellHeader(p));
      if (const std::uint32_t ifmd = std::exchange(state.pending_cell_meta, 0U); ifmd != 0U) {
        state.cell_metadata.emplace_back(state.current_row, col.col, ifmd);
      }
      if (const std::uint32_t vm = std::exchange(state.pending_value_meta, 0U); vm != 0U) {
        state.value_metadata.push_back(RichErrorCell{state.current_row, col.col, vm});
      }
      // Decode the formula's cached result so we can PRESERVE it on the
      // cell even when the Ptg stream cannot be decoded to a formula.
      Value cached = Value::blank();
      switch (type) {
        case XlsbRecordType::BrtFmlaNum: {
          if (p.size < 8) {
            return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFmlaNum value truncated",
                              kReadXlsbContext);
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
            return std::move(s.error());
          }
          in.text_storage.push_back(std::move(s.value()));
          cached = Value::text(in.text_storage.back());
          break;
        }
        case XlsbRecordType::BrtFmlaBool: {
          auto b = read_u8(p);
          if (!b) {
            return std::move(b.error());
          }
          cached = Value::boolean(b.value() != 0);
          break;
        }
        case XlsbRecordType::BrtFmlaError: {
          auto b = read_u8(p);
          if (!b) {
            return std::move(b.error());
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
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula flags truncated", kReadXlsbContext);
      }
      p.data += 2;
      p.size -= 2;
      // CellParsedFormula: u32 cce (rgce byte length) + cce bytes of
      // Ptg stream + u32 cb + cb bytes of rgcb (the array-constant
      // extra-data area `PtgArray` consumes; empty for formulas with
      // no array literals).
      ByteSpan rgce;
      ByteSpan rgcb;
      RETURN_IF_ERROR(ReadCellParsedFormula(p, "formula", &rgce, &rgcb));
      const std::string formula_text =
          DecodeFormulaText(rgce, rgcb, in.sheet_names, in.name_table, in.sheet_ranges, in.external_books, sheet_index,
                            state.current_row, col.col, in.undecoded_formula_count, &wb.tables());
      RETURN_IF_ERROR(ChargeLoad(state, sheet_index,
                                 sheet_load_write_bytes(wb.sheet(sheet_index), state.current_row, col.col, true),
                                 col.col));
      if (!formula_text.empty()) {
        // Register the real formula via the workbook-level entry so the
        // dep graph tracks it (matching the OOXML reader). The cached
        // value is preserved separately below.
        auto wf = wb.set_cell_formula(sheet_index, state.current_row, col.col,
                                      wb.ingest_stored_formula(formula_text, Workbook::FormulaSource::kRendered),
                                      Workbook::FormulaSource::kRendered);
        if (!wf) {
          return std::move(wf.error());
        }
      }
      // Always preserve the cached value (for undecodable formulas this
      // is the only correct datum the cell carries; for decoded ones it
      // matches Excel's stored result until the next recalc).
      if (!cached.is_blank()) {
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col.col, cached);
      }
      // Excel drops a formula cell's phonetic flag on load.
      col.show_phonetic = false;
      RETURN_IF_ERROR(CommitCell(wb, sheet_index, state, col, /*explicit_xf=*/true));
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
    case XlsbRecordType::BrtValueMeta: {
      // The value-metadata index (the rich value) of the next cell; the
      // record itself is re-derived by the writer.
      ByteSpan p = rec.payload;
      if (auto vm = read_u32(p); vm) {
        state.pending_value_meta = vm.value();
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
      ASSIGN_OR_RETURN(const MergeRange rfx, read_rfx(p));
      const std::uint32_t rw_first = rfx.first_row;
      const std::uint32_t rw_last = rfx.last_row;
      const std::uint32_t col_first = rfx.first_col;
      const std::uint32_t col_last = rfx.last_col;
      // The RfX rect must lie inside the grid and be well-ordered on
      // BOTH axes before any of it is used: the anchor guard below is
      // an OR, so a rect reversed on only one axis would otherwise
      // still be recorded and wrap the size math in
      // `RegisterArraySpills`.
      if (rw_first >= Sheet::kMaxRows || rw_last >= Sheet::kMaxRows || col_first >= Sheet::kMaxCols ||
          col_last >= Sheet::kMaxCols || rw_last < rw_first || col_last < col_first) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtArrFmla range out of bounds",
                          kReadXlsbContext);
      }
      if (p.size < 1) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla flag truncated",
                          kReadXlsbContext);
      }
      p.data += 1;
      p.size -= 1;
      ByteSpan rgce;
      ByteSpan rgcb;
      RETURN_IF_ERROR(ReadCellParsedFormula(p, "BrtArrFmla", &rgce, &rgcb));
      const std::string formula_text =
          DecodeFormulaText(rgce, rgcb, in.sheet_names, in.name_table, in.sheet_ranges, in.external_books, sheet_index,
                            rw_first, col_first, in.undecoded_formula_count, &wb.tables());
      if (formula_text.empty()) {
        return RecordDisposition::kModelled;
      }
      // The anchor's shell record, read before this one, carries its cached value.
      const Cell* shell = wb.sheet(sheet_index).cell_at(rw_first, col_first);
      Value cached = shell != nullptr ? shell->cached_value : Value::blank();
      const std::string cached_text = cached.is_text() ? std::string(cached.as_text()) : std::string();
      auto wf = wb.set_cell_formula(sheet_index, rw_first, col_first,
                                    wb.ingest_stored_formula(formula_text, Workbook::FormulaSource::kRendered),
                                    Workbook::FormulaSource::kRendered);
      if (!wf) {
        return std::move(wf.error());
      }
      if (cached.is_text()) {
        cached = Value::text(cached_text);  // the cell's own copy went with the shell
      }
      wb.sheet(sheet_index).set_cell_cached_value(rw_first, col_first, cached);
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
      state.array_anchors.push_back(ArrayAnchor{rw_first, col_first, rw_last, col_last});
      return RecordDisposition::kModelled;
    }
    case XlsbRecordType::BrtTable: {
      DataTable table;
      if (!DecodeDataTableRecord(rec.payload, &table)) {
        return ResolveUnmodelledRecord(state, type, framed, framed_size);
      }
      // TABLECELL: the anchor's plain value record follows and carries its cached result.
      wb.sheet(sheet_index).mutable_data_tables().push_back(table);
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
    std::uint32_t* undecoded_formula_count, const RichErrorSet* rich_errors, std::vector<DeferredSpillAnchor>* deferred,
    std::uint64_t max_cell_bytes) {
  SheetDecodeState state(max_cell_bytes);
  const FeatureFormulaReadContext feature_ctx{sheet_names, name_table, sheet_ranges, external_books, sheet_index};
  const SheetDecodeInputs inputs{sst_entries, sst_phonetic, sst_phonetic_props, text_storage,           sheet_names,
                                 name_table,  sheet_ranges, external_books,     undecoded_formula_count};
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    const std::uint8_t* const framed = cursor.data;
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
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
        return std::move(feature.error());
      }
      if (feature.value()) {
        continue;
      }
    }
    // Every record resolves to a disposition; the result is consumed rather
    // than discarded so that no record can pass through unclassified.
    ASSIGN_OR_RETURN(auto disposition,
                     DispatchSheetRecord(rec, type, framed, framed_size, state, sheet_index, wb, inputs));
    if (disposition == RecordDisposition::kAccounted) {
      ++state.dropped_records;
    }
    // Grammar phase advances after the record has been dispatched, so the
    // marker itself belongs to the phase it closes rather than the one it
    // opens.
    if (type == XlsbRecordType::BrtEndSheetData) {
      state.in_tail = true;
    } else if (type == XlsbRecordType::BrtEndWsView) {
      if (state.in_extra_ws_view) {
        append_extra_view(wb.sheet(sheet_index).mutable_view(), state.extra_view, state.extra_workbook_view);
        state.in_extra_ws_view = false;
      }
      state.in_main_ws_view = false;
    } else if (type == XlsbRecordType::BrtEndMergeCells) {
      state.merges_seen = true;
    }
  }
  // A stream without BrtEndColInfos still ends with canonical column spans.
  normalize_column_spans(wb.sheet(sheet_index).mutable_layout());
  ApplyDecodedRowHeaders(wb.sheet(sheet_index), state.row_headers);
  wb.sheet(sheet_index).apply_blank_slot_rule();
  if (!state.protected_ranges.empty()) {
    write_protected_ranges(wb.sheet(sheet_index).mutable_raw_extensions(), state.protected_ranges);
  }
  if (rich_errors != nullptr && deferred != nullptr) {
    restore_rich_error_cells(wb, sheet_index, state.value_metadata, state.array_anchors, *rich_errors, deferred);
  }
  auto spills = RegisterArraySpills(wb, sheet_index, state.array_anchors);
  if (!spills) {
    return std::move(spills.error());
  }
  if (!state.tail.empty()) {
    std::vector<cf::ConditionalFormat>& formats = wb.sheet(sheet_index).mutable_conditional_formats();
    for (std::vector<std::uint8_t>* slot :
         {&state.tail.before_merges, &state.tail.after_merges_before_hyperlinks, &state.tail.after_hyperlinks}) {
      take_x14_rules(*slot, feature_ctx, formats);
      apply_x14_data_bar_overlays(ByteSpan{slot->data(), slot->size()}, formats);
      take_x14_validations(*slot, feature_ctx, wb.sheet(sheet_index).mutable_validations());
    }
    // Sparkline groups the records describe move into the sheet's extension XML.
    std::string ext_lst = wb.sheet(sheet_index).ext_lst_xml();
    take_sparkline_groups(state.tail, feature_ctx, ext_lst);
    wb.sheet(sheet_index).set_ext_lst_xml(std::move(ext_lst));
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
      // Another workbook's entry keeps its link and that book's sheet names,
      // so the writer can re-point it at the saved link.
      if (range.external_book != 0U && range.external_book <= external_books.size()) {
        const std::vector<std::string>& sheets = external_books[range.external_book - 1U].sheet_names;
        const auto in_link = [&sheets](std::int32_t itab) {
          return itab >= 0 && static_cast<std::size_t>(itab) < sheets.size();
        };
        if (sheetless || (in_link(range.itab_first) && in_link(range.itab_last))) {
          entry.external_book = range.external_book;
          if (!sheetless) {
            entry.first = sheets[static_cast<std::size_t>(range.itab_first)];
            entry.last = sheets[static_cast<std::size_t>(range.itab_last)];
          }
        }
      }
      state.tail.extern_sheets.push_back(std::move(entry));
    }
    wb.sheet(sheet_index).set_xlsb_tail(std::move(state.tail));
  }
  return state;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
