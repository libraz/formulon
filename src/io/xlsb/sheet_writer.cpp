//
// Implementation of the per-sheet XLSB stream emitter. See
// `io/xlsb/sheet_writer.h` for the contract.

#include "io/xlsb/sheet_writer.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "hyperlink_address.h"
#include "io/cf_layout.h"
#include "io/dv_overlay.h"
#include "io/dynamic_array_formula.h"
#include "io/ooxml_writer_cell.h"
#include "io/protection_xml.h"
#include "io/xlsb/cell_writer.h"
#include "io/xlsb/cf_records.h"
#include "io/xlsb/drawing_record.h"
#include "io/xlsb/dv_records.h"
#include "io/xlsb/feature_formula.h"
#include "io/xlsb/hyperlink_records.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_properties_records.h"
#include "io/xlsb/sheet_record_flags.h"
#include "io/xlsb/sparkline_records.h"
#include "io/xlsb/sst_writer.h"
#include "io/xlsb/table_records.h"
#include "io/xlsb/view_records.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_resolve.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// True when the stored cell at `(row, col)` writes no record: the .xlsx
/// writer's `CellIsEmitted` rule, so an explicit default is a `BrtCellBlank`
/// only where it overrides a row or column format.
bool IsEmptySlot(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell) {
  return !CellIsEmitted(sheet, row, col, cell);
}

/// Adds columns `[first, end)` to a row's spans, which `add_col` keeps per
/// 1,024-column segment: each segment the piece crosses gains its clipped ends.
template <typename AddCol>
void AddPieceToSpans(std::uint32_t first, std::uint64_t end, AddCol&& add_col) {
  for (std::uint64_t seg_first = first; seg_first < end;) {
    const std::uint64_t seg_end = std::min<std::uint64_t>(end, (seg_first / 1024U + 1U) * 1024U);
    add_col(static_cast<std::uint32_t>(seg_first));
    add_col(static_cast<std::uint32_t>(seg_end - 1U));
    seg_first = seg_end;
  }
}

/// First and last emitted column within one 1,024-column segment of a row.
struct ColSpan {
  std::uint32_t segment;
  std::uint32_t first;
  std::uint32_t last;
};

/// Emits a `BrtRowHdr` for `row`. [MS-XLSB] §2.4.770 layout:
///   * rw         : u32 (0-based row index)
///   * iStyleRef  : u32 (row style xf index, or 0)
///   * miyRw      : u16 (custom row height in twips; 0 = use default)
///   * flags1     : u8
///   * flags2     : u8  (outline / hidden / customHeight / fGhostDirty bits)
///   * fPhShow    : u8  (phonetic-guide default)
///   * ccolspan   : u32 (number of following BrtColSpan records)
///   * rgBrtColspan: repeated (colMic, colLast) u32 pairs
///
/// A row with cell records must describe every 1,024-column segment which
/// contains one.  Excel repairs a stream that claims `ccolspan == 0` while
/// following it with cells, even though our own permissive reader can decode
/// it.  An empty layout-only row legitimately has no spans.
void EmitRowHeader(std::vector<std::uint8_t>& dst, std::uint32_t row, const RowLayout* layout, bool default_hidden,
                   const std::vector<ColSpan>& spans) {
  std::vector<std::uint8_t> p;
  emit_u32(p, row);
  emit_u32(p, layout != nullptr && layout->has_style ? layout->style_xf : 0U);  // iStyleRef
  const bool has_height = layout != nullptr && (layout->has_height || layout->height > 0.0);
  const double twips = has_height ? std::round(layout->height * 20.0) : 0.0;
  emit_u16(p, static_cast<std::uint16_t>(std::clamp(twips, 0.0, 65535.0)));  // miyRw
  emit_u8(p, 0);                                                             // flags1
  std::uint8_t flags2 =
      layout == nullptr ? 0U : static_cast<std::uint8_t>(layout->outline_level & kRowOutlineLevelMask);
  // Every header states visibility, so a row without an override carries the sheet default.
  if (layout != nullptr ? layout->hidden : default_hidden) {
    flags2 |= kRowZeroHeight;
  }
  if (layout != nullptr && layout->collapsed) {
    flags2 |= kRowCollapsed;
  }
  if (layout != nullptr && layout->custom_height) {
    flags2 |= kRowUnsynced;
  }
  if (layout != nullptr && layout->has_style) {
    flags2 |= kRowGhostDirty;
  }
  emit_u8(p, flags2);
  // fPhShow is a distinct byte, not part of flags2.  It is required even
  // when the phonetic guide is disabled; otherwise ccolspan is misaligned.
  emit_u8(p, 0U);
  emit_u32(p, static_cast<std::uint32_t>(spans.size()));  // ccolspan
  for (const ColSpan& span : spans) {
    emit_u32(p, span.first);  // colMic
    emit_u32(p, span.last);   // colLast
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr), p);
}

void EmitColumnInfos(std::vector<std::uint8_t>& dst, const SheetLayout& layout) {
  if (layout.columns.empty()) {
    return;
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginColInfos), ByteSpan{});
  std::vector<std::uint8_t> p;
  for (const ColumnLayout& column : layout.columns) {
    if (column.first > column.last || column.last >= Sheet::kMaxCols) {
      continue;
    }
    p.clear();
    emit_u32(p, column.first);
    emit_u32(p, column.last);
    const bool has_width = HasExplicitColumnWidth(column);
    const double width256 = has_width ? std::floor(std::max(0.0, column.width) * 256.0) : 0.0;
    emit_u32(p, static_cast<std::uint32_t>(std::clamp(width256, 0.0, 65535.0)));
    // BrtColInfo has no style-presence bit; ixfe is mandatory. OOXML spans
    // without a style therefore canonicalize to style 0 on an XLSB reload.
    emit_u32(p, column.has_style ? column.style_xf : 0U);  // ixfe
    std::uint16_t flags = has_width ? kColUserSet : 0U;
    if (column.hidden) {
      flags |= kColHidden;
    }
    flags |= static_cast<std::uint16_t>((column.outline_level & kColOutlineLevelMask) << kColOutlineShift);
    if (column.collapsed) {
      flags |= kColCollapsed;
    }
    emit_u16(p, flags);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtColInfo), p);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndColInfos), ByteSpan{});
}

/// Emits the mandatory worksheet-view and default-formatting suffix. Omitting
/// these records leaves a technically decodable stream that desktop Excel
/// repairs before opening. The defaults match an Excel-created normal
/// worksheet: visible grid/headings and 100% zoom. The default metrics come
/// from the sheet's modelled `<sheetFormatPr>` values.
void EmitWorksheetViewsAndFormatting(std::vector<std::uint8_t>& dst, const SheetView& sheet_view,
                                     const SheetFormatDefaults& defaults, std::uint32_t workbook_views) {
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsViews), ByteSpan{});
  emit_ws_view(dst, sheet_view, 0U);
  emit_extra_ws_views(dst, sheet_view, workbook_views);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsViews), ByteSpan{});

  emit_ws_fmt_info(dst, defaults);
}

/// Emits BrtWsDim, the `sheet_dimension` box the .xlsx `<dimension>` shares,
/// or A1 (0..0 / 0..0) when there is nothing to cover. BrtWsDim is optional in
/// the grammar, but Excel writes it for every sheet, empty ones included, and
/// uses it to establish the used range before reading the cell table.
void EmitWorksheetDimensions(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  const SheetDimension box = sheet_dimension(sheet).value_or(SheetDimension{});
  std::vector<std::uint8_t> payload;
  emit_rfx(payload, MergeRange{box.first_row, box.first_col, box.last_row, box.last_col});
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsDim), payload);
}

void EmitMerges(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  if (sheet.merges().empty()) {
    return;
  }
  std::size_t valid_count = 0;
  for (const MergeRange& merge : sheet.merges()) {
    if (in_grid(merge)) {
      ++valid_count;
    }
  }
  if (valid_count == 0U) {
    return;
  }
  std::vector<std::uint8_t> count;
  emit_u32(count, static_cast<std::uint32_t>(valid_count));
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginMergeCells), count);
  std::vector<std::uint8_t> payload;
  for (const MergeRange& merge : sheet.merges()) {
    if (!in_grid(merge)) {
      continue;
    }
    payload.clear();
    emit_rfx(payload, merge);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtMergeCell), payload);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndMergeCells), ByteSpan{});
}

Expected<void, Error> EmitHyperlinks(std::vector<std::uint8_t>& dst, const Sheet& sheet,
                                     const std::vector<std::string>& relationship_ids,
                                     std::string_view book_directory) {
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& hyperlink = sheet.hyperlinks()[i];
    const std::string_view rid =
        i < relationship_ids.size() ? std::string_view(relationship_ids[i]) : std::string_view{};
    SavedHyperlink saved = saved_hyperlink(hyperlink, book_directory);
    if (saved.display != hyperlink.display) {
      Hyperlink written = hyperlink;
      written.display = std::move(saved.display);
      if (auto s = emit_hyperlink(dst, written, rid); !s) {
        return s;
      }
      continue;
    }
    if (auto s = emit_hyperlink(dst, hyperlink, rid); !s) {
      return s;
    }
  }
  return Expected<void, Error>::Ok();
}

/// Emits the retained records between the merge and hyperlink blocks
/// (BrtPhoneticInfo, and any block the reader could not decode) with the
/// model's conditional formats and validations in grammar order: phonetic
/// info, CF blocks, then the single DVals container.
Expected<void, Error> EmitFormattingAndValidation(std::vector<std::uint8_t>& body, const Sheet& sheet,
                                                  const std::vector<std::uint8_t>& retained,
                                                  const FeatureFormulaWriteContext& ctx,
                                                  const std::unordered_set<std::string>& x14_linked) {
  const auto raw_dvals = static_cast<std::ptrdiff_t>(find_record_offset(retained, kBrtBeginDVals));
  body.insert(body.end(), retained.begin(), retained.begin() + raw_dvals);
  for (const CfBlockView& view : cf_saved_layout(sheet.conditional_formats())) {
    if (auto s = emit_cf_block(body, view, ctx, x14_linked, sheet.name()); !s) {
      return s;
    }
  }
  const bool listed_dvals =
      std::any_of(sheet.validations().begin(), sheet.validations().end(),
                  [&](const DataValidation& v) { return !data_validation_in_x14(v, sheet.name(), &ctx.sheet_names); });
  if (raw_dvals != static_cast<std::ptrdiff_t>(retained.size()) && listed_dvals) {
    return make_error(FormulonErrorCode::kIoXlsbRetainedPartStale,
                      "retained XLSB data validations cannot be merged with model validations", kWriteXlsbContext);
  }
  if (auto s = emit_dv_block(body, sheet.validations(), ctx, sheet.name()); !s) {
    return s;
  }
  body.insert(body.end(), retained.begin() + raw_dvals, retained.end());
  return Expected<void, Error>::Ok();
}

}  // namespace

std::size_t unwritten_protected_range_count(const Sheet& sheet) {
  const std::vector<ProtectedRange> ranges = read_protected_ranges(sheet.raw_extensions());
  return static_cast<std::size_t>(std::count_if(ranges.begin(), ranges.end(),
                                                [](const ProtectedRange& r) { return !range_protection_writable(r); }));
}

std::vector<std::string> hyperlink_relationship_ids(const Sheet& sheet) {
  std::unordered_set<std::string> used;
  std::unordered_map<std::string, std::string> assigned_targets;
  used.reserve(sheet.unknown_relationships().size() + sheet.hyperlinks().size());
  assigned_targets.reserve(sheet.hyperlinks().size());
  for (const UnknownRelationship& relationship : sheet.unknown_relationships()) {
    if (!relationship.id.empty()) {
      used.insert(relationship.id);
    }
  }
  std::size_t next_id = 1U;
  std::vector<std::string> ids;
  ids.reserve(sheet.hyperlinks().size());
  auto fresh_id = [&]() {
    std::string id;
    do {
      id = "rId" + std::to_string(next_id++);
    } while (used.count(id) != 0U);
    used.insert(id);
    return id;
  };
  for (const Hyperlink& hyperlink : sheet.hyperlinks()) {
    if (hyperlink.target.empty()) {
      ids.emplace_back();
      continue;
    }
    if (!hyperlink.rid.empty()) {
      const auto assigned = assigned_targets.find(hyperlink.rid);
      if (assigned != assigned_targets.end() && assigned->second == hyperlink.target) {
        // Multiple BrtHLink records may legitimately share one relationship
        // when they point at the same external target. Preserve that source
        // sharing instead of minting a semantically equivalent duplicate.
        ids.push_back(hyperlink.rid);
        continue;
      }
      if (used.count(hyperlink.rid) == 0U) {
        ids.push_back(hyperlink.rid);
        used.insert(hyperlink.rid);
        assigned_targets.emplace(hyperlink.rid, hyperlink.target);
        continue;
      }
    }
    const std::string fresh = fresh_id();
    assigned_targets.emplace(fresh, hyperlink.target);
    ids.push_back(fresh);
  }
  return ids;
}

Expected<std::vector<std::uint8_t>, Error> emit_sheet(
    const Sheet& sheet, SstBuilder& sst, const std::vector<std::string>& sheet_names,
    const SheetRangeTable& sheet_ranges, const NameTable& name_table, const std::vector<std::string>& hyperlink_ids,
    std::uint32_t* downgraded_formula_count, std::uint32_t dynamic_array_ifmd, const NameShapes& name_shapes,
    const std::vector<std::string>& table_part_rel_ids, std::uint32_t workbook_views,
    std::uint32_t* deferred_sparkline_groups, std::string_view drawing_rel_id, bool owns_drawing,
    const RichErrorVmMap* rich_error_vm, std::string_view book_directory) {
  std::vector<std::uint8_t> body;

  // Frame: BrtBeginSheet | BrtBeginSheetData | ... | BrtEndSheetData |
  // BrtEndSheet.
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheet), ByteSpan{});
  emit_ws_prop(body, sheet);
  EmitWorksheetDimensions(body, sheet);
  // BrtWsFmtInfo always carries a row height, so an unset default is written
  // as Excel's standard row height.
  SheetFormatDefaults format_defaults = WithDerivedOutlineLevels(sheet.format_defaults(), sheet.layout());
  if (!format_defaults.has_default_row_height) {
    format_defaults.default_row_height = ooxml_defaults::kStandardRowHeightPt;
    format_defaults.has_default_row_height = true;
  }
  EmitWorksheetViewsAndFormatting(body, sheet.view(), format_defaults, workbook_views);
  EmitColumnInfos(body, sheet.layout());
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheetData), ByteSpan{});

  // Spilled dynamic-array formulas need special treatment: Excel stores
  // every cell of the spill footprint (the anchor and its phantoms) as a
  // PtgExp "shell" cell record, with the real tokens carried once in the
  // anchor's `BrtArrFmla`. Phantom cells are not stored in `Sheet::rows()`
  // (they live in the spill table), so gather them here keyed by their
  // absolute (row, col) before walking rows. Only genuine spills (a
  // footprint larger than one cell) are treated this way; a scalar result
  // stays an ordinary formula cell.
  struct PhantomShell {
    std::uint32_t anchor_row = 0;
    std::uint32_t anchor_col = 0;
    const Value* value = nullptr;  // into the sheet's spill region, or `blank`
    bool always_calculates = false;
  };
  const Value blank = Value::blank();
  std::map<std::uint32_t, std::map<std::uint32_t, PhantomShell>> phantoms;
  for (const auto& [anchor_row, cells] : sheet.rows()) {
    for (const auto& [col, cell] : cells) {
      if (cell.formula_text.empty()) {
        continue;
      }
      const SpillRegion* region = sheet.spill_region_at_anchor(anchor_row, col);
      if (region == nullptr || static_cast<std::uint64_t>(region->rows) * region->cols <= 1U) {
        continue;
      }
      const bool always = formula_cell_always_calculates(sheet, anchor_row, col, cell, name_shapes);
      for (std::uint32_t r = 0; r < region->rows; ++r) {
        for (std::uint32_t c = 0; c < region->cols; ++c) {
          if (r == 0 && c == 0) {
            continue;  // anchor: emitted as an anchor record, not a phantom
          }
          const std::size_t idx = static_cast<std::size_t>(r) * region->cols + c;
          const Value* value = idx < region->cells.size() ? &region->cells[idx] : &blank;
          phantoms[anchor_row + r][col + c] = PhantomShell{anchor_row, col, value, always};
        }
      }
    }
  }

  // Walk rows in ascending order. The Sheet's row map is unordered, so we
  // collect indices first (union of stored rows and phantom-only rows) and
  // sort. Workbooks rarely reach more than a few thousand populated rows,
  // so this is comfortably cheap.
  std::vector<std::uint32_t> row_indices;
  row_indices.reserve(sheet.rows().size() + phantoms.size() + sheet.layout().row_overrides.size());
  for (const auto& kv : sheet.rows()) {
    row_indices.push_back(kv.first);
  }
  for (const auto& kv : phantoms) {
    row_indices.push_back(kv.first);
  }
  for (const RowLayout& layout : sheet.layout().row_overrides) {
    if (layout.row < Sheet::kMaxRows) {
      row_indices.push_back(layout.row);
    }
  }
  sort_ascending(row_indices);
  row_indices.erase(std::unique(row_indices.begin(), row_indices.end()), row_indices.end());

  std::unordered_set<std::uint64_t> downgraded_array_anchors;
  auto anchor_key = [](std::uint32_t row, std::uint32_t col) { return (static_cast<std::uint64_t>(row) << 32U) | col; };

  std::vector<ColSpan> spans;
  std::vector<std::uint8_t> cell_payload;
  // Writes the `BrtValueMeta` of an error cell saved with a rich error, ahead of its cell record.
  const auto emit_value_meta = [&](std::uint32_t row, std::uint32_t col, const Value& shown) {
    if (rich_error_vm == nullptr || !shown.is_error()) {
      return;
    }
    const std::uint32_t vm = rich_error_vm_at(*rich_error_vm, row, col);
    if (vm == 0U) {
      return;
    }
    cell_payload.clear();
    emit_u32(cell_payload, vm);
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtValueMeta), cell_payload);
  };
  for (const std::uint32_t row : row_indices) {
    const RowLayout* row_layout = find_row_override(sheet.layout(), row);
    const auto stored_it = sheet.rows().find(row);
    const RowCells* row_cells = stored_it != sheet.rows().end() ? &stored_it->second : nullptr;
    const auto phantom_it = phantoms.find(row);
    const std::map<std::uint32_t, PhantomShell>* row_phantoms =
        phantom_it != phantoms.end() ? &phantom_it->second : nullptr;

    // Blank formatted columns, written as `BrtCellBlank` records, or the
    // row's explicit defaults written as its own xf-0 row format.
    const RowEmitPlan plan =
        row_cells != nullptr ? row_emit_plan(sheet, row, *row_cells, row_phantoms != nullptr) : RowEmitPlan();
    const std::vector<EmittedRun>& runs = plan.runs;
    RowLayout converted;
    if (plan.as_row_format) {
      if (row_layout != nullptr) {
        converted = *row_layout;
      }
      converted.row = row;
      converted.has_style = true;
      converted.style_xf = 0U;
      row_layout = &converted;
    }

    // A row with nothing to emit is skipped so we do not emit
    // a bare BrtRowHdr.
    bool any = !runs.empty();
    if (row_cells != nullptr && !any) {
      for (const auto& [col, cell] : *row_cells) {
        if (!IsEmptySlot(sheet, row, col, cell)) {
          any = true;
          break;
        }
      }
    }
    if (row_phantoms != nullptr) {
      any = any || !row_phantoms->empty();
    }
    if (!any && row_layout == nullptr) {
      continue;
    }

    // BrtRowHdr represents cells in 1,024-column segments.  Include only
    // cells actually emitted below; leading default-constructed slots must
    // not enlarge a span.
    spans.clear();
    // Spans are kept ordered by segment as they are added.
    auto add_to_span = [&spans](std::uint32_t col) {
      const std::uint32_t segment = col / 1024U;
      std::size_t at = 0;
      while (at < spans.size() && spans[at].segment < segment) {
        ++at;
      }
      if (at < spans.size() && spans[at].segment == segment) {
        spans[at].first = std::min(spans[at].first, col);
        spans[at].last = std::max(spans[at].last, col);
        return;
      }
      spans.push_back(ColSpan{segment, col, col});
      for (std::size_t i = spans.size() - 1U; i > at; --i) {
        std::swap(spans[i], spans[i - 1U]);
      }
    };
    if (row_cells != nullptr) {
      for (const auto& [col, cell] : *row_cells) {
        if (!IsEmptySlot(sheet, row, col, cell)) {
          add_to_span(col);
        }
      }
    }
    for (const EmittedRun& piece : runs) {
      AddPieceToSpans(piece.first, piece.end, add_to_span);
    }
    if (row_phantoms != nullptr) {
      for (const auto& [col, phantom] : *row_phantoms) {
        (void)phantom;
        add_to_span(col);
      }
    }

    EmitRowHeader(body, row, row_layout, sheet.format_defaults().zero_height, spans);
    // Ascending three-way merge of stored entries, phantoms and run columns; a
    // column in both a stored entry and a phantom takes the entry path unless
    // it is style-only, a phantom wins over a run column, and a stored entry
    // and a run never share a column.
    RowCells::const_iterator entry_it = row_cells != nullptr ? row_cells->begin() : RowCells::const_iterator();
    const RowCells::const_iterator entry_end = row_cells != nullptr ? row_cells->end() : RowCells::const_iterator();
    auto ph = row_phantoms != nullptr ? row_phantoms->begin() : std::map<std::uint32_t, PhantomShell>::const_iterator();
    const auto ph_end =
        row_phantoms != nullptr ? row_phantoms->end() : std::map<std::uint32_t, PhantomShell>::const_iterator();
    auto run = runs.begin();
    std::uint64_t run_col = run != runs.end() ? run->first : UINT64_MAX;
    const auto advance_run_past = [&](std::uint64_t col) {
      while (run != runs.end() && run_col <= col) {
        run_col = col + 1U;
        if (run_col >= run->end) {
          ++run;
          run_col = run != runs.end() ? run->first : UINT64_MAX;
        }
      }
    };
    while (entry_it != entry_end || ph != ph_end || run != runs.end()) {
      const std::uint32_t entry_col = entry_it != entry_end ? (*entry_it).col : UINT32_MAX;
      const std::uint32_t phantom_col = ph != ph_end ? ph->first : UINT32_MAX;
      if (run_col < entry_col && run_col < phantom_col) {
        const std::uint64_t stop = std::min<std::uint64_t>({run->end, entry_col, phantom_col});
        for (std::uint64_t run_at = run_col; run_at < stop; ++run_at) {
          emit_blank(body, cell_payload, static_cast<std::uint32_t>(run_at), run->xf);
        }
        advance_run_past(stop - 1U);
        continue;
      }
      const std::uint32_t col = std::min(entry_col, phantom_col);
      advance_run_past(col);
      const Cell* cell = nullptr;
      if (entry_col == col) {
        cell = &(*entry_it).cell;
        ++entry_it;
      }
      const bool phantom = phantom_col == col;
      const auto cur_ph = ph;
      if (phantom) {
        ++ph;
      }
      // A spilled cell carrying only its own style is still the spill's.
      if (cell != nullptr && !IsEmptySlot(sheet, row, col, *cell) && !(phantom && cell->formula_text.empty())) {
        const bool always = formula_cell_always_calculates(sheet, row, col, *cell, name_shapes);
        const SpillRegion* region = cell->formula_text.empty() ? nullptr : sheet.spill_region_at_anchor(row, col);
        // A dynamic-array formula keeps that form (its `BrtCellMeta` naming
        // the XLDAPR entry) even when it does not spill; a legacy CSE block
        // keeps the array form alone, over its own extent.
        const bool dynamic = dynamic_array_ifmd != 0U && is_dynamic_array_formula(*cell);
        const bool cse = cell->array_block_rows != 0U;
        if (region != nullptr || dynamic || cse) {
          const std::uint32_t rows = cse ? cell->array_block_rows : region != nullptr ? region->rows : 1U;
          const std::uint32_t cols = cse ? cell->array_block_cols : region != nullptr ? region->cols : 1U;
          const std::uint32_t last_row = row + rows - 1U;
          const std::uint32_t last_col = col + cols - 1U;
          const Value& anchor_value =
              region != nullptr && !region->cells.empty() ? region->cells.front() : cell->cached_value;
          bool downgraded_to_literal = false;
          if (dynamic) {
            cell_payload.clear();
            emit_u32(cell_payload, dynamic_array_ifmd);  // XLDAPR dynamic-array metadata entry
            emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta), cell_payload);
          }
          emit_value_meta(row, col, anchor_value);
          if (auto r = emit_array_anchor(body, cell_payload, *cell, anchor_value, always, col, row, last_row, last_col,
                                         sheet_names, sheet_ranges, name_table, sst, downgraded_formula_count,
                                         &downgraded_to_literal, name_shapes);
              !r) {
            return std::move(r.error());
          }
          if (downgraded_to_literal) {
            downgraded_array_anchors.insert(anchor_key(row, col));
          }
        } else if (const DataTable* table = sheet.data_table_anchored_at(row, col);
                   table != nullptr && cell->formula_text.empty()) {
          emit_data_table_anchor(body, cell_payload, *cell, col, *table, sst);
        } else {
          emit_value_meta(row, col, cell->cached_value);
          if (auto r = emit_cell(body, cell_payload, *cell, row, col, always, sst, sheet_names, sheet_ranges,
                                 name_table, downgraded_formula_count, name_shapes);
              !r) {
            return std::move(r.error());
          }
        }
        continue;
      }
      if (phantom) {
        if (downgraded_array_anchors.count(anchor_key(cur_ph->second.anchor_row, cur_ph->second.anchor_col)) != 0U) {
          continue;
        }
        const std::uint32_t xf_index = select_effective_xf(sheet, row, col).xf_index;
        emit_value_meta(row, col, *cur_ph->second.value);
        emit_array_phantom(body, cell_payload, col, xf_index, *cur_ph->second.value, cur_ph->second.anchor_row,
                           cur_ph->second.anchor_col, cur_ph->second.always_calculates);
      }
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData), ByteSpan{});
  // Retained tail records bracket the merged-cell block so a sheet read from
  // an .xlsb keeps its conditional formatting, data validation, hyperlinks,
  // auto-filter and print setup in their original stream positions. The
  // buffers hold already-framed records, so this is a byte append.
  // The x14 data-bar rules follow the model: retained ones are brought in
  // line with it, missing ones are added.
  XlsbSheetTail tail = sheet.xlsb_tail();
  std::unordered_set<std::string> x14_linked;
  for (std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    reconcile_x14_data_bars(*slot, sheet.conditional_formats(), x14_linked);
  }
  if (auto added = add_x14_data_bars(tail.after_hyperlinks, sheet.conditional_formats(), x14_linked); !added) {
    return std::move(added.error());
  }
  // Standalone x14 CF rules (custom icon sets, rules naming another sheet)
  // and sparkline groups are model-owned future records; retained rules are
  // rewritten in place, missing ones are added.
  const FeatureFormulaWriteContext frt_ctx{sheet_names, sheet_ranges, name_table};
  std::unordered_set<std::string> rules14_linked;
  for (std::vector<std::uint8_t>* slot :
       {&tail.before_merges, &tail.after_merges_before_hyperlinks, &tail.after_hyperlinks}) {
    if (auto s = reconcile_x14_rules(*slot, sheet.conditional_formats(), frt_ctx, sheet.name(), tail.extern_sheets,
                                     rules14_linked);
        !s) {
      return std::move(s.error());
    }
  }
  const auto own_index =
      static_cast<std::size_t>(std::find(sheet_names.begin(), sheet_names.end(), sheet.name()) - sheet_names.begin());
  if (auto rules14 = add_x14_rules(tail.after_hyperlinks, sheet.conditional_formats(), frt_ctx, sheet.name(), own_index,
                                   rules14_linked);
      !rules14) {
    return std::move(rules14.error());
  }
  // Validations naming another sheet are model-owned BrtDVal14 records,
  // rewritten in place like the x14 CF rules.
  if (auto dvals14 = place_x14_validations(tail, sheet.validations(), frt_ctx, sheet.name()); !dvals14) {
    return std::move(dvals14.error());
  }
  if (auto sparklines = add_sparkline_groups(tail.after_hyperlinks, sheet.ext_lst_xml(), frt_ctx, sheet.name(),
                                             deferred_sparkline_groups);
      !sparklines) {
    return std::move(sparklines.error());
  }
  // A model-owned drawing replaces any retained BrtDrawing record.
  if (owns_drawing) {
    strip_drawing_records(tail.after_hyperlinks);
  }
  if (!drawing_rel_id.empty()) {
    splice_drawing(tail.after_hyperlinks, drawing_rel_id);
  }
  // Table parts are model-owned; their BrtListPart records join the tail.
  splice_list_parts(tail.after_hyperlinks, table_part_rel_ids);
  // Sheet and range protection, conditional formatting and data validation
  // are emitted from the model. Protection leads the tail: the one record the grammar
  // puts ahead of it (BrtSheetCalcProp) is one Excel does not persist.
  if (auto protection = emit_sheet_protection(body, sheet.protection()); !protection) {
    return std::move(protection.error());
  }
  if (auto ranges = emit_range_protections(body, read_protected_ranges(sheet.raw_extensions())); !ranges) {
    return std::move(ranges.error());
  }
  body.insert(body.end(), tail.before_merges.begin(), tail.before_merges.end());
  EmitMerges(body, sheet);
  if (auto features =
          EmitFormattingAndValidation(body, sheet, tail.after_merges_before_hyperlinks, frt_ctx, x14_linked);
      !features) {
    return std::move(features.error());
  }
  if (auto hyperlinks = EmitHyperlinks(body, sheet, hyperlink_ids, book_directory); !hyperlinks) {
    return std::move(hyperlinks.error());
  }
  body.insert(body.end(), tail.after_hyperlinks.begin(), tail.after_hyperlinks.end());
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
