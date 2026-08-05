//
// Implementation of the per-sheet XLSB stream emitter. See
// `io/xlsb/sheet_writer.h` for the contract.

#include "io/xlsb/sheet_writer.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <map>
#include <string>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "io/xlsb/cell_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sst_writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Returns `true` when a `Cell` is "empty" in the sense that nothing
/// useful would be carried by emitting a `BrtCellBlank` for it. Used
/// to skip implicitly-default-constructed columns produced by sheet
/// row growth (see `Sheet::set_cell_value` docs).
bool IsEmptySlot(const Cell& cell) {
  return cell.formula_text.empty() && cell.cached_value.is_blank();
}

/// Emits a `BrtRowHdr` for `row`. [MS-XLSB] §2.4.770 layout:
///   * rw         : u32 (0-based row index)
///   * iStyleRef  : u32 (0)
///   * miyRw      : u16 (custom row height in twips; 0 = use default)
///   * flags1     : u8
///   * flags2     : u8  (outline / hidden / customHeight bits)
///   * fPhShow    : u8  (phonetic-guide default)
///   * ccolspan   : u32 (number of following BrtColSpan records)
///   * rgBrtColspan: repeated (colMic, colLast) u32 pairs
///
/// A row with cell records must describe every 1,024-column segment which
/// contains one.  Excel repairs a stream that claims `ccolspan == 0` while
/// following it with cells, even though our own permissive reader can decode
/// it.  An empty layout-only row legitimately has no spans.
void EmitRowHeader(std::vector<std::uint8_t>& dst, std::uint32_t row, const RowLayout* layout,
                   const std::map<std::uint32_t, std::pair<std::uint32_t, std::uint32_t>>& spans) {
  std::vector<std::uint8_t> p;
  emit_u32(p, row);
  emit_u32(p, 0);  // iStyleRef
  const bool has_height = layout != nullptr && (layout->has_height || layout->height > 0.0);
  const double twips = has_height ? std::round(layout->height * 20.0) : 0.0;
  emit_u16(p, static_cast<std::uint16_t>(std::clamp(twips, 0.0, 65535.0)));  // miyRw
  emit_u8(p, 0);                                                             // flags1
  std::uint8_t flags2 = layout == nullptr ? 0U : static_cast<std::uint8_t>(layout->outline_level & 0x07U);
  if (layout != nullptr && layout->hidden) {
    flags2 |= 0x10U;  // fDyZero
  }
  if (has_height) {
    flags2 |= 0x20U;  // fUnsynced
  }
  emit_u8(p, flags2);
  // fPhShow is a distinct byte, not part of flags2.  It is required even
  // when the phonetic guide is disabled; otherwise ccolspan is misaligned.
  emit_u8(p, 0U);
  emit_u32(p, static_cast<std::uint32_t>(spans.size()));  // ccolspan
  for (const auto& [segment, span] : spans) {
    (void)segment;
    emit_u32(p, span.first);   // colMic
    emit_u32(p, span.second);  // colLast
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr), p);
}

void EmitColumnInfos(std::vector<std::uint8_t>& dst, const SheetLayout& layout) {
  if (layout.columns.empty()) {
    return;
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginColInfos), ByteSpan{});
  for (const ColumnLayout& column : layout.columns) {
    if (column.first > column.last || column.last >= Sheet::kMaxCols) {
      continue;
    }
    std::vector<std::uint8_t> p;
    emit_u32(p, column.first);
    emit_u32(p, column.last);
    const double width256 = std::floor(std::max(0.0, column.width) * 256.0);
    emit_u32(p, static_cast<std::uint32_t>(std::clamp(width256, 0.0, 65535.0)));
    emit_u32(p, 0);                 // ixfe
    std::uint16_t flags = 0x0002U;  // fUserSet
    if (column.hidden) {
      flags |= 0x0001U;
    }
    flags |= static_cast<std::uint16_t>((column.outline_level & 0x07U) << 8U);
    emit_u16(p, flags);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtColInfo), p);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndColInfos), ByteSpan{});
}

/// Emits the worksheet-properties record which starts the mandatory worksheet
/// prefix. `BrtWsDim` must follow this record before the view collection.
void EmitWorksheetProperties(std::vector<std::uint8_t>& dst) {
  std::vector<std::uint8_t> properties;
  emit_u16(properties, 0x04C9U);  // page breaks, publish, outline defaults
  emit_u8(properties, 0x02U);     // evaluate conditional formatting
  // BrtColor: automatic tab color (type=automatic, index=0x40).
  emit_u8(properties, 0U);
  emit_u8(properties, 0x40U);
  emit_u16(properties, 0U);
  emit_u8(properties, 0U);
  emit_u8(properties, 0U);
  emit_u8(properties, 0U);
  emit_u8(properties, 0U);
  emit_u32(properties, 0xFFFFFFFFU);  // rwSync: unused
  emit_u32(properties, 0xFFFFFFFFU);  // colSync: unused
  emit_u32(properties, 0U);           // empty CodeName
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsProp), properties);
}

/// Emits the mandatory worksheet-view and default-formatting suffix. Omitting
/// these records leaves a technically decodable stream that desktop Excel
/// repairs before opening. The defaults match an Excel-created normal
/// worksheet: visible grid/headings, 100% zoom, 10-character default column
/// width, and 20-point default row height.
void EmitWorksheetViewsAndFormatting(std::vector<std::uint8_t>& dst, const SheetView& sheet_view) {
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsViews), ByteSpan{});
  std::vector<std::uint8_t> view;
  std::uint16_t flags = 0x0380U;  // ruler, outline symbols, default gridline color
  if (sheet_view.show_grid_lines) {
    flags |= 0x0004U;
  }
  if (sheet_view.show_row_col_headers) {
    flags |= 0x0008U;
  }
  if (sheet_view.show_zeros) {
    flags |= 0x0010U;
  }
  if (sheet_view.right_to_left) {
    flags |= 0x0020U;
  }
  if (sheet_view.tab_selected) {
    flags |= 0x0040U;
  }
  emit_u16(view, flags);
  emit_u32(view, 0U);    // XLVNORMAL
  emit_u32(view, 0U);    // rwTop
  emit_u32(view, 0U);    // colLeft
  emit_u8(view, 0x40U);  // default gridline color
  emit_u8(view, 0U);
  emit_u16(view, 0U);
  const std::uint32_t zoom = std::clamp(sheet_view.zoom_scale, 10U, 400U);
  emit_u16(view, static_cast<std::uint16_t>(zoom));
  emit_u16(view, 0U);
  emit_u16(view, 0U);
  emit_u16(view, 0U);
  emit_u32(view, 0U);  // workbook view index
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginWsView), view);
  if (sheet_view.freeze_rows != 0U || sheet_view.freeze_cols != 0U) {
    // BrtPane uses the frozen row and column counts in the two Xnum fields,
    // and identifies the bottom-right pane as the active pane. fFrozenNoSplit
    // represents Excel's ordinary "Freeze Panes" state.
    std::vector<std::uint8_t> pane;
    emit_double(pane, static_cast<double>(sheet_view.freeze_rows));
    emit_double(pane, static_cast<double>(sheet_view.freeze_cols));
    emit_u32(pane, sheet_view.freeze_rows);
    emit_u32(pane, sheet_view.freeze_cols);
    emit_u32(pane, 0U);  // PNNBOTRIGHT
    emit_u8(pane, 0x02U);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtPane), pane);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsView), ByteSpan{});
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndWsViews), ByteSpan{});

  std::vector<std::uint8_t> formatting;
  emit_u32(formatting, 0xFFFFFFFFU);  // default column width is character-count based
  emit_u16(formatting, 10U);          // default column width
  emit_u16(formatting, 400U);         // default row height (twips)
  emit_u32(formatting, 0U);           // default visibility/borders/outline levels
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsFmtInfo), formatting);
}

/// Emits the bounding RfX of actual emitted cells.  Although BrtWsDim is
/// optional in the grammar, Excel itself writes it for every non-empty sheet
/// and uses it to establish the worksheet's used range before reading the
/// cell table.
void EmitWorksheetDimensions(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  std::uint32_t first_row = Sheet::kMaxRows;
  std::uint32_t last_row = 0U;
  std::uint32_t first_col = Sheet::kMaxCols;
  std::uint32_t last_col = 0U;
  bool any = false;
  auto include = [&](std::uint32_t row, std::uint32_t col) {
    first_row = std::min(first_row, row);
    last_row = std::max(last_row, row);
    first_col = std::min(first_col, col);
    last_col = std::max(last_col, col);
    any = true;
  };

  for (const auto& [row, cells] : sheet.rows()) {
    for (std::uint32_t col = 0; col < cells.size(); ++col) {
      if (!IsEmptySlot(cells[col])) {
        include(row, col);
      }
    }
  }
  if (!any) {
    return;
  }
  std::vector<std::uint8_t> payload;
  emit_u32(payload, first_row);
  emit_u32(payload, last_row);
  emit_u32(payload, first_col);
  emit_u32(payload, last_col);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtWsDim), payload);
}

void EmitMerges(std::vector<std::uint8_t>& dst, const Sheet& sheet) {
  if (sheet.merges().empty()) {
    return;
  }
  std::size_t valid_count = 0;
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row <= merge.last_row && merge.first_col <= merge.last_col && merge.last_row < Sheet::kMaxRows &&
        merge.last_col < Sheet::kMaxCols) {
      ++valid_count;
    }
  }
  if (valid_count == 0U) {
    return;
  }
  std::vector<std::uint8_t> count;
  emit_u32(count, static_cast<std::uint32_t>(valid_count));
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginMergeCells), count);
  for (const MergeRange& merge : sheet.merges()) {
    if (merge.first_row > merge.last_row || merge.first_col > merge.last_col || merge.last_row >= Sheet::kMaxRows ||
        merge.last_col >= Sheet::kMaxCols) {
      continue;
    }
    std::vector<std::uint8_t> payload;
    emit_u32(payload, merge.first_row);
    emit_u32(payload, merge.last_row);
    emit_u32(payload, merge.first_col);
    emit_u32(payload, merge.last_col);
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtMergeCell), payload);
  }
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtEndMergeCells), ByteSpan{});
}

}  // namespace

Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst,
                                                      const std::vector<std::string>& sheet_names,
                                                      const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                                      std::uint32_t* downgraded_formula_count,
                                                      bool emit_dynamic_metadata) {
  std::vector<std::uint8_t> body;

  // Frame: BrtBeginSheet | BrtBeginSheetData | ... | BrtEndSheetData |
  // BrtEndSheet.
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheet), ByteSpan{});
  EmitWorksheetProperties(body);
  EmitWorksheetDimensions(body, sheet);
  EmitWorksheetViewsAndFormatting(body, sheet.view());
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
    Value value = Value::blank();
  };
  std::map<std::uint32_t, std::map<std::uint32_t, PhantomShell>> phantoms;
  for (const auto& [anchor_row, cells] : sheet.rows()) {
    for (std::uint32_t col = 0; col < cells.size(); ++col) {
      const Cell& cell = cells[col];
      if (cell.formula_text.empty()) {
        continue;
      }
      const SpillRegion* region = sheet.spill_region_at_anchor(anchor_row, col);
      if (region == nullptr || static_cast<std::uint64_t>(region->rows) * region->cols <= 1U) {
        continue;
      }
      for (std::uint32_t r = 0; r < region->rows; ++r) {
        for (std::uint32_t c = 0; c < region->cols; ++c) {
          if (r == 0 && c == 0) {
            continue;  // anchor: emitted as an anchor record, not a phantom
          }
          const std::size_t idx = static_cast<std::size_t>(r) * region->cols + c;
          Value value = idx < region->cells.size() ? region->cells[idx] : Value::blank();
          phantoms[anchor_row + r][col + c] = PhantomShell{anchor_row, col, std::move(value)};
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
  std::sort(row_indices.begin(), row_indices.end());
  row_indices.erase(std::unique(row_indices.begin(), row_indices.end()), row_indices.end());

  std::unordered_set<std::uint64_t> downgraded_array_anchors;
  auto anchor_key = [](std::uint32_t row, std::uint32_t col) { return (static_cast<std::uint64_t>(row) << 32U) | col; };

  for (const std::uint32_t row : row_indices) {
    const RowLayout* row_layout = nullptr;
    for (const RowLayout& candidate : sheet.layout().row_overrides) {
      if (candidate.row == row) {
        row_layout = &candidate;
        break;
      }
    }
    const auto stored_it = sheet.rows().find(row);
    const std::vector<Cell>* row_cells = stored_it != sheet.rows().end() ? &stored_it->second : nullptr;
    const auto phantom_it = phantoms.find(row);
    const std::map<std::uint32_t, PhantomShell>* row_phantoms =
        phantom_it != phantoms.end() ? &phantom_it->second : nullptr;

    // Highest column carrying anything in this row (stored non-empty cell
    // or phantom). A row with nothing to emit is skipped so we do not emit
    // a bare BrtRowHdr.
    std::uint32_t max_col = 0;
    bool any = false;
    if (row_cells != nullptr) {
      for (std::uint32_t col = 0; col < row_cells->size(); ++col) {
        if (!IsEmptySlot((*row_cells)[col])) {
          max_col = std::max(max_col, col);
          any = true;
        }
      }
    }
    if (row_phantoms != nullptr) {
      for (const auto& kv : *row_phantoms) {
        max_col = std::max(max_col, kv.first);
        any = true;
      }
    }
    if (!any && row_layout == nullptr) {
      continue;
    }

    // BrtRowHdr represents cells in 1,024-column segments.  Include only
    // cells actually emitted below; leading default-constructed slots must
    // not enlarge a span.
    std::map<std::uint32_t, std::pair<std::uint32_t, std::uint32_t>> spans;
    auto add_to_span = [&spans](std::uint32_t col) {
      const std::uint32_t segment = col / 1024U;
      const auto [it, inserted] = spans.emplace(segment, std::make_pair(col, col));
      if (!inserted) {
        it->second.first = std::min(it->second.first, col);
        it->second.second = std::max(it->second.second, col);
      }
    };
    if (row_cells != nullptr) {
      for (std::uint32_t col = 0; col < row_cells->size(); ++col) {
        if (!IsEmptySlot((*row_cells)[col])) {
          add_to_span(col);
        }
      }
    }
    if (row_phantoms != nullptr) {
      for (const auto& [col, phantom] : *row_phantoms) {
        (void)phantom;
        add_to_span(col);
      }
    }

    EmitRowHeader(body, row, row_layout, spans);
    for (std::uint32_t col = 0; col <= max_col; ++col) {
      const Cell* cell = row_cells != nullptr && col < row_cells->size() ? &(*row_cells)[col] : nullptr;
      if (cell != nullptr && !IsEmptySlot(*cell)) {
        const SpillRegion* region = cell->formula_text.empty() ? nullptr : sheet.spill_region_at_anchor(row, col);
        if (region != nullptr) {
          const std::uint32_t last_row = row + region->rows - 1U;
          const std::uint32_t last_col = col + region->cols - 1U;
          const Value& anchor_value = !region->cells.empty() ? region->cells.front() : cell->cached_value;
          bool downgraded_to_literal = false;
          if (emit_dynamic_metadata) {
            std::vector<std::uint8_t> metadata_index;
            emit_u32(metadata_index, 1U);  // XLDAPR dynamic-array metadata entry
            emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta), metadata_index);
          }
          if (auto r =
                  emit_array_anchor(body, *cell, anchor_value, col, row, last_row, last_col, sheet_names, sheet_ranges,
                                    name_table, sst, downgraded_formula_count, &downgraded_to_literal);
              !r) {
            return r.error();
          }
          if (downgraded_to_literal) {
            downgraded_array_anchors.insert(anchor_key(row, col));
          }
        } else if (auto r = emit_cell(body, *cell, row, col, sst, sheet_names, sheet_ranges, name_table,
                                      downgraded_formula_count);
                   !r) {
          return r.error();
        }
        continue;
      }
      if (row_phantoms != nullptr) {
        const auto ph = row_phantoms->find(col);
        if (ph != row_phantoms->end()) {
          if (downgraded_array_anchors.count(anchor_key(ph->second.anchor_row, ph->second.anchor_col)) != 0U) {
            continue;
          }
          const std::uint32_t xf_index = cell != nullptr ? cell->xf_index : 0U;
          emit_array_phantom(body, col, xf_index, ph->second.value, ph->second.anchor_row, ph->second.anchor_col);
        }
      }
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData), ByteSpan{});
  EmitMerges(body, sheet);
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
