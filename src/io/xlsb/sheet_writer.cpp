// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the per-sheet XLSB stream emitter. See
// `io/xlsb/sheet_writer.h` for the contract.

#include "io/xlsb/sheet_writer.h"

#include <algorithm>
#include <cstdint>
#include <map>
#include <string>
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

/// Emits a `BrtRowHdr` for `row`. [MS-XLSB] §2.4.660 layout:
///   * rw         : u32 (0-based row index)
///   * iStyleRef  : u32 (0)
///   * miyRw      : u16 (custom row height in twips; 0 = use default)
///   * flags1     : u8  (collapsed / hidden / customHeight bits)
///   * flags2     : u8  (custom format / phonetic guide bits)
///   * colMic     : u32 (first non-blank column hint; 0 is fine)
///   * colMac     : u32 (first all-blank column hint past data; 0 is fine)
///
/// The reader only consumes the leading `rw` field and skips over the
/// rest, so we only need the wire layout to match the size the reader
/// expects past the row index. Excel itself accepts a minimal
/// `(rw, 0, 0, 0, 0, 0, 0)` payload.
void EmitRowHeader(std::vector<std::uint8_t>& dst, std::uint32_t row) {
  std::vector<std::uint8_t> p;
  emit_u32(p, row);
  emit_u32(p, 0);  // iStyleRef
  emit_u16(p, 0);  // miyRw
  emit_u8(p, 0);   // flags1
  emit_u8(p, 0);   // flags2
  emit_u32(p, 0);  // colMic
  emit_u32(p, 0);  // colMac
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr), p);
}

}  // namespace

Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst,
                                                      const std::vector<std::string>& sheet_names,
                                                      const SheetRangeTable& sheet_ranges,
                                                      const NameTable& name_table) {
  std::vector<std::uint8_t> body;

  // Frame: BrtBeginSheet | BrtBeginSheetData | ... | BrtEndSheetData |
  // BrtEndSheet.
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheet), ByteSpan{});
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
  row_indices.reserve(sheet.rows().size() + phantoms.size());
  for (const auto& kv : sheet.rows()) {
    row_indices.push_back(kv.first);
  }
  for (const auto& kv : phantoms) {
    row_indices.push_back(kv.first);
  }
  std::sort(row_indices.begin(), row_indices.end());
  row_indices.erase(std::unique(row_indices.begin(), row_indices.end()), row_indices.end());

  for (const std::uint32_t row : row_indices) {
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
    if (!any) {
      continue;
    }

    EmitRowHeader(body, row);
    for (std::uint32_t col = 0; col <= max_col; ++col) {
      const Cell* cell = row_cells != nullptr && col < row_cells->size() ? &(*row_cells)[col] : nullptr;
      if (cell != nullptr && !IsEmptySlot(*cell)) {
        const SpillRegion* region = cell->formula_text.empty() ? nullptr : sheet.spill_region_at_anchor(row, col);
        if (region != nullptr && static_cast<std::uint64_t>(region->rows) * region->cols > 1U) {
          const std::uint32_t last_row = row + region->rows - 1U;
          const std::uint32_t last_col = col + region->cols - 1U;
          const Value& anchor_value = !region->cells.empty() ? region->cells.front() : cell->cached_value;
          if (auto r = emit_array_anchor(body, *cell, anchor_value, col, row, last_row, last_col, sheet_names,
                                         sheet_ranges, name_table);
              !r) {
            return r.error();
          }
        } else if (auto r = emit_cell(body, *cell, row, col, sst, sheet_names, sheet_ranges, name_table); !r) {
          return r.error();
        }
        continue;
      }
      if (row_phantoms != nullptr) {
        const auto ph = row_phantoms->find(col);
        if (ph != row_phantoms->end()) {
          const std::uint32_t xf_index = cell != nullptr ? cell->xf_index : 0U;
          emit_array_phantom(body, col, xf_index, ph->second.value, ph->second.anchor_row, ph->second.anchor_col);
        }
      }
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData), ByteSpan{});
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
