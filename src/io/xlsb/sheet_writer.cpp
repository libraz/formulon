// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the per-sheet XLSB stream emitter. See
// `io/xlsb/sheet_writer.h` for the contract.

#include "io/xlsb/sheet_writer.h"

#include <algorithm>
#include <cstdint>
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
/// Bundle 4.1's reader only consumes the leading `rw` field and skips
/// over the rest, so we only need the wire layout to match the size
/// the reader expects past the row index. Excel itself accepts a
/// minimal `(rw, 0, 0, 0, 0, 0, 0)` payload.
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

Expected<std::vector<std::uint8_t>, Error> emit_sheet(const Sheet& sheet, SstBuilder& sst) {
  std::vector<std::uint8_t> body;

  // Frame: BrtBeginSheet | BrtBeginSheetData | ... | BrtEndSheetData |
  // BrtEndSheet.
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheet), ByteSpan{});
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheetData), ByteSpan{});

  // Walk rows in ascending order. The Sheet's row map is unordered, so
  // we collect indices first and sort. Workbooks rarely reach more than
  // a few thousand populated rows, so this is comfortably cheap.
  std::vector<std::uint32_t> row_indices;
  row_indices.reserve(sheet.rows().size());
  for (const auto& kv : sheet.rows()) {
    row_indices.push_back(kv.first);
  }
  std::sort(row_indices.begin(), row_indices.end());

  for (const std::uint32_t row : row_indices) {
    const auto it = sheet.rows().find(row);
    if (it == sheet.rows().end()) {
      continue;
    }
    const std::vector<Cell>& row_cells = it->second;
    EmitRowHeader(body, row);
    for (std::uint32_t col = 0; col < row_cells.size(); ++col) {
      const Cell& cell = row_cells[col];
      if (IsEmptySlot(cell)) {
        // Skip implicitly-default-constructed columns: writing a
        // `BrtCellBlank` would round-trip through the reader as an
        // explicit blank cell, which inflates the populated-cell
        // count for no observable gain.
        continue;
      }
      emit_cell(body, cell, row, col, sst);
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData), ByteSpan{});
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheet), ByteSpan{});
  return body;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
