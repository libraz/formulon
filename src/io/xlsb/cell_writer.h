// Copyright 2026 libraz. Licensed under the MIT License.
//
// MS-XLSB per-cell record emitter. The sheet writer walks each
// populated row and calls `emit_cell` for every column slot, which
// dispatches on the cell's `Value::kind()` (and on the presence of a
// formula text) to produce the appropriate `BrtCell*` /
// `BrtFmla*` record bytes.
//
// The Bundle 4.2 dispatcher covers literal cells (number, boolean,
// text, error, blank) and the round-trip path for formula cells whose
// `formula_text` is one of Bundle 4.1's reader stubs
// (`__FORMULON_XLSB_PTG__(...)`). Cells whose formula was authored
// fresh in-engine (no captured Ptg bytes) currently emit the cached
// value as a literal record and surface a `xlsb.writer.formula_lost`
// structured-log warning — the full AST→Ptg encoder lands in a later
// bundle.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.6 (XLSB record stream layout)
//   * backup/plans/21-xlsb-ptg.md (formula round-trip plan)
//   * [MS-XLSB] §2.4.x (per-cell record types)

#ifndef FORMULON_IO_XLSB_CELL_WRITER_H_
#define FORMULON_IO_XLSB_CELL_WRITER_H_

#include <cstdint>
#include <vector>

#include "cell.h"
#include "io/xlsb/sst_writer.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Appends the byte-level XLSB record(s) for `cell` at column `col`
/// into `dst`. The caller is responsible for emitting the enclosing
/// `BrtRowHdr` for `row` exactly once before the run of cells that
/// share that row.
///
/// Behaviour:
///   * `Number`  → `BrtCellRk` when the value round-trips through RK
///     encoding (`rk_round_trips_value(...)`); otherwise `BrtCellReal`
///     with the full IEEE 754 double payload.
///   * `Bool`    → `BrtCellBool` (1-byte payload).
///   * `Text`    → `BrtCellIsst` referencing the SST index returned by
///     `sst.intern(...)`. Inline strings are not emitted; the SST
///     interns deduplicate text cells.
///   * `Error`   → `BrtCellError` with the OOXML wire code (or `0x09`
///     for `ErrorCode::Unknown`).
///   * `Blank`   → `BrtCellBlank` (cell-header only).
///   * Formula cells are recognised by `cell.formula_text` being non-
///     empty. When `formula_text` matches the Bundle 4.1 reader stub
///     (`__FORMULON_XLSB_PTG__(<hex>...)`) the captured Ptg bytes are
///     spliced back into a `BrtFmla*` record matching the cached
///     value's kind; otherwise the writer falls back to emitting the
///     cached value as a literal and logs a `xlsb.writer.formula_lost`
///     warning. `Array` and `Lambda` cells are not yet supported and
///     fall through to the literal-blank path with a deferred-log
///     warning.
///
/// `row` is currently used only for diagnostic logs; it does not
/// appear in the cell payload (the enclosing `BrtRowHdr` carries the
/// row index).
void emit_cell(std::vector<std::uint8_t>& dst, const Cell& cell, std::uint32_t row, std::uint32_t col, SstBuilder& sst);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CELL_WRITER_H_
