// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// MS-XLSB per-cell record emitter. The sheet writer walks each
// populated row and calls `emit_cell` for every column slot, which
// dispatches on the cell's `Value::kind()` (and on the presence of a
// formula text) to produce the appropriate `BrtCell*` /
// `BrtFmla*` record bytes.
//
// The dispatcher covers literal cells (number, boolean, text, error,
// blank) and formula cells. A formula cell's `formula_text` is parsed to
// the engine AST and lowered to a Ptg (`rgce`) byte stream via
// `io::xlsb::encode_ptgs`, spliced into a `BrtFmla*` record matching the
// cached value's kind. A formula that cannot be parsed or lowered (a
// token outside the supported Ptg set, e.g. a defined-name or external
// reference) is NOT silently dropped to a literal: `emit_cell` returns
// the encode error so `write_xlsb` can propagate it to the caller.
//
// Design references:
//   * [MS-XLSB] §2.4.x (per-cell record types)

#ifndef FORMULON_IO_XLSB_CELL_WRITER_H_
#define FORMULON_IO_XLSB_CELL_WRITER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "cell.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/sst_writer.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

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
///   * Formula cells (`cell.formula_text` non-empty) → `BrtFmla*` with
///     the parsed-and-encoded Ptg stream. `sheet_names` resolves a
///     qualified reference's sheet to its 0-based `ixti`.
///
/// Returns `kIoXlsbUnsupportedPtg` when the formula cannot be parsed or
/// lowered to the supported Ptg token set; the caller (`write_xlsb`)
/// propagates the failure rather than losing the formula.
///
/// `row` is used only for diagnostic logs; it does not appear in the
/// cell payload (the enclosing `BrtRowHdr` carries the row index).
Expected<void, Error> emit_cell(std::vector<std::uint8_t>& dst, const Cell& cell, std::uint32_t row, std::uint32_t col,
                                SstBuilder& sst, const std::vector<std::string>& sheet_names,
                                const SheetRangeTable& sheet_ranges, const NameTable& name_table);

/// Emits the anchor cell of a spilled dynamic-array formula, matching how
/// Excel structures one: a `BrtFmla*` "shell" record (typed by
/// `anchor_value`) whose `CellParsedFormula` is a lone `PtgExp` token
/// pointing back at the anchor, immediately followed by a `BrtArrFmla`
/// record carrying the real Ptg tokens and the spill footprint
/// (`anchor_row`/`col` .. `last_row`/`last_col`). Every phantom cell in
/// the footprint carries the same shell (see `emit_array_phantom`); the
/// real tokens appear exactly once, here. Returns the encode error when
/// `cell.formula_text` cannot be lowered to the supported Ptg set.
Expected<void, Error> emit_array_anchor(std::vector<std::uint8_t>& dst, const Cell& cell, const Value& anchor_value,
                                        std::uint32_t col, std::uint32_t anchor_row, std::uint32_t last_row,
                                        std::uint32_t last_col, const std::vector<std::string>& sheet_names,
                                        const SheetRangeTable& sheet_ranges, const NameTable& name_table);

/// Emits a phantom cell of a spilled dynamic-array formula: a `BrtFmla*`
/// "shell" record (typed by `cached`, the spilled value at this position)
/// whose `CellParsedFormula` is a lone `PtgExp` token pointing at the
/// region's anchor (`anchor_row`, `anchor_col`). The phantom has no
/// formula of its own; the real tokens live in the anchor's `BrtArrFmla`.
void emit_array_phantom(std::vector<std::uint8_t>& dst, std::uint32_t col, std::uint32_t xf_index, const Value& cached,
                        std::uint32_t anchor_row, std::uint32_t anchor_col);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CELL_WRITER_H_
