// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// MS-XLSB function-id ↔ name mapping. `BrtCellFmla`'s embedded `PtgFunc`
// and `PtgFuncVar` tokens carry a 16-bit function id (rather than a
// textual name) drawn from a stable enumeration that goes back to the
// XLS binary format. This module exposes a flat `constexpr` table sorted
// by id so the Reader can resolve `PtgFunc(id)` to a name and the Writer
// can do the reverse.
//
// The table covers the classic Excel function set up to ~0x17F. Newer
// (`_xlfn.*`) functions live in higher id ranges and are added here as
// they become Reader/Writer-relevant. Unknown ids are surfaced by
// returning `nullptr`; the caller's policy (the Reader emits `#NAME?`
// per backup/plans/21-xlsb-ptg.md §21.6) lives outside this table.
//
// Design references:
//   * backup/plans/21-xlsb-ptg.md §21.6 (function-id mapping)
//   * [MS-XLSB] §2.5.97.74 (PtgFunc) and §2.5.97.75 (PtgFuncVar)
//   * [MS-XLS] §2.5.198.16 (Cetab — historical id assignments)

#ifndef FORMULON_IO_XLSB_FUNC_ID_TABLE_H_
#define FORMULON_IO_XLSB_FUNC_ID_TABLE_H_

#include <cstdint>
#include <string_view>

namespace formulon {
namespace io {
namespace xlsb {

/// One row of the function-id mapping.
///
///   * `id`        — 16-bit MS-XLSB function id ([MS-XLSB] §2.5.97.74).
///   * `name`      — uppercase Excel function name (`"SUM"`, `"VLOOKUP"`, …).
///                   Pointer references a static string literal with
///                   program lifetime.
///   * `arg_min`   — minimum arity. For `PtgFunc` (fixed-arity) this is
///                   also the exact arity; for `PtgFuncVar` it's the
///                   floor.
///   * `arg_max`   — maximum arity. For variadic functions (`variadic =
///                   true`) the maximum is logically unbounded; we
///                   record `255` as a sentinel.
///   * `variadic`  — whether the function is encoded with `PtgFuncVar`
///                   (variable arity, length byte preceding the id) or
///                   with `PtgFunc` (fixed arity baked into the id).
struct XlsbFuncEntry {
  std::uint16_t id;
  const char* name;
  std::uint8_t arg_min;
  std::uint8_t arg_max;
  bool variadic;
};

/// Number of rows in the dispatch table. Exposed so tests can assert
/// the table is sorted without hard-coding the count.
extern const std::size_t kXlsbFuncEntryCount;

/// Pointer to the first row of the table; the table is sorted by `id`
/// for O(log N) binary-search lookup.
extern const XlsbFuncEntry* const kXlsbFuncEntries;

/// Returns the table row whose `id` matches, or `nullptr` when no row
/// matches.
const XlsbFuncEntry* lookup_func_by_id(std::uint16_t id);

/// Returns the table row whose `name` matches case-insensitively, or
/// `nullptr` when no row matches. Linear scan — the lookup is rare
/// (Writer side) and the table is small.
const XlsbFuncEntry* lookup_func_by_name(std::string_view name);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_FUNC_ID_TABLE_H_
