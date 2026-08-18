//
// Projection of a spilled range onto the value model, shared by every
// path that can name a spill anchor: the `#` spill operator in the tree
// walker and in the bytecode VM, and ANCHORARRAY.
//
// Keeping one implementation is what stops the three from drifting: the
// error each failure mode reports is Excel-observable, and the three
// paths must agree on it cell for cell.

#ifndef FORMULON_EVAL_SPILL_ANCHOR_H_
#define FORMULON_EVAL_SPILL_ANCHOR_H_

#include <cstdint>
#include <string_view>

#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

// Resolves the spill region anchored at (`sheet`, `row`, `col`) and
// copies it into a fresh `ArrayValue` allocated from `arena`. An empty
// `sheet` means the context's current sheet. The array's cells are
// shallow copies of the region's; text payloads keep pointing into the
// region's owned strings, which outlive any single evaluation arena.
//
// Returns `nullptr` on failure with `*out_err` written:
//
//   * `#NAME?` — the context has no current sheet.
//   * `#REF!`  — the named sheet is unknown, the anchor is off the grid,
//                or no spill region is anchored at that address.
//   * `#NUM!`  — the arena is exhausted. It is the only failure that
//                reports `#NUM!`, so a caller that carries a structured
//                out-of-memory error can key on it.
ArrayValue* project_spill_at_anchor(std::string_view sheet, std::uint32_t row, std::uint32_t col, Arena& arena,
                                    const EvalContext& ctx, ErrorCode* out_err);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_SPILL_ANCHOR_H_
