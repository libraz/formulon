//
// The single allocation seam for every `ArrayValue` the evaluator produces.
//
// Every array result — dynamic-array builtins, broadcast operators, regex
// capture matrices, VM `MakeArray`, range materialisation — is a
// `(rows, cols)` pair turned into one flat `Value[]` buffer. Computing that
// buffer length as a bare `rows * cols` is the hazard this header exists to
// remove: on wasm32 `std::size_t` is 32 bits, so a crafted pair of dimensions
// wraps, the buffer comes back far smaller than the loop that fills it
// expects, and the fill writes past its end. That is a memory-safety bug, not
// a performance one, which is why the seam is mandatory rather than advisory.
//
// The seam therefore does all four checks in one place:
//   * per-axis bounds against the Excel grid (`Sheet::kMaxRows` / `kMaxCols`),
//   * `checked_mul_size_t` for the buffer length,
//   * a cell-count ceiling chosen by the caller, and
//   * arena-OOM propagation (`nullptr`, which callers surface as `#NUM!`).
//
// Two ceilings are in use. `kMaxDynamicArrayCells` bounds arrays a formula
// conjures out of nothing (SEQUENCE, RANDARRAY, MAKEARRAY, EXPAND); a request
// past it is a user error. `kMaxDerivedArrayCells` bounds arrays derived from
// an input that was itself already admitted under the range-expansion bound,
// where the shape is a consequence of the data rather than of an argument.

#ifndef FORMULON_EVAL_ARRAY_ALLOC_H_
#define FORMULON_EVAL_ARRAY_ALLOC_H_

#include <cstdint>

#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {

class Arena;

namespace eval {

/// Cell ceiling for an array derived from an already-admitted input (an
/// elementwise operator result, a reshape, a filtered copy). The source array
/// passed `kMaxRangeExpansionCells` on the way in, so reusing that bound keeps
/// derivation from being narrower than the read that produced its operand.
inline constexpr std::uint64_t kMaxDerivedArrayCells = kMaxRangeExpansionCells;

// The evaluation arena's byte ceiling has to admit the largest array this
// seam is willing to hand out, or the two limits contradict each other and
// a legitimate result fails as an out-of-memory instead of succeeding.
// `resource_budget.h` cannot state this itself — it must not depend on
// `Value` — so the check lives with the seam that couples them.
static_assert(kMaxDerivedArrayCells * sizeof(Value) <= kMaxEvalArenaBytes,
              "kMaxEvalArenaBytes must admit one full-size derived array");

/// Allocates a `(rows, cols)` `ArrayValue` and its row-major `Value[]` backing
/// buffer from `arena`, writing the buffer to `out_buffer`.
///
/// Returns `nullptr` — with `out_buffer` set to `nullptr` — when any of the
/// following holds, so a caller that only checks the return value cannot
/// proceed with a stale buffer:
///   * either axis is zero or exceeds its Excel grid bound,
///   * `rows * cols` overflows `std::size_t`,
///   * `rows * cols` exceeds `max_cells`,
///   * the arena cannot satisfy the request.
///
/// The buffer is uninitialised; the caller must write all `rows * cols` cells.
/// Neither pointer is owned by the caller: both live as long as the arena.
ArrayValue* allocate_array_value(std::uint32_t rows, std::uint32_t cols, Arena& arena, Value*& out_buffer,
                                 std::uint64_t max_cells = kMaxDynamicArrayCells);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ARRAY_ALLOC_H_
