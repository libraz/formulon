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
// Past every ceiling sits Excel's own bound, `kMaxEvaluatedArrayElements`:
// an array that large abandons the whole formula, which then evaluates to the
// number 0 whatever encloses the array (`abandon_formula`).
//
// Two ceilings are in use. `kMaxDynamicArrayCells` bounds arrays a formula
// conjures out of nothing (SEQUENCE, RANDARRAY, MAKEARRAY, EXPAND); a request
// past it is a user error. `kMaxDerivedArrayCells` bounds arrays derived from
// an input that was itself already admitted under the range-expansion bound,
// where the shape is a consequence of the data rather than of an argument.

#ifndef FORMULON_EVAL_ARRAY_ALLOC_H_
#define FORMULON_EVAL_ARRAY_ALLOC_H_

#include <cstddef>
#include <cstdint>

#include "eval/formula_abandon.h"
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

/// Element count past which Excel abandons a formula that evaluates an array,
/// whatever its shape or origin: floor(2^30 / 10). Mac Excel sizes and
/// evaluates `A:CX+0` (106,954,752) and refuses `A:CY+0` and
/// `SEQUENCE(1048576,103)` (108,003,328). The whole formula then evaluates to
/// the number 0, so `SUM(A:A+1:1)`, `ISERROR(...)` and `ROWS(...)` of such an
/// array are all 0.
inline constexpr std::uint64_t kMaxEvaluatedArrayElements = 107'374'182U;

/// True when an array of `rows` x `cols` elements abandons its formula, which
/// is then marked (`abandon_formula`). The caller must not build the array.
inline bool abandons_formula(std::uint64_t rows, std::uint64_t cols) noexcept {
  if (rows * cols <= kMaxEvaluatedArrayElements) {
    return false;
  }
  abandon_formula();
  return true;
}

/// Allocates a `(rows, cols)` `ArrayValue` and its row-major `Value[]` backing
/// buffer from `arena`, writing the buffer to `out_buffer`.
///
/// Returns `nullptr` — with `out_buffer` set to `nullptr` — when any of the
/// following holds, so a caller that only checks the return value cannot
/// proceed with a stale buffer:
///   * either axis is zero or exceeds its Excel grid bound,
///   * `rows * cols` exceeds `kMaxEvaluatedArrayElements`, which also abandons
///     the formula (`abandons_formula`),
///   * `rows * cols` overflows `std::size_t`,
///   * `rows * cols` exceeds `max_cells`,
///   * the arena cannot satisfy the request.
///
/// The buffer is uninitialised; the caller must write all `rows * cols` cells.
/// Neither pointer is owned by the caller: both live as long as the arena.
ArrayValue* allocate_array_value(std::uint32_t rows, std::uint32_t cols, Arena& arena, Value*& out_buffer,
                                 std::uint64_t max_cells = kMaxDynamicArrayCells);

/// Allocates a `(rows, cols)` array under `kMaxDerivedArrayCells` and fills it
/// row-major from the first `count` entries of `values`, padding the rest with
/// Blank (a clamped range expansion can return fewer cells than its rectangle).
/// Returns `nullptr` under the same conditions as `allocate_array_value`.
ArrayValue* array_from_values(std::uint32_t rows, std::uint32_t cols, const Value* values, std::size_t count,
                              Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_ARRAY_ALLOC_H_
