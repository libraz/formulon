//
// Reading a `Ref` as values: the one place a reference becomes cells, and the
// one place those reads are reported.
//
//   * A bare reference (`kArea`, `kUnion`, `kSpan3D`) is read whole, every
//     cell of every area; a 3-D span only where the caller accepts it.
//   * A reference that is an element of a grid reads its top-left cell.
//   * A counted read goes to `EvalContext::note_value_reads` (through the
//     reads delegated to); an uncounted one is not seen by a state observing
//     reads. A `Ref` marked `kDynamic` reports the rectangles it read to
//     `EvalContext::note_dynamic_read` either way, one note per distinct
//     (sheet, rectangle) per call and never merged into a bounding box.

#ifndef FORMULON_EVAL_REFERENCE_DEREF_H_
#define FORMULON_EVAL_REFERENCE_DEREF_H_

#include <cstdint>
#include <vector>

#include "eval/function_registry.h"
#include "eval/tail_array.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

/// Whether a read counts toward circularity.
enum class ReadCount : std::uint8_t {
  kCounted,    ///< A read of the cells' values.
  kUncounted,  ///< A read of something else (extent, formula text, phonetic text).
};

/// The cells of a bare reference in row-major order, area by area.
struct RefCells {
  std::vector<Value> cells;
  /// Shape of the rectangle when the reference is one `kArea`; 0 x 0 for a
  /// union or a span.
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
};

/// Reads every cell of every area of `ref` (a `kUnion` area by area, a
/// `kSpan3D` sheet by sheet), a whole-axis area narrowed as a range expansion
/// narrows it. `#REF!` for a `kSpan3D` unless `allow_span3d`; `#VALUE!` for a
/// grid or a non-`Ref`.
Expected<RefCells, ErrorCode> deref_cells(const Value& ref, bool allow_span3d, Arena& arena,
                                          const FunctionRegistry& registry, const EvalContext& ctx, ReadCount count);

/// Reads a `kArea` as the walker reads the same range: the cell for one cell,
/// an `Array` for a rectangle, a `TailArray` for a whole-axis one. A grid reads
/// as an array of `#VALUE!` (an error element keeps its error); a `kUnion` is
/// `#VALUE!` and a `kSpan3D` `#REF!`.
Shaped deref_shaped(const Value& ref, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    ReadCount count);

/// Reads one cell of `ref`: the top-left of its first area, or, for a grid, of
/// its first element (an error element is its error).
Value deref_top_left(const Value& ref, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                     ReadCount count);

/// Reads the grid `ref` as the function with class `cls` takes it.
///
///   * `kElementError`: an `Array` of the grid's shape, each element `#VALUE!`.
///   * `kPerValue`: an `Array` of the grid's shape; a one-cell element gives
///     its value, a larger one `#VALUE!`.
///   * `kTopLeft`: an `n x 1` `Array` of the elements' top-left cells, kept
///     through `def`'s range-sourced filters (all kept when `def` is null);
///     when none is kept, a 1 x 1 `Array` of a blank. A filter that rejects a
///     cell gives that error.
///   * `kOpaque`: an `Array` of the grid's shape, each element a blank marked
///     `kOpaqueElement`.
///   * `kWholeValue` / `kWholeNA`: `#VALUE!` / `#N/A`.
///
/// In every `Array` an error element keeps its error. The classes the
/// dispatcher handles itself (`kPerRef`, `kKeepRef`, `kReshape`,
/// `kLambdaElement`), and a `ref` that is not a grid, are returned unchanged.
Value materialise_grid(const Value& ref, GridArg cls, const FunctionDef* def, Arena& arena,
                       const FunctionRegistry& registry, const EvalContext& ctx, ReadCount count);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_DEREF_H_
