//
// Evaluator-internal representation of a whole-column / whole-row array: a
// dense head followed by one repeated tail row (or column). It never becomes a
// `Value`; `Shaped` carries it between evaluator layers and `densify` is the
// only way back to a dense `Value::array`.

#ifndef FORMULON_EVAL_TAIL_ARRAY_H_
#define FORMULON_EVAL_TAIL_ARRAY_H_

#include <cstdint>

#include "value.h"

namespace formulon {

class Arena;

namespace eval {

/// Axis along which a `TailArray` is compressed.
enum class TailAxis : std::uint8_t {
  kRows,  ///< Whole column: the tail continues down the rows.
  kCols,  ///< Whole row: the tail continues across the columns.
};

/// Declared-size array stored as a dense head plus a repeated tail.
/// Trivially destructible; lives in an `Arena`.
struct TailArray {
  std::uint32_t rows;  ///< Declared row count.
  std::uint32_t cols;  ///< Declared column count.
  std::uint32_t head;  ///< Dense length along `axis` (0 allowed).
  TailAxis axis;
  const Value* cells;   ///< Row-major head: `head x cols` (kRows) or `rows x head` (kCols).
  const Value* tail;    ///< Repeated row (`cols` cells, kRows) or column (`rows` cells, kCols).
  bool from_reference;  ///< Read straight from a reference, not produced by an operation.
};

/// Evaluator-internal result: a plain `Value`, or a `TailArray` when
/// `tail_array` is non-null (then `value` is unused).
struct Shaped {
  Value value = Value::blank();
  const TailArray* tail_array = nullptr;
};

/// Allocates a `TailArray` in `arena`. Returns `nullptr` when the arena is
/// exhausted. `cells` and `tail` are stored as given, not copied.
const TailArray* make_tail_array(Arena& arena, std::uint32_t rows, std::uint32_t cols, std::uint32_t head,
                                 TailAxis axis, const Value* cells, const Value* tail, bool from_reference);

/// The read of a whole column / row: `cells` (the walked head, possibly null
/// when `head` is 0) followed by reference-grid blanks to the declared size.
/// `#NUM!` when the arena is exhausted.
Shaped make_reference_tail(Arena& arena, std::uint32_t rows, std::uint32_t cols, std::uint32_t head, TailAxis axis,
                           const Value* cells);

/// Value at position (`r`, `c`) of `ta`; both indices must be within the declared size.
const Value& tail_array_at(const TailArray& ta, std::uint32_t r, std::uint32_t c);

/// Dense value of `s`: `s.value` for a non-tail result, otherwise a row-major
/// `Value::array`. Over `kMaxDerivedArrayCells` it is the error the dense path
/// gives: `#CALC!` for a reference read (as range expansion), `#NUM!` otherwise.
Value densify(const Shaped& s, Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TAIL_ARRAY_H_
