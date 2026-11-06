//
// Streamed fold of whole column / row arguments for reducers that carry
// `FoldOps`: a derived tail array is walked head-then-repeated-tail instead
// of being built, and the result matches the dense row-major walk.

#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/function_registry.h"
#include "eval/range_args.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/dispatch_internal.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Feeds `line`'s filtered cells through `def`'s fold `repeat` times in order.
bool fold_line(const FunctionDef& def, const Value* line, std::size_t count, std::uint64_t repeat,
               std::vector<Value>* scratch, FoldState* state, Value* err) {
  scratch->clear();
  if (!append_array_sourced_values(def, line, count, scratch, err)) {
    return false;
  }
  for (std::uint64_t k = 0; k < repeat; ++k) {
    for (const Value& v : *scratch) {
      if (!def.fold->step(*state, v, err)) {
        return false;
      }
    }
  }
  return true;
}

// Streams `ta` through `def`'s fold in row-major order, the repeated tail
// taken once per position it stands for, so the result matches the dense
// walk without building the array.
bool fold_tail(const FunctionDef& def, const TailArray& ta, FoldState* state, Value* err) {
  std::vector<Value> scratch;
  if (ta.axis == TailAxis::kRows) {
    return fold_line(def, ta.cells, static_cast<std::size_t>(ta.head) * ta.cols, 1U, &scratch, state, err) &&
           (ta.head >= ta.rows || fold_line(def, ta.tail, ta.cols, ta.rows - ta.head, &scratch, state, err));
  }
  for (std::uint32_t r = 0; r < ta.rows; ++r) {
    if (!fold_line(def, ta.cells + static_cast<std::size_t>(r) * ta.head, ta.head, 1U, &scratch, state, err) ||
        (ta.head < ta.cols && !fold_line(def, ta.tail + r, 1U, ta.cols - ta.head, &scratch, state, err))) {
      return false;
    }
  }
  return true;
}

}  // namespace

// Folds the collected `values` with the streamed tails at their positions.
Value fold_with_tails(const FunctionDef& def, const std::vector<Value>& values,
                      const std::vector<StreamedTail>& streamed) {
  FoldState state;
  Value err = Value::blank();
  std::size_t next = 0;
  const auto feed_values = [&](std::size_t end) {
    for (; next < end; ++next) {
      if (!def.fold->step(state, values[next], &err)) {
        return false;
      }
    }
    return true;
  };
  for (const StreamedTail& segment : streamed) {
    if (!feed_values(segment.position) || !fold_tail(def, *segment.tail, &state, &err)) {
      return err;
    }
  }
  if (!feed_values(values.size())) {
    return err;
  }
  return def.fold->finish(state);
}

}  // namespace eval
}  // namespace formulon
