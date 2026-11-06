//
// Internal seam between the call-dispatch path (`tree_walker/dispatch.cpp`),
// the range-argument collector (`tree_walker/range_args_collect.cpp`) and the
// streamed tail fold (`tree_walker/tail_fold.cpp`). Not part of the public
// evaluator surface.

#ifndef FORMULON_EVAL_TREE_WALKER_DISPATCH_INTERNAL_H_
#define FORMULON_EVAL_TREE_WALKER_DISPATCH_INTERNAL_H_

#include <cstddef>
#include <vector>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tail_array.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

// A failed argument either aborts the call or, for a function that inspects
// its error arguments, lands in `values` like any other value. Returns false
// when the call must abort with `err`.
bool absorb_arg_error(const FunctionDef& def, const Value& err, std::vector<Value>* values);

enum class RangeArgOutcome {
  kNotRange,   // not a range-producing argument; the caller handles it as a scalar
  kCollected,  // consumed into `values`
  kReturn      // the call ends with `result`
};

// Collects the cells of an argument that denotes a range into `values`: for
// a range-aware function every range-shaped spelling, for a scalar function
// a bounded rectangle (which the caller then lifts element-wise). A failed
// argument aborts the call or lands in `values` per `propagate_errors`.
RangeArgOutcome collect_range_arg(const FunctionDef& def, const parser::AstNode& arg_node, Arena& arena,
                                  const FunctionRegistry& registry, const EvalContext& ctx, std::vector<Value>& values,
                                  bool& had_range_shaped_arg, Value& result);

// A derived whole column / row argument of a reducer with `FoldOps`, streamed
// instead of built; `position` counts the collected values before it.
struct StreamedTail {
  std::size_t position;
  const TailArray* tail;
};

// The first error a row-major walk of `ta` reaches, or Blank.
Value first_tail_error(const TailArray& ta);

// Folds the collected `values` with the streamed tails at their positions.
Value fold_with_tails(const FunctionDef& def, const std::vector<Value>& values,
                      const std::vector<StreamedTail>& streamed);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_DISPATCH_INTERNAL_H_
