// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's dynamic-array (spilling) built-ins. The dynamic-array
// family is the producer side of the cell-level spill machinery owned by
// `Sheet` / `EvalContext::dispatch_array_result`: each function returns a
// `Value::Array`, which the dispatcher commits as a spill region anchored
// at the formula cell.
//
// Currently registered:
//
//   * `SEQUENCE(rows, [cols], [start], [step])` -- a deterministic
//     `rows x cols` numeric grid. The simplest dynamic-array function (no
//     range-arg dependency, no shape inference) and therefore the canonical
//     end-to-end acceptance test for the spill pipeline.
//
// Other dynamic-array functions (TRANSPOSE, FILTER, SORT, UNIQUE, ...) will
// land in this same translation unit as the spill pipeline grows.

#ifndef FORMULON_EVAL_BUILTINS_DYNAMIC_ARRAY_H_
#define FORMULON_EVAL_BUILTINS_DYNAMIC_ARRAY_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the dynamic-array (spilling) built-in functions into
/// `registry`. Currently registers SEQUENCE; other dynamic-array functions
/// will follow as they are implemented. Intended to be invoked from
/// `register_builtins`.
void register_dynamic_array_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_DYNAMIC_ARRAY_H_
