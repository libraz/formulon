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

#include <cstdint>

#include "value.h"

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Parsed arguments of `SEQUENCE(rows, [cols], [start], [step])` or
/// `RANDARRAY([rows], [cols], [min], [max], [whole_number])`: `first` / `second`
/// are start / step or min / max.
struct GeneratorParams {
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  double first = 0.0;
  double second = 0.0;
  bool whole_number = false;
};

/// Validates the evaluated arguments of SEQUENCE (`sequence`) or RANDARRAY and
/// reports the grid it returns, without building it. Returns false, setting
/// `*out_err` to the error the function returns instead, when an argument is
/// rejected.
bool generator_params(const Value* args, std::uint32_t arity, bool sequence, GeneratorParams* out, ErrorCode* out_err);

/// Registers the dynamic-array (spilling) built-in functions into
/// `registry`. Currently registers SEQUENCE; other dynamic-array functions
/// will follow as they are implemented. Intended to be invoked from
/// `register_builtins`.
void register_dynamic_array_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_DYNAMIC_ARRAY_H_
