//
// Registers Excel's SUBTOTAL aggregator into a FunctionRegistry. Lives in
// its own translation unit so the multi-mode dispatch can evolve without
// pulling the rest of the aggregate catalog along.

#ifndef FORMULON_EVAL_BUILTINS_SUBTOTAL_H_
#define FORMULON_EVAL_BUILTINS_SUBTOTAL_H_

#include <cstdint>

#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Runs SUBTOTAL's mode dispatch over already-collected argument values.
/// `args[0]` is the function code (1..11 or 101..111); the remainder are the
/// data cells in row-major order, exactly as the range-aware dispatcher
/// flattens them.
///
/// This is the single implementation of SUBTOTAL's semantics. The lazy
/// front-end in `eval/aggregate_lazy.cpp` calls it after dropping cells that
/// sit on hidden rows, and the eager registration below calls it directly.
Value subtotal_apply(const Value* args, std::uint32_t arity, Arena& arena);

/// Registers SUBTOTAL into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_subtotal_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_SUBTOTAL_H_
