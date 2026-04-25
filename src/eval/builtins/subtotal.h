// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's SUBTOTAL aggregator into a FunctionRegistry. Lives in
// its own translation unit so the multi-mode dispatch can evolve without
// pulling the rest of the aggregate catalog along.

#ifndef FORMULON_EVAL_BUILTINS_SUBTOTAL_H_
#define FORMULON_EVAL_BUILTINS_SUBTOTAL_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers SUBTOTAL into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_subtotal_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_SUBTOTAL_H_
