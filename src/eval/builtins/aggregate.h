// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's aggregate built-ins (SUM/MIN/MAX/AVERAGE/PRODUCT,
// COUNT/COUNTA/COUNTBLANK, CONCAT/CONCATENATE, LEN) into a FunctionRegistry.
// Kept in its own translation unit so the aggregate family can evolve
// independently of the rest of the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_AGGREGATE_H_
#define FORMULON_EVAL_BUILTINS_AGGREGATE_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the aggregate built-in functions (SUM, MIN, MAX, AVERAGE, PRODUCT,
/// COUNT, COUNTA, COUNTBLANK, CONCAT, CONCATENATE, LEN) into `registry`.
/// Intended to be invoked from `register_builtins`.
void register_aggregate_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_AGGREGATE_H_
