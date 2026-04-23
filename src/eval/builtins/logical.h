// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's logical built-ins (TRUE/FALSE/NOT/AND/OR) into a
// FunctionRegistry. Kept in its own translation unit so the logical family
// can evolve independently of the rest of the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_LOGICAL_H_
#define FORMULON_EVAL_BUILTINS_LOGICAL_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the logical built-in functions (TRUE, FALSE, NOT, AND, OR) into
/// `registry`. Intended to be invoked from `register_builtins`.
void register_logical_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_LOGICAL_H_
