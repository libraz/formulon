// Copyright 2026 libraz. Licensed under the MIT License.
//
// Built-in function registration entry point. The set of functions wired in
// here is the canonical Formulon function library; host extensions may add
// further entries via `FunctionRegistry::register_function` after the default
// registry is constructed.

#ifndef FORMULON_EVAL_BUILTINS_H_
#define FORMULON_EVAL_BUILTINS_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers every built-in function definition into `registry`. Intended to
/// be called once, immediately after the registry is constructed (the default
/// registry does this lazily inside `default_registry()`).
void register_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_H_
