// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's info-style built-ins (ISNUMBER, ISTEXT, ISBLANK,
// ISLOGICAL, ISERROR, ISERR, ISNA, N, T) into a FunctionRegistry. Kept in
// its own translation unit so the info family can evolve independently of
// the rest of the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_INFO_H_
#define FORMULON_EVAL_BUILTINS_INFO_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the info built-in functions (ISNUMBER, ISTEXT, ISBLANK,
/// ISLOGICAL, ISERROR, ISERR, ISNA, N, T) into `registry`. Intended to be
/// invoked from `register_builtins`.
void register_info_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_INFO_H_
