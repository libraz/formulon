// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's statistical built-ins (MEDIAN, MODE/MODE.SNGL,
// LARGE/SMALL, PERCENTILE[.INC], QUARTILE[.INC], STDEV[.S]/STDEV.P,
// VAR[.S]/VAR.P) into a FunctionRegistry. Kept in its own translation
// unit so the statistics family can evolve independently of the rest of
// the builtin catalog.

#ifndef FORMULON_EVAL_BUILTINS_STATS_H_
#define FORMULON_EVAL_BUILTINS_STATS_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the statistical built-in functions (MEDIAN, MODE, MODE.SNGL,
/// LARGE, SMALL, PERCENTILE, PERCENTILE.INC, QUARTILE, QUARTILE.INC, STDEV,
/// STDEV.S, STDEV.P, VAR, VAR.S, VAR.P) into `registry`. Intended to be
/// invoked from `register_builtins`.
void register_stats_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_STATS_H_
