// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's reference-manipulation builtins whose impls do not need
// to inspect their arguments' AST shape: currently `ADDRESS`. The lazy
// reference forms (`INDIRECT`, `OFFSET`) live in `eval/reference_lazy.*`
// because they dispatch on AST structure (INDIRECT evaluates its text arg,
// OFFSET needs a Ref or RangeOp node for the base).
//
// ADDRESS is a pure scalar text builder: given (row, column, [abs_num],
// [a1], [sheet_text]) it formats the corresponding A1 or R1C1 spelling
// without touching any cell. It therefore rides the eager dispatch path.

#ifndef FORMULON_EVAL_BUILTINS_REFERENCES_H_
#define FORMULON_EVAL_BUILTINS_REFERENCES_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers `ADDRESS` into `registry`. Invoked from `register_builtins`.
void register_reference_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_REFERENCES_H_
