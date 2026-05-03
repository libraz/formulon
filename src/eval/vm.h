// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stack-machine VM that executes the bytecode produced by `eval::compile()`.
//
// `execute()` is a stateless entry point: every transient (operand stack,
// LET-slot vector, lambda-frame chain) is built on the supplied evaluation
// arena for the duration of a single call and freed by the time the
// function returns. The VM never mutates `bc`, `registry`, or `ctx`.
//
// Design contracts:
//   * Errors as values, not as out-of-band escapes. Excel-visible faults
//     (`#DIV/0!`, `#VALUE!`, ...) flow as `Value::Error` results returned
//     by `execute()`. The `Expected<>` channel reports VM-level faults
//     (`kVmEmptyBytecode`, `kVmStackUnderflow`, ...) which the caller
//     should treat as engine bugs rather than business errors.
//   * Tree-walker parity. Bundle 5.2 is purely additive: the tree-walker
//     remains the single source of truth, and the VM is exercised in
//     parallel under `FORMULON_VM_PARITY_CHECK` to surface any drift.
//     Bundle 5.3 will swap the entry point.
//   * No `accepts_ranges` AST inspection. The bytecode IR represents only
//     scalar arguments to function calls; range-aware aggregators that
//     receive a literal `A1:B2` argument lose their AST shape during
//     compile and the VM evaluates them through whatever `Value` the
//     compiler-side LoadRange path produced. This is the documented IR
//     limitation that bundles 5.3+ will revisit.
//
// See `backup/plans/02-calc-engine.md` §2.4 for the design and
// `backup/plans/26-implementation-plan.md` Phase 5 for the bundle roadmap.

#ifndef FORMULON_EVAL_VM_H_
#define FORMULON_EVAL_VM_H_

#include "eval/bytecode.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

class FunctionRegistry;
class EvalContext;

/// Executes `bc` against `registry` and `ctx`, returning the result of the
/// final `Return` / `Halt` instruction.
///
/// `eval_arena` backs every transient `Value` payload produced during
/// execution (concatenated text, broadcast arrays, lambda closures). It
/// must outlive the returned `Value`; the result borrows from it the same
/// way the tree-walk evaluator's result borrows from its arena.
///
/// Excel-visible business errors (`#DIV/0!`, `#VALUE!`, ...) are reported
/// as `Value::Error` results. The `Expected<>` channel is reserved for
/// VM-level faults: empty bytecode, stack underflow / overflow, malformed
/// jump targets, out-of-range opcodes, and other corruption indicators.
Expected<Value, Error> execute(const ByteCode& bc, Arena& eval_arena, const FunctionRegistry& registry,
                               const EvalContext& ctx);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_VM_H_
