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
//   * Range arguments via LoadRange. A bare `Ref:Ref` range (`A1:B2`) is
//     lowered to two `LoadRef`s plus a `LoadRange` marker; the VM expands
//     the rectangle into an `Array` value, which range-aware aggregators
//     (`accepts_ranges`) then flatten during `Call`. Complex range
//     endpoints (`OFFSET(...):B5`) are not reconstructable from the IR and
//     surface `#VALUE!`, matching the tree-walker's non-Ref-endpoint
//     fallback. This residual gap is the documented IR limitation that
//     future bundles will revisit.

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
