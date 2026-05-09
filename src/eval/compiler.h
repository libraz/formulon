// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// AST -> ByteCode compiler.
//
// `compile()` walks a parser `AstNode` tree and lowers it to a
// `ByteCode` body suitable for the (forthcoming) stack-machine VM.
// The bundle is "emit-only": the resulting bytecode is not yet wired
// into the tree-walk evaluator. The VM (Bundle 5.2) will consume it.
//
// Lowering rules:
//
//   * Literals, refs, name refs, structured refs, spill refs, external
//     refs lower to a single `LoadConst` / `LoadRef` / ... opcode.
//   * Arithmetic / concat / comparison / unary lower to the
//     corresponding `BinaryOp` / `UnaryOp` / `Concat` opcode after
//     compiling their operands.
//   * `Call`s lower to `Call op=name argc=arity` after compiling each
//     argument left-to-right, except for the lazy / short-circuit
//     family `IF`, `IFERROR`, `IFNA`, which are lowered to branch
//     instructions: the short-circuit semantics live in the bytecode,
//     not in a runtime function table.
//   * `LET` and `LAMBDA` allocate slots and lower to
//     `StoreLet` / `LoadLet` / `LoadLambdaArg` opcodes; `LAMBDA` itself
//     becomes a `MakeLambda` that points at an inline-compiled body.
//   * `ArrayLiteral` lowers to row-major `LoadConst` of each element
//     followed by `MakeArray rows cols`.
//   * `ErrorPlaceholder` is rejected with `kVmUnsupportedNode`; the
//     parser is responsible for reporting the underlying parse error
//     long before the compiler runs.
//
// All operands are bounds-checked against the 24-bit / 32-bit budgets
// (see `Instruction::kMaxA`). On overflow the compiler returns the
// matching `kVm...Overflow` error.

#ifndef FORMULON_EVAL_COMPILER_H_
#define FORMULON_EVAL_COMPILER_H_

#include "eval/bytecode.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace eval {

/// Lowers `root` to a fresh `ByteCode` body.
///
/// `arena` is used for temporary scratch allocations during compilation
/// (per-kind helpers may stash short-lived metadata there); the resulting
/// `ByteCode` is fully self-contained and does not borrow from the arena
/// once `compile()` returns.
///
/// Returns the compiled body on success, or an `Error` carrying one of
/// the `kVm*` error codes on failure. The compiler never aborts on a
/// well-formed AST: every shape representable by `parser::NodeKind`
/// either lowers successfully or yields an error result.
Expected<ByteCode, Error> compile(const parser::AstNode& root, Arena& arena);

/// Convenience: lowers `root` and then runs the bytecode optimiser
/// (`eval::optimize`) on the result. Equivalent to calling `compile()`
/// followed by `optimize()`. Useful for tests and benchmarks that want
/// the post-optimisation stream in a single call. The default `compile()`
/// remains untouched so callers that need the raw lowering can still
/// access it.
Expected<ByteCode, Error> compile_and_optimize(const parser::AstNode& root, Arena& arena);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COMPILER_H_
