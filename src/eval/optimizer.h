//
// ByteCode optimiser (compile-time, behaviour-preserving).
//
// The optimiser runs after `eval::compile()` and before `eval::execute()`.
// It is purely additive: every pass MUST produce bytecode that, when fed
// to the VM, yields the bit-exact same `Value` (raw IEEE-754 bits for
// `Number`, identical text bytes for `Text`, identical error code for
// `Error`) as the un-optimised input. Anything a pass cannot prove safe,
// it leaves alone. The optimiser never invents Excel-visible faults; if
// constant-folding `1/0` reveals `#DIV/0!`, that is the same value the
// VM would have produced at run time, so the IR keeps it.
//
// Four passes run in fixed order:
//
//   1. Constant folding (active in this bundle).
//      Pattern `LoadConst K1; LoadConst K2; (BinaryOp|UnaryOp)` collapses
//      to a single `LoadConst K_new` whose pooled value is the result of
//      evaluating the operator at compile time using the same
//      `apply_arithmetic` / `apply_unary` / `apply_comparison` helpers
//      the VM uses, so the two paths cannot diverge.
//
//      Skipped intentionally:
//        * `Concat` and the `&` form: the result is a `Text` borrowing
//          arena memory whose lifetime cannot be guaranteed across the
//          `ByteCode` boundary. A future bundle that interns folded
//          text into `bc.string_storage` will lift this restriction.
//        * Any binary op whose operand is an `Array`, `Ref`, or `Lambda`
//          constant — these never arise from the current compiler but
//          we defend against them.
//
//   2. Name inlining (stub in this bundle).
//      Reserved for a future pass that replaces `LoadName N` with
//      `LoadConst K` when the workbook-scope name resolves to a literal.
//      `compile()` does not currently take a `Workbook&`, so the pass
//      is a no-op and only the dispatch order is fixed in stone here.
//
//   3. Range canonicalisation (active in this bundle).
//      Pattern `LoadRef A; LoadRef B; LoadRange 0xFF` is canonicalised
//      so that the two endpoints are emitted in lexicographic order
//      (`(sheet, col, row)`), matching `tree_walker.cpp::eval_range`'s
//      run-time normalisation. Degenerate ranges (`A == B`) collapse to
//      a single `LoadRef A`.
//
//   4. Branch hoisting (skeleton in this bundle).
//      The pass walks the bytecode and identifies `cond; JumpIfFalse;
//      LoadConst K1; Jump; LoadConst K2;` sequences but does not yet
//      rewrite them because there is no `Select` opcode in the IR. The
//      pass increments `branch_hoist_opportunities` so a later bundle
//      that adds the opcode can quantify the gain before the IR change.
//
// Errors:
//   The optimiser only surfaces engine-level faults via `Expected<>`:
//   pool overflow (`kVmConstPoolOverflow`) when re-pooling a folded
//   constant exhausts the 24-bit `Instruction::a` budget, or
//   `kVmOptimizerFailed` for an internal invariant violation. It never
//   surfaces Excel-visible faults — those flow through the bytecode as
//   `LoadConst Value::error(...)`.

#ifndef FORMULON_EVAL_OPTIMIZER_H_
#define FORMULON_EVAL_OPTIMIZER_H_

#include <cstdint>

#include "eval/bytecode.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace eval {

/// Per-run telemetry for the optimiser. `optimize()` increments these
/// counters so callers (tests, benchmarks, the WASM size profiler) can
/// quantify what the passes did. The struct is zero-initialised by
/// default so callers do not need to reset it before each call.
struct OptimizerStats {
  /// Number of `(LoadConst, LoadConst, BinaryOp)` and
  /// `(LoadConst, UnaryOp)` triplets/pairs that the constant-folding
  /// pass replaced with a single `LoadConst`.
  std::uint32_t constants_folded = 0;
  /// Number of `LoadName` instructions the name-inlining pass replaced
  /// with `LoadConst`. The pass is a no-op in this bundle, so this
  /// counter is currently always zero.
  std::uint32_t names_inlined = 0;
  /// Number of `LoadRange` triplets the range-canonicalisation pass
  /// either collapsed (degenerate `A:A`) or reordered (`B5:A1` ->
  /// `A1:B5`).
  std::uint32_t ranges_canonicalized = 0;
  /// Number of `JumpIfFalse` triplets the branch-hoisting pass actually
  /// rewrote. Always zero in this bundle (the IR has no `Select`
  /// opcode yet).
  std::uint32_t branches_hoisted = 0;
  /// Number of `JumpIfFalse` triplets the branch-hoisting pass would
  /// have rewritten if the IR supported a `Select` opcode. Used as
  /// a sanity check for the follow-up bundle that adds the opcode.
  std::uint32_t branch_hoist_opportunities = 0;
};

/// Runs the four optimisation passes in fixed order on `bc` and returns
/// the resulting `ByteCode`.
///
/// The input `bc` is consumed (passed by value) so the caller can move
/// it in directly; the returned `ByteCode` is a fresh body that does
/// not alias the input. Test callers that want to compare pre- and
/// post-optimisation streams should clone the input via the test-only
/// equality helper before invoking `optimize()`.
///
/// `arena` is reserved for transient scratch; the current pass set does
/// not use it but the parameter is kept stable so future passes
/// (e.g. text-fold interning) need not change the public API.
///
/// `stats` is optional; when non-null the per-pass counters are
/// incremented in place. Errors are reserved for engine-level faults:
/// see `kVmOptimizerFailed` and `kVmConstPoolOverflow`. The optimiser
/// never invents Excel-visible faults.
Expected<ByteCode, Error> optimize(ByteCode bc, Arena& arena, OptimizerStats* stats = nullptr);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_OPTIMIZER_H_
