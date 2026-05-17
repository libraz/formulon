// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Private helper for the tree-walk evaluator: a tiny RAII counter guard
// that bounds runaway recursion through `EvalContext::resolve_ref` and
// user-defined LAMBDA closures. Kept header-only so the two evaluator
// translation units that need it (`tree_walker/walker.cpp` and
// `tree_walker/dispatch.cpp`) can construct guards inline without an
// extra function-call hop.
//
// This header is internal to the tree-walker family and is not part of
// the public evaluator surface — production callers reach the evaluator
// through `eval/tree_walker.h`.

#ifndef FORMULON_EVAL_TREE_WALKER_DEPTH_GUARD_H_
#define FORMULON_EVAL_TREE_WALKER_DEPTH_GUARD_H_

#include <cstdint>

namespace formulon {
namespace eval {

// Hard caps on recursion depth. The parser already enforces a parse-depth
// limit of 128 (see `parser::ParserOptions::max_parse_depth`), which bounds
// stack growth from a single formula. These two evaluator-side caps defend
// against the orthogonal vectors that bypass that bound:
//
//   * `kMaxEvalDepth` — bounds linear cell-chain recursion through
//     `EvalContext::resolve_ref`. A workbook of `A1=A2, A2=A3, ..., A1000=1`
//     is not a cycle (so `EvalState::push_cell` does not flag it), and each
//     resolved formula spawns a fresh `eval_node` recursion. Without this
//     cap a chain of ~1000 cells overflows the WASM 256-512 KB stack.
//
//   * `kMaxLambdaDepth` — bounds runtime recursion through user-defined
//     LAMBDA closures (e.g. `LET(f, LAMBDA(n, f(n+1)), f(0))`). The body
//     AST stays small so `kMaxEvalDepth` does not trigger; the recursion
//     lives in `invoke_lambda` re-entering itself.
//
// On overflow the offending sub-expression returns `#CALC!` (the same
// Excel-visible code Mac Excel surfaces for indeterminate / runaway lambda
// recursion). The internal `kEvalStackOverflow` code is reserved for the
// `Expected<T, Error>` plumbing (currently unused on this path).
constexpr std::uint32_t kMaxEvalDepth = 512;
constexpr std::uint32_t kMaxLambdaDepth = 256;

// RAII guard: bumps `*p` on construction (when `*p < cap`) and decrements
// on destruction. When the cap was already reached, `exceeded()` reports
// `true` and the counter is left untouched so a later sibling in the same
// frame does not double-decrement past zero. Null `p` disables tracking
// entirely — `exceeded()` always returns `false` — which preserves
// behaviour for ad-hoc callers that bypass `evaluate()`.
class EvalDepthGuard {
 public:
  EvalDepthGuard(std::uint32_t* p, std::uint32_t cap) noexcept : p_(p), exceeded_(p != nullptr && *p >= cap) {
    if (p_ != nullptr && !exceeded_) {
      ++(*p_);
    }
  }
  ~EvalDepthGuard() noexcept {
    if (p_ != nullptr && !exceeded_) {
      --(*p_);
    }
  }
  EvalDepthGuard(const EvalDepthGuard&) = delete;
  EvalDepthGuard& operator=(const EvalDepthGuard&) = delete;
  EvalDepthGuard(EvalDepthGuard&&) = delete;
  EvalDepthGuard& operator=(EvalDepthGuard&&) = delete;

  bool exceeded() const noexcept { return exceeded_; }

 private:
  std::uint32_t* p_;
  bool exceeded_;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_TREE_WALKER_DEPTH_GUARD_H_
