// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Iterative-calculation solver for circular reference SCCs.
//
// When a workbook opts in to iterative calculation (Excel's "Enable
// iterative calculation" workbook option), formulas that participate in a
// cyclic SCC are no longer short-circuited to `#REF!`. Instead, the recalc
// engine hands the SCC to `run_iterative_solve()`, which fixed-point
// iterates the cells up to `IterativeOptions::max_iterations` times, looking
// for either:
//
//   * **convergence** — the largest absolute change across the SCC drops
//     below `IterativeOptions::max_change`. The solver stops and the cells
//     keep their last computed values; or
//   * **divergence** — three successive iterations whose maximum delta is
//     non-decreasing (and above `max_change`). The solver writes `#NUM!`
//     to every member and signals failure.
//
// The convergence test is on the *absolute* delta of numeric values, per
// Excel's specification (no relative tolerance, no per-cell weighting). Any
// change in `ValueKind` (Number -> Text, Number -> Error, etc.) bumps the
// delta to +infinity so a "value flip" never reads as converged; in the
// general case Excel cannot encode a converged non-numeric cycle anyway.
//
// The solver is single-threaded and synchronous. It owns nothing: callers
// supply the SCC list, the per-cell evaluation lambda, and the per-cell
// commit lambda. The caller is also responsible for snapshotting and
// restoring any state that should not be visible to the solver — once
// `run_iterative_solve` returns, the cell store has been mutated either
// to the converged values or to `#NUM!` sentinels.
//
// See backup/plans/02-calc-engine.md §2.7.3 for the design corpus entry.

#ifndef FORMULON_EVAL_ITERATIVE_SOLVER_H_
#define FORMULON_EVAL_ITERATIVE_SOLVER_H_

#include <cstdint>
#include <functional>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "value.h"

namespace formulon::eval {

/// Excel's default iteration cap when iterative calc is enabled.
constexpr std::uint32_t kDefaultMaxIterations = 100U;
/// Excel's default absolute convergence threshold.
constexpr double kDefaultMaxChange = 0.001;

/// User-facing knobs governing iterative-calc behaviour. Mirrors Excel's
/// "Enable iterative calculation" workbook option. The defaults match
/// Excel: iterative calc disabled, capped at `kDefaultMaxIterations`
/// passes, `kDefaultMaxChange` absolute max-change threshold.
struct IterativeOptions {
  /// When false, circular SCCs surface `#REF!` (the legacy behaviour).
  /// When true, the recalc engine hands them to `run_iterative_solve`.
  bool enabled = false;
  /// Maximum number of evaluation passes the solver will run before
  /// giving up. Must be at least 1 to make any progress; the solver
  /// silently treats 0 as 1 to avoid an empty-loop edge case.
  std::uint32_t max_iterations = kDefaultMaxIterations;
  /// Absolute convergence threshold. The solver stops as soon as
  /// `max(|v_n - v_{n-1}|) < max_change` across all numeric members. A
  /// non-positive value forces the solver to run for the full
  /// `max_iterations` since strict-less-than against any non-negative
  /// delta will never hold.
  double max_change = kDefaultMaxChange;
};

/// Outcome of a single `run_iterative_solve` invocation.
struct IterativeOutcome {
  /// True when every member's last-pass delta dropped below
  /// `max_change`. Mutually exclusive with `diverged` and `aborted`.
  bool converged = false;
  /// True when the solver detected divergence (three successive passes
  /// whose maximum delta did not decrease and stayed above `max_change`)
  /// and aborted with `#NUM!`. Mutually exclusive with `converged` and
  /// `aborted`.
  bool diverged = false;
  /// True when the user-supplied progress callback returned `false` and
  /// the solver returned early. The cell store is left in its current
  /// partially-converged state; the caller should treat the result as
  /// "not converged". Mutually exclusive with `converged` and `diverged`.
  bool aborted = false;
  /// Number of iterations actually executed, in `[1, max_iterations]`.
  /// Always at least 1: the solver always evaluates the SCC once before
  /// inspecting the convergence condition.
  std::uint32_t iterations_run = 0U;
};

/// Iteration progress callback. Invoked after each Gauss-Seidel sweep.
///
/// Arguments:
///   iteration         — 1-based iteration counter.
///   max_residual      — the maximum |new - old| across all formula cells
///                       updated during the sweep that just completed.
///                       `+infinity` for any iteration in which a cell's
///                       value kind changed (Number -> Text etc.) or in
///                       which any cell produced NaN.
///   max_iterations    — the iteration cap configured for this solve.
///   user_data         — the opaque pointer the caller passed alongside
///                       the callback; the solver does not interpret it.
///
/// Return `true` to continue iterating, `false` to abort early. An aborted
/// solve leaves the workbook in its current partially-converged state; the
/// solver returns with `IterativeOutcome::aborted == true`.
using IterativeProgressCb = bool (*)(std::uint32_t iteration, double max_residual, std::uint32_t max_iterations,
                                     void* user_data);

/// Type-erased core of the iterative solver. The template-friendly
/// wrapper below forwards into this function so the heavy logic lives in
/// the translation unit, not in every caller.
///
/// `evaluate_one(cell)` MUST evaluate `cell`'s formula against the
/// caller's current cell store and return the resulting `Value`. It is
/// invoked once per member per iteration. `commit(cell, value)` MUST
/// update the cell store so the *next* `evaluate_one` call observes the
/// new value (typically `Sheet::set_cell_cached_value`). Both callbacks
/// are invoked synchronously and serially; no thread safety is required.
///
/// `progress` is an optional callback invoked once per Gauss-Seidel
/// sweep. Pass `nullptr` to opt out — behaviour is then identical to
/// the legacy callback-less contract. The callback returns `true` to
/// continue iterating; returning `false` makes the solver stop early
/// with `IterativeOutcome::aborted == true`. `progress_user_data` is
/// forwarded verbatim to every callback invocation.
///
/// On divergence, the solver itself calls `commit(cell, #NUM!)` for every
/// SCC member before returning. On convergence or iteration-limit
/// exhaustion, the cell store retains whatever the last iteration
/// produced. On callback-driven abort the cell store is also left as-is.
IterativeOutcome run_iterative_solve_impl(const std::vector<CellNodeId>& scc, const IterativeOptions& opts,
                                          const std::function<Value(CellNodeId)>& evaluate_one,
                                          const std::function<void(CellNodeId, Value)>& commit,
                                          IterativeProgressCb progress, void* progress_user_data);

/// Runs the iterative solver on a single SCC. See
/// `run_iterative_solve_impl` for the contract. The template wrapper
/// exists purely so callers can pass any callable (lambdas, function
/// pointers, etc.) without manually constructing `std::function`.
template <typename Eval, typename Commit>
IterativeOutcome run_iterative_solve(const std::vector<CellNodeId>& scc, const IterativeOptions& opts,
                                     Eval&& evaluate_one, Commit&& commit) {
  return run_iterative_solve_impl(scc, opts, std::function<Value(CellNodeId)>(std::forward<Eval>(evaluate_one)),
                                  std::function<void(CellNodeId, Value)>(std::forward<Commit>(commit)),
                                  /*progress=*/nullptr, /*progress_user_data=*/nullptr);
}

/// Overload that threads an optional progress callback into the solver.
/// The callback is invoked after each Gauss-Seidel sweep with the
/// 1-based iteration index, the max delta observed during the sweep,
/// and the configured iteration cap. Returning `false` from the
/// callback aborts the solve and surfaces `IterativeOutcome::aborted`.
template <typename Eval, typename Commit>
IterativeOutcome run_iterative_solve(const std::vector<CellNodeId>& scc, const IterativeOptions& opts,
                                     Eval&& evaluate_one, Commit&& commit, IterativeProgressCb progress,
                                     void* progress_user_data) {
  return run_iterative_solve_impl(scc, opts, std::function<Value(CellNodeId)>(std::forward<Eval>(evaluate_one)),
                                  std::function<void(CellNodeId, Value)>(std::forward<Commit>(commit)), progress,
                                  progress_user_data);
}

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_ITERATIVE_SOLVER_H_
