// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `run_iterative_solve_impl`. See `iterative_solver.h`
// for the public contract and the design rationale.

#include "eval/iterative_solver.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <limits>
#include <unordered_map>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "value.h"

namespace formulon::eval {
namespace {

// Returns the absolute delta between `prev` and `curr` for the purpose of
// the convergence test.
//
// Numeric -> Numeric is the only "small delta" case: we return
// `|curr.as_number() - prev.as_number()|`, treating NaN as +infinity so a
// formula that converged onto NaN never claims convergence (Excel surfaces
// NaN as `#NUM!` in practice, but we play it safe at the solver layer).
//
// Any kind change, or any non-numeric pair, returns +infinity. This makes
// "Number -> Text" or "Number -> Error" or "Text -> Text" never read as
// converged, which is the correct Excel-observable behaviour: iterative
// calc is fundamentally a numeric fixed-point search, and a kind flip
// means the formula has not stabilised. (A genuine `=42` cell does not
// participate in a cycle SCC in the first place, so this is not a
// regression for literal cells.)
double abs_delta(const Value& prev, const Value& curr) noexcept {
  if (prev.kind() != curr.kind()) {
    return std::numeric_limits<double>::infinity();
  }
  if (curr.kind() != ValueKind::Number) {
    // Same kind, non-numeric: equal -> 0, otherwise infinity. Equality
    // for the non-numeric kinds we care about (Text, Bool, Error) is
    // total and cheap.
    return (prev == curr) ? 0.0 : std::numeric_limits<double>::infinity();
  }
  const double prev_num = prev.as_number();
  const double curr_num = curr.as_number();
  if (std::isnan(prev_num) || std::isnan(curr_num)) {
    return std::numeric_limits<double>::infinity();
  }
  return std::fabs(curr_num - prev_num);
}

}  // namespace

IterativeOutcome run_iterative_solve_impl(const std::vector<CellNodeId>& scc, const IterativeOptions& opts,
                                          const std::function<Value(CellNodeId)>& evaluate_one,
                                          const std::function<void(CellNodeId, Value)>& commit,
                                          IterativeProgressCb progress, void* progress_user_data) {
  IterativeOutcome out;
  if (scc.empty()) {
    // Defensive: an empty component is degenerate. Treat as immediate
    // convergence so callers do not see a spurious failure flag, but
    // leave `iterations_run = 0` to make the no-op visible.
    out.converged = true;
    return out;
  }

  // Excel docs cap `max_iterations` at 32767 in the dialog. We accept the
  // full uint32 range; the only constraint we apply is "at least 1
  // iteration" so the loop body always runs once before the convergence
  // check inspects a delta.
  const std::uint32_t max_iters = (opts.max_iterations == 0U) ? 1U : opts.max_iterations;

  // Per-cell snapshot of the *previous* iteration's value. Initialised
  // from the per-cell `evaluate_one` snapshot caller is responsible for —
  // the caller passes a live evaluator, so the natural seed is the
  // current `cached_value`. We do not have direct access to that here, so
  // we rely on the protocol: the *first* iteration always reports an
  // infinite delta (because `prev_values` starts empty / Blank), which
  // means the first iteration is never confused for convergence.
  std::unordered_map<CellNodeId, Value, CellNodeIdHash> prev_values;
  prev_values.reserve(scc.size());
  for (CellNodeId cell : scc) {
    // Seed with Blank; `abs_delta(Blank, anything)` is +infinity for
    // numeric outputs (kind mismatch) which is exactly what we want for
    // the first iteration.
    prev_values.emplace(cell, Value::blank());
  }

  // Track the last three max-deltas for divergence detection. We only
  // arm divergence after we have at least three observed deltas, so the
  // sentinel `-1.0` reads as "no observation yet" without colliding with
  // any real delta (which is always non-negative or +infinity).
  double d_minus_2 = -1.0;
  double d_minus_1 = -1.0;

  for (std::uint32_t iter = 0U; iter < max_iters; ++iter) {
    double max_delta = 0.0;

    // Single Gauss-Seidel-style pass: evaluate each member in `scc` order
    // and commit its new value before evaluating the next. Excel's
    // observable behaviour is order-dependent on the natural-language
    // "iteration" anyway (it depends on traversal order), and Tarjan
    // gives us a stable order per SCC. The last-iteration commit means
    // each member sees its peers' freshest values during the same pass,
    // which is what Excel does when iterative calc is enabled.
    for (CellNodeId cell : scc) {
      Value curr = evaluate_one(cell);
      // Compute delta against the snapshot *before* committing the new
      // value, so the next member of the SCC observes the freshest
      // value but the convergence calculation still sees the inter-
      // iteration change. `prev_values` was seeded for every member in
      // `scc` above; `find()` is therefore guaranteed to hit, but we
      // guard defensively for SCCs with duplicate ids (a logic-bug
      // shape that should never reach here).
      auto entry = prev_values.find(cell);
      const double delta =
          (entry != prev_values.end()) ? abs_delta(entry->second, curr) : std::numeric_limits<double>::infinity();
      if (delta > max_delta) {
        max_delta = delta;
      }
      // Commit so subsequent members in this iteration (and the next
      // iteration) see the new value.
      commit(cell, curr);
      if (entry != prev_values.end()) {
        entry->second = curr;
      } else {
        prev_values.emplace(cell, curr);
      }
    }

    out.iterations_run = iter + 1U;

    // Optional progress callback: invoked AFTER the sweep so the caller
    // sees the residual that resulted from the work just performed and
    // BEFORE the convergence / divergence checks so the caller can
    // cancel a still-progressing solve. The callback returns `false` to
    // abort; we surface `aborted` and leave the cell store in its
    // current partially-converged state (the converged-or-not
    // accounting belongs to the caller).
    if (progress != nullptr) {
      const bool keep_going = progress(out.iterations_run, max_delta, max_iters, progress_user_data);
      if (!keep_going) {
        out.aborted = true;
        return out;
      }
    }

    // Convergence: the largest absolute change in this pass dropped
    // below the user-specified threshold. Strict-less-than mirrors the
    // Excel spec (`max_change` is a hard upper bound, not inclusive).
    if (max_delta < opts.max_change) {
      out.converged = true;
      return out;
    }

    // Divergence: three successive passes with non-decreasing max-delta
    // and a max-delta that is itself above the convergence threshold
    // (otherwise we are oscillating below the floor and the next test
    // would have converged). Requires at least three observed deltas
    // (`d_minus_2` and `d_minus_1` armed) before the test arms.
    if (d_minus_2 >= 0.0 && d_minus_1 >= 0.0 && max_delta >= d_minus_1 && d_minus_1 >= d_minus_2 &&
        max_delta > opts.max_change) {
      // Bail out early and write `#NUM!` to every member. Excel's
      // observable behaviour for "iterative calc could not converge" is
      // a divergence-flagged surface; `#NUM!` matches the spec corpus
      // entry in §2.7.3 and is the closest sentinel we have for
      // "numerical method failed".
      for (CellNodeId cell : scc) {
        commit(cell, Value::error(ErrorCode::Num));
      }
      out.diverged = true;
      return out;
    }

    // Slide the window for divergence detection.
    d_minus_2 = d_minus_1;
    d_minus_1 = max_delta;
  }

  // Iteration limit hit without convergence or divergence detection.
  // Cell values stay at the last-iteration commits; `converged` and
  // `diverged` are both false to signal "ran out of budget".
  return out;
}

}  // namespace formulon::eval
