//
// Implementation of `run_iterative_solve_impl`. See `iterative_solver.h`
// for the public contract and the design rationale.

#include "eval/iterative_solver.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <limits>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "eval/dep_graph.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon::eval {
namespace {

// A Value may borrow the evaluator's per-pass Arena. Keeping a borrowed Value
// as the previous-iteration baseline is therefore unsafe: the next member of
// an SCC can reset the arena and free that payload before the baseline is
// compared. ConvergenceSnapshot owns the only payloads that participate in
// convergence (scalar values and Text bytes) and intentionally treats all
// pointer-shaped kinds as unstable without retaining or dereferencing them.
// Text bytes are copied into a solver-owned, capped Arena; the view is valid
// only while the snapshot's corresponding Arena remains the previous/current
// sweep buffer.
struct ConvergenceSnapshot {
  ValueKind kind = ValueKind::Blank;
  // Scalar kinds are copied by value. This member is never populated for
  // Text, Array, Lambda, or Ref, so it cannot retain a borrowed pointer.
  Value scalar = Value::blank();
  // Text is copied with an explicit byte length so embedded NUL bytes remain
  // part of the convergence baseline. An allocation failure makes the
  // snapshot incomparable rather than allowing a false convergence result.
  std::string_view text;
  bool comparable = true;

  static ConvergenceSnapshot capture(const Value& value, Arena& arena) noexcept {
    ConvergenceSnapshot snapshot;
    snapshot.kind = value.kind();
    switch (snapshot.kind) {
      case ValueKind::Blank:
      case ValueKind::Number:
      case ValueKind::Bool:
      case ValueKind::Error:
        snapshot.scalar = value;
        break;
      case ValueKind::Text: {
        const std::string_view bytes = value.as_text();
        if (bytes.empty()) {
          break;
        }
        snapshot.text = arena.intern(bytes);
        if (snapshot.text.size() != bytes.size()) {
          snapshot.comparable = false;
        }
        break;
      }
      case ValueKind::Array:
      case ValueKind::Ref:
      case ValueKind::Lambda:
        // Never call an accessor for pointer-shaped payloads. Their shape,
        // contents, and pointer identity are all arena-lifetime dependent;
        // any such result must keep the residual at +infinity.
        snapshot.comparable = false;
        break;
    }
    return snapshot;
  }
};

// Returns the absolute delta between `prev` and `curr` for the purpose of
// the convergence test.
//
// Numeric -> Numeric is the only "small delta" case: we return
// `|curr.as_number() - prev.as_number()|`, treating NaN as +infinity so a
// formula that converged onto NaN never claims convergence (Excel surfaces
// NaN as `#NUM!` in practice, but we play it safe at the solver layer).
//
// Any kind change, an incomparable snapshot, or any non-numeric pair returns
// +infinity. Same-kind Text, Bool, Error, and Blank values compare their
// copied payloads. Array, Lambda, and Ref snapshots always return +infinity:
// no arena-backed payload is retained or dereferenced by the convergence
// checker.
double abs_delta(const ConvergenceSnapshot& prev, const ConvergenceSnapshot& curr) noexcept {
  if (!prev.comparable || !curr.comparable || prev.kind != curr.kind) {
    return std::numeric_limits<double>::infinity();
  }
  switch (curr.kind) {
    case ValueKind::Blank:
    case ValueKind::Bool:
    case ValueKind::Error:
      return (prev.scalar == curr.scalar) ? 0.0 : std::numeric_limits<double>::infinity();
    case ValueKind::Number: {
      const double prev_num = prev.scalar.as_number();
      const double curr_num = curr.scalar.as_number();
      if (std::isnan(prev_num) || std::isnan(curr_num)) {
        return std::numeric_limits<double>::infinity();
      }
      return std::fabs(curr_num - prev_num);
    }
    case ValueKind::Text:
      return (prev.text == curr.text) ? 0.0 : std::numeric_limits<double>::infinity();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return std::numeric_limits<double>::infinity();
  }
  return std::numeric_limits<double>::infinity();
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

  // Per-cell snapshots of the previous and current iterations. Seed the
  // previous map with comparable Blank values to preserve the legacy solver
  // contract: a Blank result can converge on the first sweep, while Number
  // and Text results still see the initial kind mismatch as +infinity. Text
  // views point into the matching Arena only:
  // `previous_snapshots` is never reset while it is the baseline, and the
  // two map/Arena pairs are swapped only after the current sweep finishes.
  // This bounds retained Text storage to at most two sweep buffers instead of
  // accumulating one heap string per cell per iteration.
  std::unordered_map<CellNodeId, ConvergenceSnapshot, CellNodeIdHash> previous_snapshots;
  std::unordered_map<CellNodeId, ConvergenceSnapshot, CellNodeIdHash> current_snapshots;
  previous_snapshots.reserve(scc.size());
  current_snapshots.reserve(scc.size());
  Arena previous_arena(/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes);
  Arena current_arena(/*initial_chunk_bytes=*/4096, kMaxEvalArenaBytes);
  for (CellNodeId cell : scc) {
    previous_snapshots.emplace(cell, ConvergenceSnapshot{});
  }

  for (std::uint32_t iter = 0U; iter < max_iters; ++iter) {
    // At this point current_arena/current_snapshots belong to the sweep that
    // is no longer the convergence baseline. Drop its views before resetting
    // the Arena; previous_arena/previous_snapshots still hold the complete
    // prior sweep.
    current_snapshots.clear();
    current_arena.reset();
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
      // Capture the result before committing it. `commit` may copy the value
      // into workbook storage or otherwise advance/reset the evaluator
      // Arena; the convergence check must only inspect the current sweep's
      // owned snapshot.
      ConvergenceSnapshot curr_snapshot = ConvergenceSnapshot::capture(curr, current_arena);
      auto previous_entry = previous_snapshots.find(cell);
      const double delta = (previous_entry != previous_snapshots.end())
                               ? abs_delta(previous_entry->second, curr_snapshot)
                               : std::numeric_limits<double>::infinity();
      if (delta > max_delta) {
        max_delta = delta;
      }
      // Commit so subsequent members in this iteration (and the next
      // iteration) see the new value.
      commit(cell, curr);
      current_snapshots.insert_or_assign(cell, std::move(curr_snapshot));
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

    // The current map's views now become the next iteration's baseline. The
    // old baseline is moved to the current buffer and will be reset only at
    // the top of the next sweep, after all of its views are no longer read.
    std::swap(previous_snapshots, current_snapshots);
    std::swap(previous_arena, current_arena);
  }

  // Excel has no residual-growth cutoff. Iteration-limit exhaustion leaves
  // the last commits intact; `converged` and `diverged` are both false to
  // signal that the finite budget was consumed.
  return out;
}

}  // namespace formulon::eval
