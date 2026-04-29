// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the iterative-calc solver. The tests drive the solver
// directly with mock `evaluate_one` / `commit` callbacks rather than
// going through `RecalcEngine`; that keeps the tests focused on the
// fixed-point iteration logic itself (convergence detection, divergence
// detection, value-kind handling, multi-cell SCCs) without dragging in
// parser / evaluator state.

#include "eval/iterative_solver.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <vector>

#include "eval/dep_graph.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon::eval {
namespace {

// Convenience: a single-cell SCC at sheet 0, row 0, col 0.
const CellNodeId kCellA{0U, 0U, 0U};
const CellNodeId kCellB{0U, 1U, 0U};

// Mock store: maps a CellNodeId to its current `cached_value`. Used as
// the substrate that `evaluate_one` reads and `commit` writes to.
class MockStore {
 public:
  Value get(CellNodeId c) const {
    auto it = store_.find(c);
    if (it == store_.end()) {
      return Value::blank();
    }
    return it->second;
  }
  void set(CellNodeId c, Value v) {
    auto it = store_.find(c);
    if (it == store_.end()) {
      store_.emplace(c, v);
    } else {
      it->second = v;
    }
  }

 private:
  std::unordered_map<CellNodeId, Value, CellNodeIdHash> store_;
};

TEST(IterativeSolver, SimpleFixedPointConverges) {
  // Recurrence: x_{n+1} = (x_n + 10) / 2.
  // Fixed point: x = 10. Starting from blank (treated as 0), the sequence
  // converges geometrically toward 10. With max_change = 0.001 we expect
  // convergence within ~14 iterations.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 100U;
  opts.max_change = 0.001;

  std::vector<CellNodeId> scc{kCellA};
  auto eval_fn = [&](CellNodeId c) {
    Value prev = store.get(c);
    double base = prev.is_number() ? prev.as_number() : 0.0;
    return Value::number((base + 10.0) / 2.0);
  };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_GT(out.iterations_run, 0U);
  EXPECT_LT(out.iterations_run, 100U);
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_NEAR(final.as_number(), 10.0, 0.01);
}

TEST(IterativeSolver, ImmediateConvergenceAfterSinglePass) {
  // Constant function `=42` against a Blank prior. The first pass
  // commits 42 (kind change Blank -> Number, delta = +inf). The second
  // pass commits 42 again (delta = 0). Converges at iteration 2.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 100U;
  opts.max_change = 0.001;

  std::vector<CellNodeId> scc{kCellA};
  auto eval_fn = [&](CellNodeId /*c*/) { return Value::number(42.0); };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, 2U);
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_DOUBLE_EQ(final.as_number(), 42.0);
}

TEST(IterativeSolver, IterationLimitExhaustedWithoutConvergence) {
  // Tight max_iterations + large max_change == 0 forces the solver to
  // run for the full budget. We use a slow-converging recurrence
  // (x_{n+1} = (x_n + 1000) / 2) and cap iterations at 5; the sequence
  // will not converge within 5 steps to a max_change threshold of 1e-9.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 5U;
  opts.max_change = 1e-9;

  std::vector<CellNodeId> scc{kCellA};
  auto eval_fn = [&](CellNodeId c) {
    Value prev = store.get(c);
    double base = prev.is_number() ? prev.as_number() : 0.0;
    return Value::number((base + 1000.0) / 2.0);
  };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_FALSE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, 5U);
  // Cell still holds the last-iteration value (not #NUM!): the solver
  // does not overwrite on iteration-limit exhaustion; the recalc engine
  // is the one that decides the user-visible failure mode.
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_GT(final.as_number(), 0.0);
}

TEST(IterativeSolver, DivergenceTriggersNumError) {
  // Strictly-monotonic-increasing recurrence: x_{n+1} = 2 * x_n + 1.
  // Starting from Blank (treated as 0), the sequence is 1, 3, 7, 15, ...
  // Each iteration's |delta| is strictly larger than the previous, so
  // after three observed deltas the divergence detector fires.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 100U;
  opts.max_change = 0.001;

  std::vector<CellNodeId> scc{kCellA};
  auto eval_fn = [&](CellNodeId c) {
    Value prev = store.get(c);
    double base = prev.is_number() ? prev.as_number() : 0.0;
    return Value::number((2.0 * base) + 1.0);
  };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_FALSE(out.converged);
  EXPECT_TRUE(out.diverged);
  // Divergence detection arms once we have d_{n-2}, d_{n-1}, d_n. With
  // the sequence 1, 3, 7, 15, ... the deltas are 1, 2, 4, 8, ... and the
  // monotonic-increase test fires on the third comparison (iteration 4
  // or earlier depending on the seed transition).
  EXPECT_LE(out.iterations_run, 100U);
  // On divergence the solver writes #NUM! to every member.
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_error());
  EXPECT_EQ(final.as_error(), ErrorCode::Num);
}

TEST(IterativeSolver, ValueKindFlipNeverConverges) {
  // Toggles between Text and Number every iteration. The kind change
  // forces `abs_delta = +infinity`, so no iteration ever reads as
  // converged. The solver bails out on iteration limit; `diverged` is
  // false because the deltas are constant (+infinity, +infinity, ...) —
  // the divergence test requires `>=` on monotonic non-decreasing
  // deltas, which `+inf >= +inf` satisfies, so divergence DOES fire.
  // Exact behaviour depends on the divergence threshold check; the
  // test pins the observable result.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 20U;
  opts.max_change = 0.001;

  std::vector<CellNodeId> scc{kCellA};
  bool toggle = false;
  auto eval_fn = [&](CellNodeId /*c*/) {
    toggle = !toggle;
    if (toggle) {
      return Value::text("hello");
    }
    return Value::number(1.0);
  };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_FALSE(out.converged);
  // The solver detects the constant +infinity delta sequence as
  // divergent (3 successive non-decreasing deltas above max_change).
  EXPECT_TRUE(out.diverged);
  // On divergence the cell holds #NUM!.
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_error());
  EXPECT_EQ(final.as_error(), ErrorCode::Num);
}

TEST(IterativeSolver, MultiCellSccConverges) {
  // Two-cell SCC where each cell mirrors the other:
  //   A = B (read previous B)
  //   B = A (read previous A — but A was just committed in the same
  //          pass, since the solver is Gauss-Seidel-style).
  // Seeding both with Blank, the first pass commits A = Blank, B = Blank
  // (B reads the just-committed A which is Blank). Both stay Blank
  // forever; the kind-equal-Blank branch in `abs_delta` returns 0, so
  // we converge on iteration 2.
  //
  // To exercise an actual numeric handshake, we instead model:
  //   A = B + 1 (uses prior B, which is Blank=0 -> 1)
  //   B = A - 1 (uses just-committed A=1 -> 0)
  // First iter commits A=1, B=0. Second iter: A = B + 1 = 1, B = A - 1 = 0.
  // Delta is 0 across the second pass -> converged.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 50U;
  opts.max_change = 0.001;

  std::vector<CellNodeId> scc{kCellA, kCellB};
  auto eval_fn = [&](CellNodeId c) {
    if (c == kCellA) {
      Value b = store.get(kCellB);
      double bv = b.is_number() ? b.as_number() : 0.0;
      return Value::number(bv + 1.0);
    }
    Value a = store.get(kCellA);
    double av = a.is_number() ? a.as_number() : 0.0;
    return Value::number(av - 1.0);
  };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_LE(out.iterations_run, 3U);
  Value a = store.get(kCellA);
  Value b = store.get(kCellB);
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), 1.0);
  EXPECT_DOUBLE_EQ(b.as_number(), 0.0);
}

TEST(IterativeSolver, EmptySccIsNoop) {
  // Defensive: passing an empty component returns immediate
  // "convergence" with zero iterations run. Caller is responsible for
  // never invoking the solver on empty input, but the contract is
  // documented and tested.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  std::vector<CellNodeId> scc;
  bool eval_called = false;
  bool commit_called = false;
  auto eval_fn = [&](CellNodeId /*c*/) {
    eval_called = true;
    return Value::blank();
  };
  auto commit_fn = [&](CellNodeId /*c*/, Value /*v*/) { commit_called = true; };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, 0U);
  EXPECT_FALSE(eval_called);
  EXPECT_FALSE(commit_called);
}

TEST(IterativeSolver, MaxIterationsZeroTreatedAsOne) {
  // `max_iterations = 0` is treated as `1` so the loop body always
  // runs once. The sole iteration commits the new value; whether it
  // converges depends on whether the kind matches the seed (Blank).
  // For a numeric output the kind change registers as +infinity delta,
  // so iter 1 ends without convergence -> iteration limit hit ->
  // !converged && !diverged.
  MockStore store;
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 0U;
  opts.max_change = 0.001;

  std::vector<CellNodeId> scc{kCellA};
  auto eval_fn = [&](CellNodeId /*c*/) { return Value::number(7.0); };
  auto commit_fn = [&](CellNodeId c, Value v) { store.set(c, v); };

  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
  EXPECT_FALSE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, 1U);
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_DOUBLE_EQ(final.as_number(), 7.0);
}

}  // namespace
}  // namespace formulon::eval
