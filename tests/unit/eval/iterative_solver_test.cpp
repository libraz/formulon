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
#include <string>
#include <unordered_map>
#include <vector>

#include "eval/dep_graph.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
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

TEST(IterativeSolver, DivergingSequenceRunsToIterationLimit) {
  // Strictly-monotonic-increasing recurrence: x_{n+1} = 2 * x_n + 1.
  // Starting from Blank (treated as 0), the sequence is 1, 3, 7, 15, ...
  // Excel has no residual-growth cutoff, so this must consume the configured
  // iteration budget rather than being converted to #NUM! early.
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
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, 100U);
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_GT(final.as_number(), 0.0);
}

TEST(IterativeSolver, ValueKindFlipNeverConverges) {
  // Toggles between Text and Number every iteration. The kind change
  // forces `abs_delta = +infinity`, so no iteration ever reads as
  // converged. The solver consumes the configured iteration limit without
  // inferring divergence from the non-finite residual sequence.
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
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, opts.max_iterations);
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_DOUBLE_EQ(final.as_number(), 1.0);
}

TEST(IterativeSolver, TextSnapshotSurvivesArenaResetAndChunkFree) {
  // A's first Text result lives in the arena's initial chunk. B then forces a
  // new chunk, making the initial chunk an older chunk that the next reset
  // frees. The solver must copy A's bytes before B is evaluated; retaining
  // A's string_view as the previous baseline is an ASan-detectable use after
  // free on the second sweep. A changes only after the embedded NUL byte, so
  // truncating the owned copy at NUL would incorrectly converge one sweep
  // early.
  const std::string expected_a("A\0C", 3);
  const std::string expected_b("C\0D", 3);
  constexpr std::size_t kRuns = 32U;
  constexpr std::uint32_t kExpectedIterations = 3U;

  for (std::size_t run = 0U; run < kRuns; ++run) {
    Arena arena(/*initial_chunk_bytes=*/32);
    const std::string a_first_text("A\0B", 3);
    const std::string a_second_text("A\0C", 3);
    const std::string b_text("C\0D", 3);
    std::size_t a_calls = 0U;
    std::string committed_a;
    std::string committed_b;
    std::vector<CellNodeId> scc{kCellA, kCellB};
    IterativeOptions opts;
    opts.enabled = true;
    opts.max_iterations = 4U;
    opts.max_change = 0.001;

    auto eval_fn = [&](CellNodeId cell) {
      arena.reset();
      if (cell == kCellA) {
        ++a_calls;
        const std::string& source = (a_calls == 1U) ? a_first_text : a_second_text;
        const std::string_view bytes = arena.intern(std::string_view(source.data(), source.size()));
        return Value::text(bytes);
      }

      // B starts with the initial chunk reset to empty. This fills that chunk
      // exactly; the following Text allocation therefore gets a fresh head
      // chunk, and the next reset releases A's old chunk.
      if (arena.allocate(32U, alignof(char)) == nullptr) {
        ADD_FAILURE() << "the forced second arena chunk allocation failed";
        return Value::error(ErrorCode::Num);
      }
      const std::string_view bytes = arena.intern(std::string_view(b_text.data(), b_text.size()));
      return Value::text(bytes);
    };
    auto commit_fn = [&](CellNodeId cell, Value value) {
      ASSERT_TRUE(value.is_text());
      const std::string_view bytes = value.as_text();
      if (cell == kCellA) {
        committed_a.assign(bytes.data(), bytes.size());
      } else if (cell == kCellB) {
        committed_b.assign(bytes.data(), bytes.size());
      }
    };

    const IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn);
    EXPECT_TRUE(out.converged) << "run " << run;
    EXPECT_EQ(out.iterations_run, kExpectedIterations) << "run " << run;
    EXPECT_EQ(a_calls, 3U) << "run " << run;
    ASSERT_EQ(committed_a.size(), expected_a.size()) << "run " << run;
    ASSERT_EQ(committed_b.size(), expected_b.size()) << "run " << run;
    EXPECT_EQ(committed_a, expected_a) << "run " << run;
    EXPECT_EQ(committed_b, expected_b) << "run " << run;
  }
}

TEST(IterativeSolver, ArrayAndLambdaSnapshotsAlwaysRemainUnconverged) {
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 2U;
  opts.max_change = 0.001;
  const std::vector<CellNodeId> scc{kCellA};
  auto commit_fn = [](CellNodeId /*cell*/, Value /*value*/) {};

  const IterativeOutcome array_out =
      run_iterative_solve(scc, opts, [](CellNodeId /*cell*/) { return Value::array(nullptr); }, commit_fn);
  EXPECT_FALSE(array_out.converged);
  EXPECT_EQ(array_out.iterations_run, opts.max_iterations);

  const IterativeOutcome lambda_out =
      run_iterative_solve(scc, opts, [](CellNodeId /*cell*/) { return Value::lambda(nullptr); }, commit_fn);
  EXPECT_FALSE(lambda_out.converged);
  EXPECT_EQ(lambda_out.iterations_run, opts.max_iterations);
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

TEST(IterativeSolver, BlankResultConvergesOnFirstSweep) {
  // The legacy solver seeds every previous value with Blank. Preserve that
  // observable contract: a Blank-producing SCC converges in one sweep even
  // when the iteration budget is exactly one.
  IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 1U;
  opts.max_change = 0.001;
  const std::vector<CellNodeId> scc{kCellA};
  auto commit_fn = [](CellNodeId /*cell*/, Value /*value*/) {};

  const IterativeOutcome out =
      run_iterative_solve(scc, opts, [](CellNodeId /*cell*/) { return Value::blank(); }, commit_fn);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.diverged);
  EXPECT_EQ(out.iterations_run, 1U);
}

}  // namespace
}  // namespace formulon::eval
