//
// Unit tests for the iterative-solver progress callback contract added
// to `iterative_solver.h`. Drives the solver directly with a mock
// store so we can observe every callback invocation.

#include <cmath>
#include <cstdint>
#include <unordered_map>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon::eval {
namespace {

const CellNodeId kCellA{0U, 0U, 0U};

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

// Captured per-iteration data for assertions in the tests below.
struct ProgressCapture {
  std::vector<std::uint32_t> iterations;
  std::vector<double> residuals;
  std::vector<std::uint32_t> caps;
  std::uint32_t abort_after = 0;  // 0 means "never abort"
};

extern "C" bool capture_progress(std::uint32_t iteration, double max_residual, std::uint32_t max_iterations,
                                 void* user_data) {
  auto* cap = static_cast<ProgressCapture*>(user_data);
  cap->iterations.push_back(iteration);
  cap->residuals.push_back(max_residual);
  cap->caps.push_back(max_iterations);
  if (cap->abort_after != 0 && iteration >= cap->abort_after) {
    return false;
  }
  return true;
}

TEST(IterativeProgress, CallbackInvokedOncePerSweep) {
  // Simple converging recurrence: x_{n+1} = (x_n + 10) / 2.
  // Without an abort signal the callback should be invoked exactly
  // `iterations_run` times — one per Gauss-Seidel sweep.
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

  ProgressCapture cap;
  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn, &capture_progress, &cap);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.aborted);
  ASSERT_EQ(cap.iterations.size(), out.iterations_run);
  // Iteration counter should be 1-based and monotonically increasing.
  for (std::size_t i = 0; i < cap.iterations.size(); ++i) {
    EXPECT_EQ(cap.iterations[i], i + 1U);
    EXPECT_EQ(cap.caps[i], opts.max_iterations);
  }
}

TEST(IterativeProgress, ResidualDecreasesOnConvergingCase) {
  // For the same converging recurrence, the residual sequence after
  // the first iteration should be strictly decreasing (the seed
  // iteration's residual is +infinity from a Blank prior, so we skip
  // index 0 in the comparison).
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

  ProgressCapture cap;
  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn, &capture_progress, &cap);
  ASSERT_TRUE(out.converged);
  ASSERT_GE(cap.residuals.size(), 2U);
  // Skip index 0 (Blank -> Number kind change ⇒ +infinity).
  for (std::size_t i = 2; i < cap.residuals.size(); ++i) {
    EXPECT_LT(cap.residuals[i], cap.residuals[i - 1]) << "non-monotonic residual at index " << i;
  }
}

TEST(IterativeProgress, AbortReturnsAbortedOutcome) {
  // Configure the callback to return `false` after iteration 3. The
  // solver must stop, surface `aborted`, and leave the cell value at
  // its last-iteration commit (NOT #NUM!).
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

  ProgressCapture cap;
  cap.abort_after = 3U;
  IterativeOutcome out = run_iterative_solve(scc, opts, eval_fn, commit_fn, &capture_progress, &cap);
  EXPECT_FALSE(out.converged);
  EXPECT_TRUE(out.aborted);
  EXPECT_EQ(out.iterations_run, 3U);
  // The last commit before the abort wrote a partial value; the cell
  // must NOT be #NUM!.
  Value final = store.get(kCellA);
  ASSERT_TRUE(final.is_number());
  EXPECT_GT(final.as_number(), 0.0);
}

TEST(IterativeProgress, NullCallbackPreservesLegacyContract) {
  // Passing `nullptr` for the progress callback should leave behaviour
  // identical to the legacy callback-less overload — the solver must
  // converge as before with `aborted == false`.
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

  IterativeOutcome out =
      run_iterative_solve(scc, opts, eval_fn, commit_fn, /*progress=*/nullptr, /*progress_user_data=*/nullptr);
  EXPECT_TRUE(out.converged);
  EXPECT_FALSE(out.aborted);
}

}  // namespace
}  // namespace formulon::eval
