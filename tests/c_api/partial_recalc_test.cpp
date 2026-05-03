// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Smoke tests for `fm_workbook_partial_recalc` and
// `fm_workbook_set_iterative_progress`. The tests exercise the C ABI
// surface only; the deeper engine semantics live in the unit-test
// suite (`tests/unit/eval/partial_recalc_test.cpp`,
// `tests/unit/eval/iterative_progress_test.cpp`).

#include "c_api/formulon_c.h"

#include <atomic>
#include <cstdint>
#include <cstring>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

std::atomic<int> g_callback_invocations{0};

extern "C" bool always_continue(uint32_t /*iteration*/, double /*max_residual*/, uint32_t /*max_iterations*/,
                                void* /*user_data*/) {
  g_callback_invocations.fetch_add(1, std::memory_order_relaxed);
  return true;
}

extern "C" bool always_abort(uint32_t /*iteration*/, double /*max_residual*/, uint32_t /*max_iterations*/,
                             void* /*user_data*/) {
  g_callback_invocations.fetch_add(1, std::memory_order_relaxed);
  return false;
}

}  // namespace

TEST(FormulonCApiPartialRecalc, RecomputesOnlyClosure) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 5.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 1, "=A1+1"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 2, "=A1*3"), 0);

  // Viewport: B1 only (sheet 0, row 0, col 1).
  fm_viewport vp{};
  vp.sheet = 0;
  vp.first_row = 0;
  vp.last_row = 0;
  vp.first_col = 1;
  vp.last_col = 1;
  uint32_t recomputed = 0;
  ASSERT_EQ(fm_workbook_partial_recalc(wb.handle, &vp, &recomputed), 0);
  // Only B1 was a formula in the closure.
  EXPECT_EQ(recomputed, 1U);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 6.0);
  // C1 was not in the closure: still blank.
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 2, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);
}

TEST(FormulonCApiPartialRecalc, NullArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_viewport vp{};
  EXPECT_EQ(fm_workbook_partial_recalc(nullptr, &vp, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_partial_recalc(wb.handle, nullptr, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPartialRecalc, OutCountOptional) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=42"), 0);
  fm_viewport vp{};
  vp.sheet = 0;
  vp.first_row = 0;
  vp.last_row = 0;
  vp.first_col = 0;
  vp.last_col = 0;
  // Passing NULL for `out_recomputed_count` is allowed.
  EXPECT_EQ(fm_workbook_partial_recalc(wb.handle, &vp, nullptr), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 42.0);
}

TEST(FormulonCApiPartialRecalc, IterativeProgressContinueRunsToConvergence) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Enable iterative calc.
  ASSERT_EQ(fm_workbook_set_iterative(wb.handle, 1, 100, 0.001), 0);
  // Build a self-referential cell: A1 = (A1 + 10) / 2 → fixed point 10.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=(A1+10)/2"), 0);

  g_callback_invocations.store(0);
  ASSERT_EQ(fm_workbook_set_iterative_progress(wb.handle, &always_continue, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_GT(g_callback_invocations.load(), 0);
  // Cell should converge near 10.
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_NEAR(v.u.number, 10.0, 0.01);
}

TEST(FormulonCApiPartialRecalc, IterativeProgressAbortStopsEarly) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_iterative(wb.handle, 1, 100, 0.001), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=(A1+10)/2"), 0);

  g_callback_invocations.store(0);
  ASSERT_EQ(fm_workbook_set_iterative_progress(wb.handle, &always_abort, nullptr), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  // The abort callback fires after iteration 1 — the solver stops
  // immediately. The cell holds the partial first-iteration value
  // (NOT #NUM!), and the callback was invoked exactly once.
  EXPECT_EQ(g_callback_invocations.load(), 1);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  // Partial (not converged, not error).
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
}

TEST(FormulonCApiPartialRecalc, IterativeProgressNullArgumentRejected) {
  EXPECT_EQ(fm_workbook_set_iterative_progress(nullptr, nullptr, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}
