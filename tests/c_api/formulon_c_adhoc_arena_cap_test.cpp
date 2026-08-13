//
// The ad-hoc evaluation entry points (`fm_workbook_evaluate_formula`,
// `fm_workbook_evaluate_formula_array`, `fm_workbook_evaluate_cf_formula`)
// each back one caller-supplied formula with a per-call arena. That arena
// must carry a finite byte ceiling so a hostile formula surfaces as a
// recoverable status instead of an allocation failure that aborts the host.
//
// Exhausting the production ceiling would cost a gigabyte of resident memory,
// which no fast-tier test can afford, so these tests pin the wiring rather
// than the exhaustion: the ceiling is finite, an arena built the way those
// entry points build theirs refuses a request past it without allocating,
// and ordinary evaluation still succeeds under the cap.

#include <cstddef>
#include <limits>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

}  // namespace

TEST(AdhocEvalArenaCap, EvalArenaCeilingIsFinite) {
  EXPECT_LT(formulon::kMaxEvalArenaBytes, std::numeric_limits<std::size_t>::max());
}

TEST(AdhocEvalArenaCap, ArenaBuiltLikeTheAdhocEntryPointsRefusesRequestPastCeiling) {
  // Same construction the ad-hoc entry points use. The oversized request is
  // rejected from the ceiling check, so no allocation of that size is ever
  // attempted.
  formulon::Arena arena(/*initial_chunk_bytes=*/4096, formulon::kMaxEvalArenaBytes);
  EXPECT_EQ(arena.allocate(formulon::kMaxEvalArenaBytes + 1U, alignof(double)), nullptr);
  EXPECT_TRUE(arena.exhausted());
  EXPECT_EQ(arena.bytes_allocated(), 0U);
}

TEST(AdhocEvalArenaCap, EvaluationStillSucceedsUnderTheCap) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 10.0), 0);  // A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t scalar{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 1, "=A1*2", &scalar), 0);
  EXPECT_EQ(scalar.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(scalar.u.number, 20.0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 1, "=SEQUENCE(3,2)", &rows, &cols), 0);
  EXPECT_EQ(rows, 3U);
  EXPECT_EQ(cols, 2U);

  fm_value_t fired{};
  ASSERT_EQ(fm_workbook_evaluate_cf_formula(wb.handle, 0, 0, 1, 0, 1, "=A1>5", &fired), 0);
  EXPECT_EQ(fired.kind, FM_VAL_BOOL);
  EXPECT_EQ(fired.u.boolean, 1);
}
