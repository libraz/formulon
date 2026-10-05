// Stable C ABI (`src/c_api/formulon_c.h`) end-to-end tests.
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface declared in
// `formulon_c.h`.

#include <atomic>
#include <cstddef>
#include <cstring>
#include <limits>
#include <string>
#include <thread>
#include <vector>

#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

TEST(FormulonCApi, CreateAndDestroy) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NE(wb.handle, nullptr);
  // create() always seeds a single Sheet1.
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_sheet_name(wb.handle, 0, &name), 0);
  ASSERT_NE(name, nullptr);
  EXPECT_STREQ(name, "Sheet1");
}
TEST(FormulonCApi, CreateEmptyAndAddSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 0U);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Data"), 0);
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Stats"), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 2U);

  const char* n0 = nullptr;
  const char* n1 = nullptr;
  ASSERT_EQ(fm_workbook_sheet_name(wb.handle, 0, &n0), 0);
  ASSERT_EQ(fm_workbook_sheet_name(wb.handle, 1, &n1), 0);
  EXPECT_STREQ(n0, "Data");
  EXPECT_STREQ(n1, "Stats");
}
TEST(FormulonCApi, AddSheetValidatesNameAndRejectsDuplicate) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Data"), 0);
  // Duplicate (case-insensitive), forbidden character, and empty name are
  // all rejected now that the public add surface shares the rename
  // validator instead of silently accepting anything.
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "data"), 0);
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "a/b"), 0);
  EXPECT_NE(fm_workbook_add_sheet(wb.handle, ""), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
  // A valid, distinct name still succeeds.
  EXPECT_EQ(fm_workbook_add_sheet(wb.handle, "Stats"), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 2U);
}
TEST(FormulonCApi, NumberLiteralRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 42.5), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 42.5);
}
TEST(FormulonCApi, GetValueRejectsCoordinatesOutsideTheGrid) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  fm_value_t v{};
  EXPECT_EQ(fm_workbook_get_value(wb.handle, 0, formulon::Sheet::kMaxRows, 0, &v), kInvalidArgument);
  EXPECT_EQ(fm_workbook_get_value(wb.handle, 0, 0, formulon::Sheet::kMaxCols, &v), kInvalidArgument);
  // The last addressable cell is still readable.
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, formulon::Sheet::kMaxRows - 1, formulon::Sheet::kMaxCols - 1, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);
}
TEST(FormulonCApi, SetNumberRejectsNaNAndInfinity) {
  // A Number-kind cell must always hold a finite double: ISNUMBER,
  // arithmetic on the cell, and a save/reload round trip would otherwise
  // disagree about whether it is a number or an error (save() downgrades a
  // non-finite literal to #NUM!, per io/ooxml_writer_cell.cpp).
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, std::numeric_limits<double>::quiet_NaN()), kInvalidArgument);
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, std::numeric_limits<double>::infinity()), kInvalidArgument);
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, -std::numeric_limits<double>::infinity()), kInvalidArgument);

  // A rejected call must not have written anything: the cell stays blank.
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);

  // A finite value still succeeds.
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.5), 0);
}
TEST(FormulonCApi, ParallelRecalcMatchesSerialOnWideIndependentDag) {
  WorkbookGuard serial;
  WorkbookGuard parallel;
  ASSERT_EQ(fm_workbook_create(&serial.handle), 0);
  ASSERT_EQ(fm_workbook_create(&parallel.handle), 0);

  constexpr std::uint32_t kRows = 48U;
  for (std::uint32_t row = 0U; row < kRows; ++row) {
    const double input = static_cast<double>(row + 1U);
    const std::string formula = "=A" + std::to_string(row + 1U) + "*2+1";
    ASSERT_EQ(fm_workbook_set_number(serial.handle, 0U, row, 0U, input), 0);
    ASSERT_EQ(fm_workbook_set_number(parallel.handle, 0U, row, 0U, input), 0);
    ASSERT_EQ(fm_workbook_set_formula(serial.handle, 0U, row, 1U, formula.c_str()), 0);
    ASSERT_EQ(fm_workbook_set_formula(parallel.handle, 0U, row, 1U, formula.c_str()), 0);
  }

  ASSERT_EQ(fm_workbook_recalc(serial.handle), 0);
  fm_parallel_recalc_stats stats{};
  ASSERT_EQ(fm_workbook_recalc_parallel(parallel.handle, 4U, &stats), 0);
  EXPECT_EQ(stats.cells_evaluated, kRows);
  EXPECT_GE(stats.sccs_processed, kRows);
  EXPECT_GE(stats.parallel_steps, 1U);
  EXPECT_GE(stats.worker_threads_started, 2U);
  EXPECT_GE(stats.worker_threads_used, 1U);

  for (std::uint32_t row = 0U; row < kRows; ++row) {
    fm_value_t serial_value{};
    fm_value_t parallel_value{};
    ASSERT_EQ(fm_workbook_get_value(serial.handle, 0U, row, 1U, &serial_value), 0);
    ASSERT_EQ(fm_workbook_get_value(parallel.handle, 0U, row, 1U, &parallel_value), 0);
    ASSERT_EQ(serial_value.kind, FM_VAL_NUMBER);
    ASSERT_EQ(parallel_value.kind, FM_VAL_NUMBER);
    EXPECT_DOUBLE_EQ(parallel_value.u.number, serial_value.u.number);
    EXPECT_DOUBLE_EQ(parallel_value.u.number, static_cast<double>(row + 1U) * 2.0 + 1.0);
  }
}
TEST(FormulonCApi, ParallelRecalcCountOneStartsNoWorkers) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0U, 0U, 0U, 7.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0U, 0U, 1U, "=A1+1"), 0);

  fm_parallel_recalc_stats stats{};
  ASSERT_EQ(fm_workbook_recalc_parallel(wb.handle, 1U, &stats), 0);
  EXPECT_EQ(stats.cells_evaluated, 1U);
  EXPECT_EQ(stats.worker_threads_started, 0U);
  EXPECT_EQ(stats.worker_threads_used, 0U);

  // The output is optional for callers that only need the status code.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0U, 0U, 0U, 9.0), 0);
  EXPECT_EQ(fm_workbook_recalc_parallel(wb.handle, 1U, nullptr), 0);
}
TEST(FormulonCApi, ParallelRecalcInvalidAndNullPathsResetStats) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_parallel_recalc_stats poison = {1U, 2U, 3U, 4U, 5U, 6U, 7U};
  const fm_parallel_recalc_stats zero{};
  fm_parallel_recalc_stats stats = poison;
  EXPECT_EQ(fm_workbook_recalc_parallel(wb.handle, 9U, &stats),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(std::memcmp(&stats, &zero, sizeof(stats)), 0);

  stats = poison;
  EXPECT_EQ(fm_workbook_recalc_parallel(nullptr, 1U, &stats),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(std::memcmp(&stats, &zero, sizeof(stats)), 0);

  EXPECT_EQ(fm_workbook_recalc_parallel(wb.handle, 1U, nullptr), 0);
}
TEST(FormulonCApi, BoolAndBlankSetters) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_bool(wb.handle, 0, 0, 0, 1), 0);
  ASSERT_EQ(fm_workbook_set_bool(wb.handle, 0, 1, 0, 0), 0);
  ASSERT_EQ(fm_workbook_set_blank(wb.handle, 0, 2, 0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 1);

  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 0);

  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 2, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);
}
TEST(FormulonCApi, ErrorSetterStoresStaticErrorLiteral) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_error(wb.handle, 0, 0, 0, 1), 0);  // ErrorCode::Div0
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
  EXPECT_EQ(v.u.error_code, 1);
}
TEST(FormulonCApi, ErrorSetterRejectsInvalidErrorCode) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_workbook_set_error(wb.handle, 0, 0, 0, -1), 0);
  EXPECT_NE(fm_workbook_set_error(wb.handle, 0, 0, 0, 999), 0);
}
TEST(FormulonCApi, FormulaNumericResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 10.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 1, 0, "=A1*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 1, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 20.0);
}
TEST(FormulonCApi, FormulaTextResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=UPPER(\"hello\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_TEXT);
  ASSERT_NE(v.u.text, nullptr);
  EXPECT_STREQ(v.u.text, "HELLO");
}
TEST(FormulonCApi, TextSetterRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Use a stack buffer that goes out of scope before recalc to confirm
  // the handle interns the bytes.
  {
    char tmp[] = "hello";
    ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, tmp), 0);
    // Mutate the source buffer to prove we're not aliasing.
    tmp[0] = 'X';
    (void)tmp[0];
  }
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_TEXT);
  ASSERT_NE(v.u.text, nullptr);
  EXPECT_STREQ(v.u.text, "hello");
  // Verify the returned pointer is actually NUL-terminated by inspecting
  // strlen against the documented length.
  EXPECT_EQ(std::strlen(v.u.text), 5U);
}
TEST(FormulonCApi, FormulaErrorSurfacesAsValueError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1/0"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_ERROR);
  // Excel's `#DIV/0!` is `ErrorCode::Div0`, ordinal 1.
  EXPECT_EQ(v.u.error_code, 1);
}
TEST(FormulonCApi, NullWorkbookSetsBindingError) {
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(nullptr, 0, 0, 0, &v);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_GT(std::strlen(fm_last_error_message()), 0U);

  // Also exercise the create-side NULL path.
  rc = fm_workbook_create(nullptr);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_GT(std::strlen(fm_last_error_message()), 0U);
}
TEST(FormulonCApi, OutOfRangeSheetIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_status_t rc = fm_workbook_set_number(wb.handle, 99, 0, 0, 1.0);
  EXPECT_NE(rc, 0);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  // Context should reference the offending sheet index.
  std::string ctx = fm_last_error_context();
  EXPECT_NE(ctx.find("sheet_index"), std::string::npos) << "context=" << ctx;
}
TEST(FormulonCApi, OutOfGridCoordinateIsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // A column near the top of the u32 range would otherwise resize a row
  // vector to billions of cells. It must be rejected before the storage
  // layer, not turned into a multi-GB allocation.
  const std::uint32_t kBadCol = 4'000'000'000U;
  fm_status_t rc = fm_workbook_set_number(wb.handle, 0, 0, kBadCol, 1.0);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  std::string ctx = fm_last_error_context();
  EXPECT_NE(ctx.find("col"), std::string::npos) << "context=" << ctx;

  // Row at the Excel ceiling (kMaxRows) is one past the last addressable
  // row and must also be rejected.
  EXPECT_EQ(fm_workbook_set_formula(wb.handle, 0, formulon::Sheet::kMaxRows, 0, "=1"),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_set_blank(wb.handle, 0, 0, formulon::Sheet::kMaxCols),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // The last in-grid cell still round-trips.
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, formulon::Sheet::kMaxRows - 1U, formulon::Sheet::kMaxCols - 1U, 3.0),
            0);
}
TEST(FormulonCApi, SuccessClearsPreviousLastError) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // First, force an error so the thread-local diagnostic is populated.
  fm_value_t v{};
  ASSERT_NE(fm_workbook_get_value(nullptr, 0, 0, 0, &v), 0);
  EXPECT_GT(std::strlen(fm_last_error_message()), 0U);
  // Now drive a success and confirm the diagnostic was cleared.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  EXPECT_EQ(std::strlen(fm_last_error_message()), 0U);
  EXPECT_EQ(std::strlen(fm_last_error_context()), 0U);
}

TEST(FormulonCApi, IterativeOptionsConverge) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_iterative(wb.handle, 1, 100, 1e-6), 0);

  // A1 = 0.5 * (A1 + 2) converges to 2 from the blank initial cache
  // value. Unlike Newton's method, this intentionally needs no separate
  // numeric seed, because setting a formula replaces that cell's cache.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=0.5*(A1+2)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_NEAR(v.u.number, 2.0, 1e-3);
}
