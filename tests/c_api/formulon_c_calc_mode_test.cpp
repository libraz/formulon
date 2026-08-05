//
// Stable C ABI calc-mode regression tests. Covers the
// `fm_workbook_calc_mode` / `fm_workbook_set_calc_mode` getter / setter
// pair, validation of the enum domain, and OOXML round-trip via
// `<calcPr calcMode>`.

#include <cstdint>
#include <cstring>

#include "c_api/formulon_c.h"
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

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

}  // namespace

TEST(FormulonCApiCalcMode, FreshWorkbookDefaultsToAuto) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_calc_mode_t mode = FM_CALC_MODE_MANUAL;  // poison
  ASSERT_EQ(fm_workbook_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

TEST(FormulonCApiCalcMode, SetThenGetReturnsRequestedMode) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_MANUAL);

  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO_NO_TABLE), 0);
  ASSERT_EQ(fm_workbook_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO_NO_TABLE);

  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO), 0);
  ASSERT_EQ(fm_workbook_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

TEST(FormulonCApiCalcMode, IterativeEnabledTogglePreservesConfiguredLimits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_iterative(wb.handle, 0, 42, 0.25), 0);
  ASSERT_EQ(fm_workbook_set_iterative_enabled(wb.handle, 1), 0);

  int32_t enabled = 0;
  uint32_t max_iterations = 0;
  double max_change = 0.0;
  ASSERT_EQ(fm_workbook_get_iterative(wb.handle, &enabled, &max_iterations, &max_change), 0);
  EXPECT_EQ(enabled, 1);
  EXPECT_EQ(max_iterations, 42U);
  EXPECT_DOUBLE_EQ(max_change, 0.25);
}

TEST(FormulonCApiCalcMode, UnknownModeRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Domain is {0, 1, 2}; build an out-of-range value via memcpy because a
  // direct `static_cast<fm_calc_mode_t>(99)` is unspecified per the standard
  // (99 sits outside the enum's representable range, which GCC's
  // -Wconversion correctly flags). The C ABI accepts the raw byte value
  // regardless and routes it through the validation switch.
  fm_calc_mode_t bad_mode{};
  const int raw = 99;
  std::memcpy(&bad_mode, &raw, sizeof(bad_mode));
  fm_status_t rc = fm_workbook_set_calc_mode(wb.handle, bad_mode);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  // Workbook state should be unchanged after a rejected set.
  fm_calc_mode_t mode = FM_CALC_MODE_MANUAL;
  ASSERT_EQ(fm_workbook_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

TEST(FormulonCApiCalcMode, ExcelProfileIdDefaultsToWinJaAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const char* profile = nullptr;
  ASSERT_EQ(fm_workbook_excel_profile_id(wb.handle, &profile), 0);
  ASSERT_STREQ(profile, "win-365-ja_JP");

  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_excel_profile_id(wb.handle, &profile), 0);
  EXPECT_STREQ(profile, "mac-365-ja_JP");

  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_excel_profile_id(wb.handle, &profile), 0);
  EXPECT_STREQ(profile, "win-365-ja_JP");
}

TEST(FormulonCApiCalcMode, UnknownExcelProfileIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_status_t rc = fm_workbook_set_excel_profile_id(wb.handle, "linux-365-ja_JP");
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  rc = fm_workbook_set_excel_profile_id(wb.handle, "mac-365-en_US");
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  const char* profile = nullptr;
  ASSERT_EQ(fm_workbook_excel_profile_id(wb.handle, &profile), 0);
  EXPECT_STREQ(profile, "win-365-ja_JP");
}

TEST(FormulonCApiCalcMode, NullArgsReturnBindingNullPointer) {
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  EXPECT_EQ(fm_workbook_calc_mode(nullptr, &mode),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_calc_mode(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_set_calc_mode(nullptr, FM_CALC_MODE_AUTO),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  const char* profile = nullptr;
  EXPECT_EQ(fm_workbook_excel_profile_id(nullptr, &profile),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_excel_profile_id(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_set_excel_profile_id(nullptr, "mac-365-ja_JP"),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_set_excel_profile_id(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiCalcMode, ManualModeRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=1+1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_MANUAL);
}

TEST(FormulonCApiCalcMode, AutoNoTableModeRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO_NO_TABLE), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO_NO_TABLE);
}

TEST(FormulonCApiCalcMode, IterativeOptionsRoundTripThroughCalcPr) {
  // The reader/writer parses iterate / iterateCount / iterateDelta on
  // the same `<calcPr>` element. Drive the full round-trip through the
  // `fm_workbook_set_iterative` setter and confirm a reload preserves
  // the iterative-on flag together with calcMode.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);
  ASSERT_EQ(fm_workbook_set_iterative(wb.handle, /*enabled=*/1, /*max_iterations=*/25, /*max_change=*/0.005), 0);
  // Cyclic refs would surface #REF! without iterative on; using a
  // literal here keeps the test focused on metadata, not recalc paths.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=42"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  fm_calc_mode_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_MANUAL);
  // Iterative attributes are observed indirectly: a cyclic A1=B1+1,
  // B1=A1 chain only converges when the loaded workbook still has
  // iterative=on. Attach those after load to verify the flag survived.
  ASSERT_EQ(fm_workbook_set_formula(wb2.handle, 0, 0, 0, "=B1+1"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb2.handle, 0, 0, 1, "=A1"), 0);
  // Reading back via the workbook object isn't part of the C ABI here;
  // instead drive a recalc and verify it does not surface #REF!. With
  // iterative=off the cycle would fail; with iterative=on (preserved
  // across the round-trip) the solver runs and returns kOk.
  EXPECT_EQ(fm_workbook_recalc(wb2.handle), 0);
}

TEST(FormulonCApiCalcMode, AutoModeOmitsCalcPrFromOoxml) {
  // Default calc mode + default iterative options means the writer
  // should NOT emit a `<calcPr>` element. We verify behavioural
  // equivalence: a fresh load must still report Auto.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // No set_calc_mode call → Auto by default.

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  fm_calc_mode_t mode = FM_CALC_MODE_MANUAL;  // poison
  ASSERT_EQ(fm_workbook_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}
