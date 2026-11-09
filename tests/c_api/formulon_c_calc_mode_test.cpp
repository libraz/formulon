//
// Stable C ABI calc-mode regression tests. Covers the
// `fm_workbook_get_calc_mode` / `fm_workbook_set_calc_mode` getter / setter
// pair, validation of the enum domain, and OOXML round-trip via
// `<calcPr calcMode>`.

#include <cstdint>
#include <limits>

#include "c_api/formulon_c.h"
#include "formulon_c_setters.h"
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
  ~BufferGuard() { fm_buffer_destroy(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

fm_iterative_settings_info_t ReadIterative(const fm_workbook_t* wb) {
  fm_iterative_settings_info_t info = FM_STRUCT_INIT(fm_iterative_settings_info_t);
  EXPECT_EQ(fm_workbook_get_iterative(wb, &info), 0);
  return info;
}

}  // namespace

TEST(FormulonCApiCalcMode, FreshWorkbookDefaultsToAuto) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  int32_t mode = FM_CALC_MODE_MANUAL;  // poison
  ASSERT_EQ(fm_workbook_get_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

TEST(FormulonCApiCalcMode, SetThenGetReturnsRequestedMode) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);
  int32_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_get_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_MANUAL);

  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO_NO_TABLE), 0);
  ASSERT_EQ(fm_workbook_get_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO_NO_TABLE);

  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO), 0);
  ASSERT_EQ(fm_workbook_get_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

TEST(FormulonCApiCalcMode, IterativeEnabledTogglePreservesConfiguredLimits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetIterative(wb.handle, 0, 42, 0.25), 0);
  ASSERT_EQ(SetIterative(wb.handle, 1, -1, -1.0), 0);

  const fm_iterative_settings_info_t info = ReadIterative(wb.handle);
  EXPECT_EQ(info.enabled, 1);
  EXPECT_EQ(info.max_iterations, 42U);
  EXPECT_DOUBLE_EQ(info.max_change, 0.25);
}

// The reader clamps `iterateCount` on its own path; the setter a host reaches
// without any file refuses a count above Excel's limit and keeps the stored one.
TEST(FormulonCApiCalcMode, IterativeMaxIterationsAboveExcelsLimitIsRefused) {
  constexpr uint32_t kExcelLimit = 32767U;
  const std::int32_t oversized[] = {
      32768,
      100000,
      std::numeric_limits<std::int32_t>::max(),
  };
  for (const std::int32_t requested : oversized) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const std::uint32_t before = ReadIterative(wb.handle).max_iterations;
    EXPECT_EQ(SetIterative(wb.handle, 1, requested, 0.001), FM_STATUS_INVALID_ARGUMENT) << requested;
    EXPECT_STREQ(fm_get_last_error_message(), "max_iterations must be in 1..32767") << requested;
    EXPECT_EQ(ReadIterative(wb.handle).max_iterations, before) << requested;
  }
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(SetIterative(wb.handle, 1, static_cast<std::int32_t>(kExcelLimit), 0.001), 0);
  EXPECT_EQ(ReadIterative(wb.handle).max_iterations, kExcelLimit);
}

// Both ends of the accepted range are stored unchanged.
TEST(FormulonCApiCalcMode, IterativeMaxIterationsKeepsInRangeValues) {
  constexpr uint32_t kExcelLimit = 32767U;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  ASSERT_EQ(SetIterative(wb.handle, 1, static_cast<std::int32_t>(kExcelLimit), 0.001), 0);
  EXPECT_EQ(ReadIterative(wb.handle).max_iterations, kExcelLimit);

  ASSERT_EQ(SetIterative(wb.handle, 1, 1, 0.001), 0);
  EXPECT_EQ(ReadIterative(wb.handle).max_iterations, 1U);
}

TEST(FormulonCApiCalcMode, UnknownModeRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // The setter accepts a raw int32_t so FFI callers can probe the complete
  // invalid domain without constructing an out-of-domain C enum.
  const std::int32_t invalid_modes[] = {
      99,
      std::numeric_limits<std::int32_t>::min(),
      std::numeric_limits<std::int32_t>::max(),
  };
  for (const std::int32_t raw : invalid_modes) {
    const fm_status_t rc = fm_workbook_set_calc_mode(wb.handle, raw);
    EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument)) << raw;
  }
  // Workbook state should be unchanged after a rejected set.
  int32_t mode = FM_CALC_MODE_MANUAL;
  ASSERT_EQ(fm_workbook_get_calc_mode(wb.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

TEST(FormulonCApiCalcMode, ExcelProfileIdDefaultsToWinEnAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  const char* profile = nullptr;
  ASSERT_EQ(fm_workbook_get_profile_id(wb.handle, &profile), 0);
  ASSERT_STREQ(profile, "win-365-en_US");

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_get_profile_id(wb.handle, &profile), 0);
  EXPECT_STREQ(profile, "mac-365-ja_JP");

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_get_profile_id(wb.handle, &profile), 0);
  EXPECT_STREQ(profile, "win-365-ja_JP");
}

TEST(FormulonCApiCalcMode, ExcelProfileSwitchDirtiesProfileSensitiveFormulas) {
  // CODE() special-cases U+9AD9 to 38526 only under the win-365-ja_JP host
  // (eval/info_lazy.cpp); recalc alone is dirty-only, so a profile switch
  // has to mark every formula dirty itself or the cached win-profile value
  // survives the switch unrecomputed.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=CODE(\"\xe9\xab\x99\")"), 0);  // U+9AD9
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  ASSERT_DOUBLE_EQ(v.u.number, 38526.0);

  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &v), 0);
  ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_NE(v.u.number, 38526.0) << "the profile switch must dirty CODE() so it re-evaluates under the new host";
}

TEST(FormulonCApiCalcMode, UnknownExcelProfileIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);

  fm_status_t rc = fm_workbook_set_profile_id(wb.handle, "linux-365-ja_JP");
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  rc = fm_workbook_set_profile_id(wb.handle, "mac-365-pt_PT");
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  const char* profile = nullptr;
  ASSERT_EQ(fm_workbook_get_profile_id(wb.handle, &profile), 0);
  EXPECT_STREQ(profile, "win-365-en_US");
}

TEST(FormulonCApiCalcMode, NullArgsReturnBindingNullPointer) {
  int32_t mode = FM_CALC_MODE_AUTO;
  EXPECT_EQ(fm_workbook_get_calc_mode(nullptr, &mode),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_get_calc_mode(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_set_calc_mode(nullptr, FM_CALC_MODE_AUTO),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  const char* profile = nullptr;
  EXPECT_EQ(fm_workbook_get_profile_id(nullptr, &profile),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_get_profile_id(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_set_profile_id(nullptr, "mac-365-ja_JP"),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_set_profile_id(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiCalcMode, ManualModeRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=1+1"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  int32_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_get_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_MANUAL);
}

TEST(FormulonCApiCalcMode, AutoNoTableModeRoundTripsThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO_NO_TABLE), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  int32_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_get_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO_NO_TABLE);
}

TEST(FormulonCApiCalcMode, IterativeOptionsRoundTripThroughCalcPr) {
  // The reader/writer parses iterate / iterateCount / iterateDelta on
  // the same `<calcPr>` element. Drive the full round-trip through the
  // `fm_workbook_update_iterative` setter and confirm a reload preserves
  // the iterative-on flag together with calcMode.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);
  ASSERT_EQ(SetIterative(wb.handle, /*enabled=*/1, /*max_iterations=*/25, /*max_change=*/0.005), 0);
  // Cyclic refs would surface #REF! without iterative on; using a
  // literal here keeps the test focused on metadata, not recalc paths.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=42"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  int32_t mode = FM_CALC_MODE_AUTO;
  ASSERT_EQ(fm_workbook_get_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_MANUAL);
  // Iterative attributes are observed indirectly: a cyclic A1=B1+1,
  // B1=A1 chain only converges when the loaded workbook still has
  // iterative=on. Attach those after load to verify the flag survived.
  ASSERT_EQ(fm_sheet_set_formula(wb2.handle, 0, 0, 0, "=B1+1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb2.handle, 0, 0, 1, "=A1"), 0);
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // No set_calc_mode call → Auto by default.

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  int32_t mode = FM_CALC_MODE_MANUAL;  // poison
  ASSERT_EQ(fm_workbook_get_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}
