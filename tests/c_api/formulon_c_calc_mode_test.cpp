//
// Stable C ABI calc-mode regression tests. Covers the
// `fm_workbook_get_calc_mode` / `fm_workbook_set_calc_mode` getter / setter
// pair, validation of the enum domain, and OOXML round-trip via
// `<calcPr calcMode>`.

#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "support/ooxml_package_fixture.h"
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
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &wb2.handle), 0);
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
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &wb2.handle), 0);
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
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &wb2.handle), 0);
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
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &wb2.handle), 0);
  int32_t mode = FM_CALC_MODE_MANUAL;  // poison
  ASSERT_EQ(fm_workbook_get_calc_mode(wb2.handle, &mode), 0);
  EXPECT_EQ(mode, FM_CALC_MODE_AUTO);
}

namespace {

constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kReentrant = static_cast<fm_status_t>(formulon::FormulonErrorCode::kGraphRecalcReentrant);
constexpr fm_status_t kResourceLimit = static_cast<fm_status_t>(formulon::FormulonErrorCode::kSecurityResourceLimit);

// A1 = 1, B1 = A1+1 caching 99 and C1 = "a"&" " caching "a", in an automatic workbook with the current calcId
// (`full_calc_on_load` adds fullCalcOnLoad).
void LoadStaleCacheBook(WorkbookGuard* out, bool full_calc_on_load = false) {
  const std::string calc_pr =
      full_calc_on_load ? "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>" : "<calcPr calcId=\"191029\"/>";
  std::string workbook(formulon::test::kOneSheetWorkbookXml);
  workbook.insert(workbook.find("</workbook>"), calc_pr);
  const std::string sheet =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\">"
      "<c r=\"A1\"><v>1</v></c><c r=\"B1\"><f>A1+1</f><v>99</v></c>"
      "<c r=\"C1\" t=\"str\"><f>\"a\"&amp;\" \"</f><v>a</v></c></row></sheetData></worksheet>";
  const std::vector<std::uint8_t> package =
      formulon::test::BuildOoxmlPackage(formulon::test::OoxmlContentTypes(formulon::test::kXlsxWorkbookContentType),
                                        sheet, workbook, formulon::test::kOneSheetWorkbookRels);
  ASSERT_EQ(fm_workbook_load(package.data(), package.size(), nullptr, &out->handle), 0);
}

int32_t HasDirty(const fm_workbook_t* wb) {
  int32_t dirty = -1;
  EXPECT_EQ(fm_workbook_has_dirty_cells(wb, &dirty), 0);
  return dirty;
}

double NumberAt(const fm_workbook_t* wb, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, 0, col, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
  return v.kind == FM_VALUE_KIND_NUMBER ? v.u.number : -1.0;
}

std::string TextAt(const fm_workbook_t* wb, uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, 0, col, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_TEXT);
  return v.kind == FM_VALUE_KIND_TEXT && v.u.text != nullptr ? std::string(v.u.text) : std::string();
}

}  // namespace

TEST(FormulonCApiCalcMode, FullRecalcRecomputesTheValuesABookOpensWith) {
  for (const int32_t mode : {FM_CALC_MODE_AUTO, FM_CALC_MODE_MANUAL}) {
    SCOPED_TRACE(mode);
    WorkbookGuard wb;
    LoadStaleCacheBook(&wb);
    ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, mode), 0);
    EXPECT_EQ(HasDirty(wb.handle), 0);
    EXPECT_EQ(NumberAt(wb.handle, 1), 99.0);
    EXPECT_EQ(TextAt(wb.handle, 2), "a");
    ASSERT_EQ(fm_workbook_recalc_full(wb.handle), 0);
    EXPECT_EQ(NumberAt(wb.handle, 1), 2.0);
    EXPECT_EQ(TextAt(wb.handle, 2), "a ");
    EXPECT_EQ(HasDirty(wb.handle), 0);
  }
}

TEST(FormulonCApiCalcMode, FullCalcOnLoadBookOpensDirty) {
  WorkbookGuard wb;
  LoadStaleCacheBook(&wb, true);
  EXPECT_EQ(HasDirty(wb.handle), 1);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(NumberAt(wb.handle, 1), 2.0);
  EXPECT_EQ(TextAt(wb.handle, 2), "a ");
}

TEST(FormulonCApiCalcMode, ParallelFullRecalcReportsTheCellsItComputed) {
  WorkbookGuard wb;
  LoadStaleCacheBook(&wb);
  fm_parallel_recalc_stats_info_t stats = FM_STRUCT_INIT(fm_parallel_recalc_stats_info_t);
  ASSERT_EQ(fm_workbook_recalc_full_parallel(wb.handle, 1U, &stats), 0) << fm_get_last_error_message();
  EXPECT_EQ(stats.cells_evaluated, 2U);
  EXPECT_EQ(NumberAt(wb.handle, 1), 2.0);
  EXPECT_EQ(TextAt(wb.handle, 2), "a ");
  EXPECT_EQ(fm_workbook_recalc_full_parallel(wb.handle, 4U, nullptr), 0);
}

TEST(FormulonCApiCalcMode, FullRecalcRefusalsMarkNothingDirty) {
  EXPECT_EQ(fm_workbook_recalc_full(nullptr), kNullPointer);
  WorkbookGuard wb;
  LoadStaleCacheBook(&wb);
  const fm_parallel_recalc_stats_info_t poison = {sizeof(fm_parallel_recalc_stats_info_t), 1U, 2U, 3U, 4U, 5U, 6U, 7U};
  fm_parallel_recalc_stats_info_t stats = poison;
  EXPECT_EQ(fm_workbook_recalc_full_parallel(nullptr, 1U, &stats), kNullPointer);
  EXPECT_EQ(stats.cells_evaluated, 0U);
  stats = poison;
  EXPECT_EQ(fm_workbook_recalc_full_parallel(wb.handle, 9U, &stats), kResourceLimit);
  EXPECT_EQ(stats.cells_evaluated, 0U);
  stats = poison;
  stats.struct_size = 4U;
  EXPECT_EQ(fm_workbook_recalc_full_parallel(wb.handle, 1U, &stats), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(HasDirty(wb.handle), 0);
  EXPECT_EQ(NumberAt(wb.handle, 1), 99.0);
}

TEST(FormulonCApiCalcMode, FullRecalcFromAProgressCallbackIsReentrant) {
  struct Probe {
    fm_workbook_t* wb = nullptr;
    bool called = false;
    fm_status_t serial = 0;
    fm_status_t parallel = 0;
  } probe;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=B1+1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 1, "=A1/2"), 0);
  ASSERT_EQ(SetIterative(wb.handle, 1, 2, -1.0), 0);
  probe.wb = wb.handle;
  ASSERT_EQ(fm_workbook_set_iterative_progress(
                wb.handle,
                [](uint32_t, double, uint32_t, void* data) -> int32_t {
                  auto& p = *static_cast<Probe*>(data);
                  if (!p.called) {
                    p.called = true;
                    p.serial = fm_workbook_recalc_full(p.wb);
                    p.parallel = fm_workbook_recalc_full_parallel(p.wb, 1U, nullptr);
                  }
                  return 1;
                },
                &probe),
            0);
  EXPECT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_TRUE(probe.called);
  EXPECT_EQ(probe.serial, kReentrant);
  EXPECT_EQ(probe.parallel, kReentrant);
}
