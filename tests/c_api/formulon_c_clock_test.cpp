//
// Stable C ABI clock-seam regression tests. Covers the
// `fm_workbook_pinned_now` / `fm_workbook_set_pinned_now` /
// `fm_workbook_clear_pinned_now` trio, the civil-field validation domain,
// the effect of a pin on `NOW` / `TODAY`, and the fact that a pin is model
// state a save does not carry.

#include <cstdint>

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

// 2026-04-23 15:30:45 local. The serial and the time-of-day fraction are the
// same constants the engine-level clock-seam tests assert against.
constexpr fm_civil_time_t kPinned{2026, 4, 23, 15, 30, 45};
constexpr double kPinnedSerial = 46135.0;

fm_civil_time_t WithField(fm_civil_time_t base, int index, int32_t value) {
  switch (index) {
    case 0:
      base.year = value;
      break;
    case 1:
      base.month = value;
      break;
    case 2:
      base.day = value;
      break;
    case 3:
      base.hour = value;
      break;
    case 4:
      base.minute = value;
      break;
    default:
      base.second = value;
      break;
  }
  return base;
}

}  // namespace

TEST(FormulonCApiClock, FreshWorkbookFollowsTheHostClock) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_civil_time_t now{1, 1, 1, 1, 1, 1};  // poison
  int32_t pinned = 7;                     // poison
  ASSERT_EQ(fm_workbook_pinned_now(wb.handle, &now, &pinned), 0);
  EXPECT_EQ(pinned, 0);
  // Zero-filled rather than left as the caller wrote it, so a caller that
  // ignores `out_pinned` at least reads a value that cannot pass for a date.
  EXPECT_EQ(now.year, 0);
  EXPECT_EQ(now.month, 0);
  EXPECT_EQ(now.day, 0);
  EXPECT_EQ(now.hour, 0);
  EXPECT_EQ(now.minute, 0);
  EXPECT_EQ(now.second, 0);
}

TEST(FormulonCApiClock, APinReadsBackFieldForField) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &kPinned), 0);

  fm_civil_time_t now{};
  int32_t pinned = 0;
  ASSERT_EQ(fm_workbook_pinned_now(wb.handle, &now, &pinned), 0);
  EXPECT_EQ(pinned, 1);
  EXPECT_EQ(now.year, kPinned.year);
  EXPECT_EQ(now.month, kPinned.month);
  EXPECT_EQ(now.day, kPinned.day);
  EXPECT_EQ(now.hour, kPinned.hour);
  EXPECT_EQ(now.minute, kPinned.minute);
  EXPECT_EQ(now.second, kPinned.second);
}

TEST(FormulonCApiClock, ClearingRestoresTheHostClock) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &kPinned), 0);
  ASSERT_EQ(fm_workbook_clear_pinned_now(wb.handle), 0);

  fm_civil_time_t now{};
  int32_t pinned = 1;
  ASSERT_EQ(fm_workbook_pinned_now(wb.handle, &now, &pinned), 0);
  EXPECT_EQ(pinned, 0);
}

TEST(FormulonCApiClock, ClearingAnUnpinnedWorkbookIsANoOp) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_clear_pinned_now(wb.handle), 0);
  EXPECT_EQ(fm_workbook_clear_pinned_now(wb.handle), 0);
}

TEST(FormulonCApiClock, TodayAndNowResolveThroughThePin) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &kPinned), 0);

  fm_value_t today{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=TODAY()", &today), 0);
  ASSERT_EQ(today.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(today.u.number, kPinnedSerial);

  // 15:30:45 is 55,845 seconds into the day.
  fm_value_t now{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=NOW()", &now), 0);
  ASSERT_EQ(now.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(now.u.number, kPinnedSerial + 55845.0 / 86400.0);
}

TEST(FormulonCApiClock, MovingThePinMovesTheResult) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &kPinned), 0);
  const fm_civil_time_t next{2026, 4, 24, 0, 0, 0};
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &next), 0);

  fm_value_t today{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=TODAY()", &today), 0);
  ASSERT_EQ(today.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(today.u.number, kPinnedSerial + 1.0);
}

TEST(FormulonCApiClock, APinnedFormulaCellRecalcsAgainstThePin) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_pinned_now(wb.handle, &kPinned), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=TODAY()"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t cell{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &cell), 0);
  ASSERT_EQ(cell.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(cell.u.number, kPinnedSerial);
}

TEST(FormulonCApiClock, ASaveDoesNotCarryThePin) {
  WorkbookGuard source;
  ASSERT_EQ(fm_workbook_create(&source.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(source.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_pinned_now(source.handle, &kPinned), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(source.handle, &saved.data, &saved.len), 0);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  fm_civil_time_t now{};
  int32_t pinned = 1;
  ASSERT_EQ(fm_workbook_pinned_now(reloaded.handle, &now, &pinned), 0);
  EXPECT_EQ(pinned, 0);
}

TEST(FormulonCApiClock, EveryFieldIsRangeChecked) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const int32_t invalid_low[] = {1899, 0, 0, -1, -1, -1};
  const int32_t invalid_high[] = {10000, 13, 32, 24, 60, 60};
  for (int index = 0; index < 6; ++index) {
    const fm_civil_time_t low = WithField(kPinned, index, invalid_low[index]);
    const fm_civil_time_t high = WithField(kPinned, index, invalid_high[index]);
    EXPECT_EQ(fm_workbook_set_pinned_now(wb.handle, &low),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << "field index " << index;
    EXPECT_EQ(fm_workbook_set_pinned_now(wb.handle, &high),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << "field index " << index;
  }
  // None of the rejections left a pin behind.
  fm_civil_time_t now{};
  int32_t pinned = 1;
  ASSERT_EQ(fm_workbook_pinned_now(wb.handle, &now, &pinned), 0);
  EXPECT_EQ(pinned, 0);
}

TEST(FormulonCApiClock, TheDayBoundIsTheRealMonthLength) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const fm_civil_time_t leap_day{2024, 2, 29, 0, 0, 0};
  EXPECT_EQ(fm_workbook_set_pinned_now(wb.handle, &leap_day), 0);

  const fm_civil_time_t non_leap_day{2025, 2, 29, 0, 0, 0};
  EXPECT_EQ(fm_workbook_set_pinned_now(wb.handle, &non_leap_day),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // 1900-02-29 is Excel's ghost day: it exists on the serial scale but not on
  // the calendar, and a wall-clock reading is a calendar instant.
  const fm_civil_time_t ghost_day{1900, 2, 29, 0, 0, 0};
  EXPECT_EQ(fm_workbook_set_pinned_now(wb.handle, &ghost_day),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiClock, NullArgumentsAreRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_civil_time_t now{};
  int32_t pinned = 0;
  const auto null_ptr = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  EXPECT_EQ(fm_workbook_pinned_now(nullptr, &now, &pinned), null_ptr);
  EXPECT_EQ(fm_workbook_pinned_now(wb.handle, nullptr, &pinned), null_ptr);
  EXPECT_EQ(fm_workbook_pinned_now(wb.handle, &now, nullptr), null_ptr);
  EXPECT_EQ(fm_workbook_set_pinned_now(nullptr, &kPinned), null_ptr);
  EXPECT_EQ(fm_workbook_set_pinned_now(wb.handle, nullptr), null_ptr);
  EXPECT_EQ(fm_workbook_clear_pinned_now(nullptr), null_ptr);
}
