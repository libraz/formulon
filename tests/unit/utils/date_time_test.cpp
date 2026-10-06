#include "utils/date_time.h"

#include <cstdint>
#include <limits>

#include "gtest/gtest.h"

namespace formulon::date_time {
namespace {

TEST(DateTimeNormalization, ZeroDayRewindsAcrossMarchInBothEpochs) {
  EXPECT_EQ(days_from_civil(2024, 3, 0), days_from_civil(2024, 2, 29));
  EXPECT_EQ(days_from_civil(2023, 3, 0), days_from_civil(2023, 2, 28));
  for (bool date1904 : {false, true}) {
    EXPECT_DOUBLE_EQ(serial_from_ymd(2024, 3, 0, date1904), serial_from_ymd(2024, 2, 29, date1904));
  }
}

TEST(DateTimeNormalization, MonthOverflowUsesCalendarMonthLengths) {
  EXPECT_EQ(days_from_civil(2023, 15, 1), days_from_civil(2024, 3, 1));
  EXPECT_EQ(days_from_civil(2024, 0, 1), days_from_civil(2023, 12, 1));
  EXPECT_EQ(days_from_civil(2024, 25, 1), days_from_civil(2026, 1, 1));
  for (bool date1904 : {false, true}) {
    EXPECT_DOUBLE_EQ(serial_from_ymd(2023, 15, 1, date1904), serial_from_ymd(2024, 3, 1, date1904));
  }
}

TEST(DateTimeNormalization, WideDayOffsetsDoNotWrap) {
  constexpr unsigned day = std::numeric_limits<unsigned>::max();
  EXPECT_EQ(days_from_civil(2024, 1, day), days_from_civil(2024, 1, 1) + static_cast<std::int64_t>(day) - 1);
}

TEST(DateTimeNormalization, ExtremeYearsRetainTheirLeapRule) {
  constexpr int first = std::numeric_limits<int>::min();
  constexpr int last = std::numeric_limits<int>::max();
  EXPECT_EQ(days_from_civil(first, 3, 1) - days_from_civil(first, 1, 1), 60);
  EXPECT_EQ(days_from_civil(last, 3, 1) - days_from_civil(last, 1, 1), 59);
}

TEST(DateTimeNormalization, ExtremeYearsRoundTripWithoutIntermediateOverflow) {
  for (int year : {std::numeric_limits<int>::min(), std::numeric_limits<int>::max()}) {
    const YMD date = civil_from_days(days_from_civil(year, 3, 1));
    EXPECT_EQ(date.y, year);
    EXPECT_EQ(date.m, 3u);
    EXPECT_EQ(date.d, 1u);
  }
  EXPECT_EQ(days_from_civil(0, std::numeric_limits<unsigned>::max(), 1), days_from_civil(357913941, 3, 1));
}

}  // namespace
}  // namespace formulon::date_time
