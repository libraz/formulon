//
// Unit tests for the strict ISO 8601 date-time parser.

#include "io/iso_date.h"

#include "gtest/gtest.h"

namespace formulon::io {
namespace {

TEST(IsoDate, PlainDateParses) {
  double serial = 0.0;
  ASSERT_TRUE(parse_iso_date_serial("2024-02-29", &serial));
  EXPECT_GT(serial, 0.0);
}

TEST(IsoDate, DateTimeWithFractionalSecondsParses) {
  double serial = 0.0;
  ASSERT_TRUE(parse_iso_date_serial("2024-01-01T12:30:45.5", &serial));
  EXPECT_GT(serial, 0.0);
}

TEST(IsoDate, LeapDayInLeapYearIsAccepted) {
  // 2024 is a leap year (divisible by 4, not by 100).
  double serial = 0.0;
  EXPECT_TRUE(parse_iso_date_serial("2024-02-29", &serial));
}

TEST(IsoDate, LeapDayInNonLeapYearIsRejected) {
  // 2023 is not a leap year: February 29 does not exist.
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("2023-02-29", &serial));
}

TEST(IsoDate, LeapDayInCenturyNonLeapYearIsRejected) {
  // 1900 is divisible by 100 but not 400: not a leap year under the
  // Gregorian rule.
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("1900-02-29", &serial));
}

TEST(IsoDate, LeapDayInCenturyLeapYearIsAccepted) {
  // 2000 is divisible by 400: a leap year.
  double serial = 0.0;
  EXPECT_TRUE(parse_iso_date_serial("2000-02-29", &serial));
}

TEST(IsoDate, Day30OfFebruaryIsRejected) {
  // Regression: `serial_from_ymd` silently normalizes an out-of-range day
  // into the following month, so without an explicit month-length check
  // `2024-02-30` used to parse successfully as March 1.
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("2024-02-30", &serial));
}

TEST(IsoDate, Day31OfApril30DayMonthIsRejected) {
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("2024-04-31", &serial));
}

TEST(IsoDate, Day31OfMonthWith31DaysIsAccepted) {
  double serial = 0.0;
  EXPECT_TRUE(parse_iso_date_serial("2024-01-31", &serial));
}

TEST(IsoDate, MonthOutOfRangeIsRejected) {
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("2024-13-01", &serial));
}

TEST(IsoDate, DayZeroIsRejected) {
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("2024-01-00", &serial));
}

TEST(IsoDate, MalformedLexemeIsRejected) {
  double serial = 0.0;
  EXPECT_FALSE(parse_iso_date_serial("not-a-date", &serial));
}

TEST(IsoDate, TimeZoneOffsetIsAcceptedButIgnored) {
  double serial = 0.0;
  ASSERT_TRUE(parse_iso_date_serial("2024-01-01T00:00:00+09:00", &serial));
}

TEST(IsoDate, TrailingZDesignatorIsAccepted) {
  double serial = 0.0;
  EXPECT_TRUE(parse_iso_date_serial("2024-01-01T00:00:00Z", &serial));
}

}  // namespace
}  // namespace formulon::io
