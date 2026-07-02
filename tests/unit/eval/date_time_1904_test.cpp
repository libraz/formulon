// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the 1904 date-system epoch handling added to the serial
// <-> calendar conversion helpers. The 1904 system anchors serial 0 at
// 1904-01-01 and carries no 1900 leap-year bug, so a 1904 serial is
// exactly 1462 less than the 1900 serial for the same calendar day, and
// the fictitious 1900-02-29 ghost day never appears.

#include "eval/date_time.h"
#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace date_time {
namespace {

TEST(DateTime1904, SerialDiffersBy1462FromDefault) {
  // DATE(2020, 1, 1): 43831 in the 1900 system, 42369 in the 1904 system.
  const double s1900 = serial_from_ymd(2020, 1, 1, /*date1904=*/false);
  const double s1904 = serial_from_ymd(2020, 1, 1, /*date1904=*/true);
  EXPECT_DOUBLE_EQ(s1900, 43831.0);
  EXPECT_DOUBLE_EQ(s1904, 42369.0);
  EXPECT_DOUBLE_EQ(s1900 - s1904, 1462.0);
}

TEST(DateTime1904, EpochAnchorIsJan1904) {
  // 1904-01-01 is serial 0 in the 1904 system and serial 1462 in 1900.
  EXPECT_DOUBLE_EQ(serial_from_ymd(1904, 1, 1, /*date1904=*/true), 0.0);
  EXPECT_DOUBLE_EQ(serial_from_ymd(1904, 1, 1, /*date1904=*/false), 1462.0);
}

TEST(DateTime1904, RoundTripsThroughYmd) {
  for (int serial : {0, 1, 100, 1000, 42369, 50000}) {
    const YMD ymd = ymd_from_serial(static_cast<double>(serial), /*date1904=*/true);
    const double back = serial_from_ymd(ymd.y, ymd.m, ymd.d, /*date1904=*/true);
    EXPECT_DOUBLE_EQ(back, static_cast<double>(serial)) << "serial=" << serial;
  }
}

TEST(DateTime1904, NoGhostLeapDay) {
  // The 1900 system reserves serial 60 for the fictitious 1900-02-29 and
  // has a one-day discontinuity across it. The 1904 system has neither:
  // conversions are linear across the 1900-02/03 boundary region.
  //
  // In the 1904 system the serials around the 1900 ghost region map to
  // ordinary 1904-era dates with no special-casing; verify a contiguous
  // run stays strictly monotonic and gap-free.
  YMD prev = ymd_from_serial(58.0, /*date1904=*/true);
  for (int serial = 59; serial <= 62; ++serial) {
    const YMD cur = ymd_from_serial(static_cast<double>(serial), /*date1904=*/true);
    const long long prev_days = days_from_civil(prev.y, prev.m, prev.d);
    const long long cur_days = days_from_civil(cur.y, cur.m, cur.d);
    EXPECT_EQ(cur_days - prev_days, 1) << "serial=" << serial << " must advance exactly one civil day";
    // None of these may land on the fictitious 1900-02-29.
    EXPECT_FALSE(cur.y == 1900 && cur.m == 2u && cur.d == 29u);
    prev = cur;
  }
}

TEST(DateTime1904, DefaultSystemUnaffected) {
  // The default (1900) behaviour, including the ghost day, must be
  // unchanged by the added parameter.
  EXPECT_DOUBLE_EQ(serial_from_ymd(1900, 2, 29), 60.0);
  const YMD ghost = ymd_from_serial(60.0);
  EXPECT_EQ(ghost.y, 1900);
  EXPECT_EQ(ghost.m, 2u);
  EXPECT_EQ(ghost.d, 29u);
  EXPECT_DOUBLE_EQ(serial_from_ymd(2020, 1, 1), 43831.0);
}

}  // namespace
}  // namespace date_time
}  // namespace eval
}  // namespace formulon
