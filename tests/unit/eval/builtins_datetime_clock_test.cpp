// End-to-end date/time built-in tests: clock seams and serial weekday grids.

#include <cmath>

#include "builtins_datetime_test_helpers.h"
#include "utils/date_time.h"
namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Clock seam
//
// `EvalContext::wall_clock()` falls through to the host clock unless a
// reading is pinned. Pinning is what makes NOW / TODAY assertable at all:
// without it every expectation here would have to be a shape check against
// whatever day the suite happens to run on, as the test above has to be.
// ---------------------------------------------------------------------------

// 2026-04-23 15:30:45. The date half matches `DateTimeDate.CurrentDateSerial`
// above, so the expected serial 46135 is already independently pinned.
constexpr date_time::CivilTime kPinnedNow{{2026, 4U, 23U}, {15U, 30U, 45U}};
constexpr double kPinnedSerial = 46135.0;
// 15:30:45 == 55,845 seconds into the day.
constexpr double kPinnedTimeOfDay = 55845.0 / 86400.0;

Value EvalSourcePinned(std::string_view src, date_time::CivilTime pinned = kPinnedNow, bool date1904 = false) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena, default_registry(),
                  test::mac_context().with_date1904(date1904).with_pinned_now(pinned));
}

TEST(ClockSeam, TodayReturnsThePinnedDate) {
  const Value v = EvalSourcePinned("=TODAY()");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), kPinnedSerial);
}

TEST(ClockSeam, NowCarriesThePinnedTimeOfDay) {
  const Value v = EvalSourcePinned("=NOW()");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), kPinnedSerial + kPinnedTimeOfDay);
}

TEST(ClockSeam, TodayDiscardsTheTimeOfDayNowKeeps) {
  const Value today = EvalSourcePinned("=TODAY()");
  const Value now = EvalSourcePinned("=NOW()");
  ASSERT_TRUE(today.is_number());
  ASSERT_TRUE(now.is_number());
  EXPECT_DOUBLE_EQ(today.as_number(), std::floor(now.as_number()));
  EXPECT_GT(now.as_number(), today.as_number());
}

TEST(ClockSeam, MovingThePinMovesTheAnswer) {
  // Guards against the pin being accepted but ignored: a second reading a
  // day later must shift the serial by exactly one.
  constexpr date_time::CivilTime kNextDay{{2026, 4U, 24U}, {15U, 30U, 45U}};
  const Value v = EvalSourcePinned("=TODAY()", kNextDay);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), kPinnedSerial + 1.0);
}

TEST(ClockSeam, ThePinnedReadingStillHonoursTheWorkbookEpoch) {
  // The pin carries calendar fields, not a serial, so the 1904 epoch is
  // applied downstream exactly as it is for an argument-supplied date.
  const Value v = EvalSourcePinned("=TODAY()", kPinnedNow, /*date1904=*/true);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), kPinnedSerial - date_time::kDate1904EpochGap);
}

TEST(ClockSeam, AnUnpinnedContextStillFollowsTheHostClock) {
  // The seam must not disturb the shipping default: with nothing pinned both
  // functions keep reading the host clock, so only the shape is assertable.
  const Value v = EvalSource("=TODAY()");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_DOUBLE_EQ(v.as_number(), std::floor(v.as_number()));
}
// ---------------------------------------------------------------------------
// Weekdays of 1900 serials
// ---------------------------------------------------------------------------

// Excel 365 ja-JP readings for serials 0..8, around the ghost day 60, and
// across 1900 for the week-number family.
TEST(DateTimeSerialWeekday, MatchesExcelGrid) {
  struct Case {
    const char* formula;
    double number;
    const char* text;
    bool num_error;
  };
  static constexpr Case kCases[] = {
      {"=WEEKDAY(0)", 7.0, nullptr, false},
      {"=WEEKDAY(0,2)", 6.0, nullptr, false},
      {"=WEEKDAY(0,3)", 5.0, nullptr, false},
      {"=WEEKDAY(1)", 1.0, nullptr, false},
      {"=WEEKDAY(1,2)", 7.0, nullptr, false},
      {"=WEEKDAY(1,3)", 6.0, nullptr, false},
      {"=WEEKDAY(2)", 2.0, nullptr, false},
      {"=WEEKDAY(2,2)", 1.0, nullptr, false},
      {"=WEEKDAY(2,3)", 0.0, nullptr, false},
      {"=WEEKDAY(3)", 3.0, nullptr, false},
      {"=WEEKDAY(3,2)", 2.0, nullptr, false},
      {"=WEEKDAY(3,3)", 1.0, nullptr, false},
      {"=WEEKDAY(4)", 4.0, nullptr, false},
      {"=WEEKDAY(4,2)", 3.0, nullptr, false},
      {"=WEEKDAY(4,3)", 2.0, nullptr, false},
      {"=WEEKDAY(5)", 5.0, nullptr, false},
      {"=WEEKDAY(5,2)", 4.0, nullptr, false},
      {"=WEEKDAY(5,3)", 3.0, nullptr, false},
      {"=WEEKDAY(6)", 6.0, nullptr, false},
      {"=WEEKDAY(6,2)", 5.0, nullptr, false},
      {"=WEEKDAY(6,3)", 4.0, nullptr, false},
      {"=WEEKDAY(7)", 7.0, nullptr, false},
      {"=WEEKDAY(7,2)", 6.0, nullptr, false},
      {"=WEEKDAY(7,3)", 5.0, nullptr, false},
      {"=WEEKDAY(8)", 1.0, nullptr, false},
      {"=WEEKDAY(8,2)", 7.0, nullptr, false},
      {"=WEEKDAY(8,3)", 6.0, nullptr, false},
      {"=WEEKDAY(58)", 2.0, nullptr, false},
      {"=WEEKDAY(58,2)", 1.0, nullptr, false},
      {"=WEEKDAY(58,3)", 0.0, nullptr, false},
      {"=WEEKDAY(59)", 3.0, nullptr, false},
      {"=WEEKDAY(59,2)", 2.0, nullptr, false},
      {"=WEEKDAY(59,3)", 1.0, nullptr, false},
      {"=WEEKDAY(60)", 4.0, nullptr, false},
      {"=WEEKDAY(60,2)", 3.0, nullptr, false},
      {"=WEEKDAY(60,3)", 2.0, nullptr, false},
      {"=WEEKDAY(61)", 5.0, nullptr, false},
      {"=WEEKDAY(61,2)", 4.0, nullptr, false},
      {"=WEEKDAY(61,3)", 3.0, nullptr, false},
      {"=WEEKDAY(62)", 6.0, nullptr, false},
      {"=WEEKDAY(62,2)", 5.0, nullptr, false},
      {"=WEEKDAY(62,3)", 4.0, nullptr, false},
      {"=NETWORKDAYS(0,0)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,0)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,0,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,0,\"0000011\")", 0.0, nullptr, false},
      {"=NETWORKDAYS(0,1)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,1)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,1,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,1,\"0000011\")", 0.0, nullptr, false},
      {"=NETWORKDAYS(0,2)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,2)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,2,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,2,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(0,3)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,3)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,3,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,3,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(0,4)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,4)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,4,11)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,4,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(0,5)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,5)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,5,11)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,5,\"0000011\")", 4.0, nullptr, false},
      {"=NETWORKDAYS(0,6)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,6)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,6,11)", 6.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,6,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(0,7)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,7)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,7,11)", 7.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,7,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(0,8)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,8)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,8,11)", 7.0, nullptr, false},
      {"=NETWORKDAYS.INTL(0,8,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(1,1)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,1)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,1,11)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,1,\"0000011\")", 0.0, nullptr, false},
      {"=NETWORKDAYS(1,2)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,2)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,2,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,2,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(1,3)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,3)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,3,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,3,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(1,4)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,4)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,4,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,4,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(1,5)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,5)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,5,11)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,5,\"0000011\")", 4.0, nullptr, false},
      {"=NETWORKDAYS(1,6)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,6)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,6,11)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,6,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(1,7)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,7)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,7,11)", 6.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,7,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(1,8)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,8)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,8,11)", 6.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,8,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(2,2)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,2)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,2,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,2,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(2,3)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,3)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,3,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,3,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(2,4)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,4)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,4,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,4,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(2,5)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,5)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,5,11)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,5,\"0000011\")", 4.0, nullptr, false},
      {"=NETWORKDAYS(2,6)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,6)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,6,11)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,6,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(2,7)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,7)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,7,11)", 6.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,7,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(2,8)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,8)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,8,11)", 6.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,8,\"0000011\")", 5.0, nullptr, false},
      {"=NETWORKDAYS(3,3)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,3)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,3,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,3,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(3,4)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,4)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,4,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,4,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(3,5)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,5)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,5,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,5,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(3,6)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,6)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,6,11)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,6,\"0000011\")", 4.0, nullptr, false},
      {"=NETWORKDAYS(3,7)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,7)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,7,11)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,7,\"0000011\")", 4.0, nullptr, false},
      {"=NETWORKDAYS(3,8)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,8)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,8,11)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(3,8,\"0000011\")", 4.0, nullptr, false},
      {"=NETWORKDAYS(4,4)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,4)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,4,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,4,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(4,5)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,5)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,5,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,5,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(4,6)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,6)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,6,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,6,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(4,7)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,7)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,7,11)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,7,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(4,8)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,8)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,8,11)", 4.0, nullptr, false},
      {"=NETWORKDAYS.INTL(4,8,\"0000011\")", 3.0, nullptr, false},
      {"=NETWORKDAYS(5,5)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,5)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,5,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,5,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(5,6)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,6)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,6,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,6,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(5,7)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,7)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,7,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,7,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(5,8)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,8)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,8,11)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(5,8,\"0000011\")", 2.0, nullptr, false},
      {"=NETWORKDAYS(6,6)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,6)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,6,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,6,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(6,7)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,7)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,7,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,7,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(6,8)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,8)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,8,11)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(6,8,\"0000011\")", 1.0, nullptr, false},
      {"=NETWORKDAYS(7,7)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(7,7)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(7,7,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(7,7,\"0000011\")", 0.0, nullptr, false},
      {"=NETWORKDAYS(7,8)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(7,8)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(7,8,11)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(7,8,\"0000011\")", 0.0, nullptr, false},
      {"=NETWORKDAYS(8,8)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(8,8)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(8,8,11)", 0.0, nullptr, false},
      {"=NETWORKDAYS.INTL(8,8,\"0000011\")", 0.0, nullptr, false},
      {"=NETWORKDAYS(55,65)", 7.0, nullptr, false},
      {"=NETWORKDAYS.INTL(55,65)", 7.0, nullptr, false},
      {"=NETWORKDAYS(58,62)", 5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(58,62)", 5.0, nullptr, false},
      {"=NETWORKDAYS(59,61)", 3.0, nullptr, false},
      {"=NETWORKDAYS.INTL(59,61)", 3.0, nullptr, false},
      {"=NETWORKDAYS(60,60)", 1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(60,60)", 1.0, nullptr, false},
      {"=NETWORKDAYS(59,60)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(59,60)", 2.0, nullptr, false},
      {"=NETWORKDAYS(60,61)", 2.0, nullptr, false},
      {"=NETWORKDAYS.INTL(60,61)", 2.0, nullptr, false},
      {"=NETWORKDAYS(1,70)", 50.0, nullptr, false},
      {"=NETWORKDAYS.INTL(1,70)", 50.0, nullptr, false},
      {"=NETWORKDAYS(2,1)", -1.0, nullptr, false},
      {"=NETWORKDAYS.INTL(2,1)", -1.0, nullptr, false},
      {"=NETWORKDAYS(8,0)", -5.0, nullptr, false},
      {"=NETWORKDAYS.INTL(8,0)", -5.0, nullptr, false},
      {"=WORKDAY(0,1)", 2.0, nullptr, false},
      {"=WORKDAY.INTL(0,1)", 2.0, nullptr, false},
      {"=WORKDAY.INTL(0,1,11)", 2.0, nullptr, false},
      {"=WORKDAY(1,1)", 2.0, nullptr, false},
      {"=WORKDAY.INTL(1,1)", 2.0, nullptr, false},
      {"=WORKDAY.INTL(1,1,11)", 2.0, nullptr, false},
      {"=WORKDAY(0,0)", 0.0, nullptr, false},
      {"=WORKDAY.INTL(0,0)", 0.0, nullptr, false},
      {"=WORKDAY.INTL(0,0,11)", 0.0, nullptr, false},
      {"=WORKDAY(1,0)", 1.0, nullptr, false},
      {"=WORKDAY.INTL(1,0)", 1.0, nullptr, false},
      {"=WORKDAY.INTL(1,0,11)", 1.0, nullptr, false},
      {"=WORKDAY(0,5)", 6.0, nullptr, false},
      {"=WORKDAY.INTL(0,5)", 6.0, nullptr, false},
      {"=WORKDAY.INTL(0,5,11)", 6.0, nullptr, false},
      {"=WORKDAY(1,5)", 6.0, nullptr, false},
      {"=WORKDAY.INTL(1,5)", 6.0, nullptr, false},
      {"=WORKDAY.INTL(1,5,11)", 6.0, nullptr, false},
      {"=WORKDAY(2,-1)", 0.0, nullptr, true},
      {"=WORKDAY.INTL(2,-1)", 0.0, nullptr, true},
      {"=WORKDAY.INTL(2,-1,11)", 0.0, nullptr, false},
      {"=WORKDAY(1,-1)", 0.0, nullptr, true},
      {"=WORKDAY.INTL(1,-1)", 0.0, nullptr, true},
      {"=WORKDAY.INTL(1,-1,11)", 0.0, nullptr, false},
      {"=WORKDAY(57,3)", 60.0, nullptr, false},
      {"=WORKDAY.INTL(57,3)", 60.0, nullptr, false},
      {"=WORKDAY.INTL(57,3,11)", 60.0, nullptr, false},
      {"=WORKDAY(58,1)", 59.0, nullptr, false},
      {"=WORKDAY.INTL(58,1)", 59.0, nullptr, false},
      {"=WORKDAY.INTL(58,1,11)", 59.0, nullptr, false},
      {"=WORKDAY(59,1)", 60.0, nullptr, false},
      {"=WORKDAY.INTL(59,1)", 60.0, nullptr, false},
      {"=WORKDAY.INTL(59,1,11)", 60.0, nullptr, false},
      {"=WORKDAY(61,-1)", 60.0, nullptr, false},
      {"=WORKDAY.INTL(61,-1)", 60.0, nullptr, false},
      {"=WORKDAY.INTL(61,-1,11)", 60.0, nullptr, false},
      {"=WORKDAY(62,-3)", 59.0, nullptr, false},
      {"=WORKDAY.INTL(62,-3)", 59.0, nullptr, false},
      {"=WORKDAY.INTL(62,-3,11)", 59.0, nullptr, false},
      {"=TEXT(0,\"ddd\")", 0.0, "Sat", false},
      {"=TEXT(0,\"dddd\")", 0.0, "Saturday", false},
      {"=TEXT(0,\"aaa\")", 0.0, "土", false},
      {"=WEEKNUM(0)", 0.0, nullptr, false},
      {"=WEEKNUM(0,2)", 1.0, nullptr, false},
      {"=WEEKNUM(0,21)", 52.0, nullptr, false},
      {"=ISOWEEKNUM(0)", 52.0, nullptr, false},
      {"=WEEKNUM(0,11)", 1.0, nullptr, false},
      {"=WEEKNUM(0,17)", 0.0, nullptr, false},
      {"=TEXT(1,\"ddd\")", 0.0, "Sun", false},
      {"=TEXT(1,\"dddd\")", 0.0, "Sunday", false},
      {"=TEXT(1,\"aaa\")", 0.0, "日", false},
      {"=WEEKNUM(1)", 1.0, nullptr, false},
      {"=WEEKNUM(1,2)", 1.0, nullptr, false},
      {"=WEEKNUM(1,21)", 52.0, nullptr, false},
      {"=ISOWEEKNUM(1)", 52.0, nullptr, false},
      {"=WEEKNUM(1,11)", 1.0, nullptr, false},
      {"=WEEKNUM(1,17)", 1.0, nullptr, false},
      {"=TEXT(2,\"ddd\")", 0.0, "Mon", false},
      {"=TEXT(2,\"dddd\")", 0.0, "Monday", false},
      {"=TEXT(2,\"aaa\")", 0.0, "月", false},
      {"=WEEKNUM(2)", 1.0, nullptr, false},
      {"=WEEKNUM(2,2)", 2.0, nullptr, false},
      {"=WEEKNUM(2,21)", 1.0, nullptr, false},
      {"=ISOWEEKNUM(2)", 1.0, nullptr, false},
      {"=WEEKNUM(2,11)", 2.0, nullptr, false},
      {"=WEEKNUM(2,17)", 1.0, nullptr, false},
      {"=TEXT(6,\"ddd\")", 0.0, "Fri", false},
      {"=TEXT(6,\"dddd\")", 0.0, "Friday", false},
      {"=TEXT(6,\"aaa\")", 0.0, "金", false},
      {"=WEEKNUM(6)", 1.0, nullptr, false},
      {"=WEEKNUM(6,2)", 2.0, nullptr, false},
      {"=WEEKNUM(6,21)", 1.0, nullptr, false},
      {"=ISOWEEKNUM(6)", 1.0, nullptr, false},
      {"=WEEKNUM(6,11)", 2.0, nullptr, false},
      {"=WEEKNUM(6,17)", 1.0, nullptr, false},
      {"=TEXT(7,\"ddd\")", 0.0, "Sat", false},
      {"=TEXT(7,\"dddd\")", 0.0, "Saturday", false},
      {"=TEXT(7,\"aaa\")", 0.0, "土", false},
      {"=WEEKNUM(7)", 1.0, nullptr, false},
      {"=WEEKNUM(7,2)", 2.0, nullptr, false},
      {"=WEEKNUM(7,21)", 1.0, nullptr, false},
      {"=ISOWEEKNUM(7)", 1.0, nullptr, false},
      {"=WEEKNUM(7,11)", 2.0, nullptr, false},
      {"=WEEKNUM(7,17)", 1.0, nullptr, false},
      {"=TEXT(8,\"ddd\")", 0.0, "Sun", false},
      {"=TEXT(8,\"dddd\")", 0.0, "Sunday", false},
      {"=TEXT(8,\"aaa\")", 0.0, "日", false},
      {"=WEEKNUM(8)", 2.0, nullptr, false},
      {"=WEEKNUM(8,2)", 2.0, nullptr, false},
      {"=WEEKNUM(8,21)", 1.0, nullptr, false},
      {"=ISOWEEKNUM(8)", 1.0, nullptr, false},
      {"=WEEKNUM(8,11)", 2.0, nullptr, false},
      {"=WEEKNUM(8,17)", 2.0, nullptr, false},
      {"=TEXT(14,\"ddd\")", 0.0, "Sat", false},
      {"=TEXT(14,\"dddd\")", 0.0, "Saturday", false},
      {"=TEXT(14,\"aaa\")", 0.0, "土", false},
      {"=WEEKNUM(14)", 2.0, nullptr, false},
      {"=WEEKNUM(14,2)", 3.0, nullptr, false},
      {"=WEEKNUM(14,21)", 2.0, nullptr, false},
      {"=ISOWEEKNUM(14)", 2.0, nullptr, false},
      {"=WEEKNUM(14,11)", 3.0, nullptr, false},
      {"=WEEKNUM(14,17)", 2.0, nullptr, false},
      {"=TEXT(15,\"ddd\")", 0.0, "Sun", false},
      {"=TEXT(15,\"dddd\")", 0.0, "Sunday", false},
      {"=TEXT(15,\"aaa\")", 0.0, "日", false},
      {"=WEEKNUM(15)", 3.0, nullptr, false},
      {"=WEEKNUM(15,2)", 3.0, nullptr, false},
      {"=WEEKNUM(15,21)", 2.0, nullptr, false},
      {"=ISOWEEKNUM(15)", 2.0, nullptr, false},
      {"=WEEKNUM(15,11)", 3.0, nullptr, false},
      {"=WEEKNUM(15,17)", 3.0, nullptr, false},
      {"=TEXT(59,\"ddd\")", 0.0, "Tue", false},
      {"=TEXT(59,\"dddd\")", 0.0, "Tuesday", false},
      {"=TEXT(59,\"aaa\")", 0.0, "火", false},
      {"=WEEKNUM(59)", 9.0, nullptr, false},
      {"=WEEKNUM(59,2)", 10.0, nullptr, false},
      {"=WEEKNUM(59,21)", 9.0, nullptr, false},
      {"=ISOWEEKNUM(59)", 9.0, nullptr, false},
      {"=WEEKNUM(59,11)", 10.0, nullptr, false},
      {"=WEEKNUM(59,17)", 9.0, nullptr, false},
      {"=TEXT(60,\"ddd\")", 0.0, "Wed", false},
      {"=TEXT(60,\"dddd\")", 0.0, "Wednesday", false},
      {"=TEXT(60,\"aaa\")", 0.0, "水", false},
      {"=WEEKNUM(60)", 9.0, nullptr, false},
      {"=WEEKNUM(60,2)", 10.0, nullptr, false},
      {"=WEEKNUM(60,21)", 9.0, nullptr, false},
      {"=ISOWEEKNUM(60)", 9.0, nullptr, false},
      {"=WEEKNUM(60,11)", 10.0, nullptr, false},
      {"=WEEKNUM(60,17)", 9.0, nullptr, false},
      {"=TEXT(61,\"ddd\")", 0.0, "Thu", false},
      {"=TEXT(61,\"dddd\")", 0.0, "Thursday", false},
      {"=TEXT(61,\"aaa\")", 0.0, "木", false},
      {"=WEEKNUM(61)", 9.0, nullptr, false},
      {"=WEEKNUM(61,2)", 10.0, nullptr, false},
      {"=WEEKNUM(61,21)", 9.0, nullptr, false},
      {"=ISOWEEKNUM(61)", 9.0, nullptr, false},
      {"=WEEKNUM(61,11)", 10.0, nullptr, false},
      {"=WEEKNUM(61,17)", 9.0, nullptr, false},
      {"=TEXT(62,\"ddd\")", 0.0, "Fri", false},
      {"=TEXT(62,\"dddd\")", 0.0, "Friday", false},
      {"=TEXT(62,\"aaa\")", 0.0, "金", false},
      {"=WEEKNUM(62)", 9.0, nullptr, false},
      {"=WEEKNUM(62,2)", 10.0, nullptr, false},
      {"=WEEKNUM(62,21)", 9.0, nullptr, false},
      {"=ISOWEEKNUM(62)", 9.0, nullptr, false},
      {"=WEEKNUM(62,11)", 10.0, nullptr, false},
      {"=WEEKNUM(62,17)", 9.0, nullptr, false},
      {"=TEXT(365,\"ddd\")", 0.0, "Sun", false},
      {"=TEXT(365,\"dddd\")", 0.0, "Sunday", false},
      {"=TEXT(365,\"aaa\")", 0.0, "日", false},
      {"=WEEKNUM(365)", 53.0, nullptr, false},
      {"=WEEKNUM(365,2)", 53.0, nullptr, false},
      {"=WEEKNUM(365,21)", 52.0, nullptr, false},
      {"=ISOWEEKNUM(365)", 52.0, nullptr, false},
      {"=WEEKNUM(365,11)", 53.0, nullptr, false},
      {"=WEEKNUM(365,17)", 53.0, nullptr, false},
      {"=TEXT(366,\"ddd\")", 0.0, "Mon", false},
      {"=TEXT(366,\"dddd\")", 0.0, "Monday", false},
      {"=TEXT(366,\"aaa\")", 0.0, "月", false},
      {"=WEEKNUM(366)", 53.0, nullptr, false},
      {"=WEEKNUM(366,2)", 54.0, nullptr, false},
      {"=WEEKNUM(366,21)", 1.0, nullptr, false},
      {"=ISOWEEKNUM(366)", 1.0, nullptr, false},
      {"=WEEKNUM(366,11)", 54.0, nullptr, false},
      {"=WEEKNUM(366,17)", 53.0, nullptr, false},
  };
  for (const Case& c : kCases) {
    const Value v = EvalSource(c.formula);
    if (c.num_error) {
      ASSERT_TRUE(v.is_error()) << c.formula;
      EXPECT_EQ(v.as_error(), ErrorCode::Num) << c.formula;
    } else if (c.text != nullptr) {
      ASSERT_TRUE(v.is_text()) << c.formula;
      EXPECT_EQ(v.as_text(), c.text) << c.formula;
    } else {
      ASSERT_TRUE(v.is_number()) << c.formula;
      EXPECT_EQ(v.as_number(), c.number) << c.formula;
    }
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
