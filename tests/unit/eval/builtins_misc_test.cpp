// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for a miscellaneous batch of builtins that do not fit
// neatly into the per-category test files: XOR (logical), DAYS360,
// NETWORKDAYS.INTL, WORKDAY.INTL, NOW, TODAY (date/time), and AREAS
// (reference). Each test parses a formula, evaluates it through the
// default registry, and asserts the resulting Value. The structured cases
// for NETWORKDAYS.INTL / WORKDAY.INTL reuse `EvalSourceIn` against a
// bound workbook so the range-sourced holiday path is exercised alongside
// the inline-array form.

#include <chrono>
#include <cmath>
#include <ctime>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it via the default function registry. The
// thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
Value EvalSource(std::string_view src) {
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
  return evaluate(*root, eval_arena);
}

// Bound-workbook variant for tests that need A1-style references to resolve.
Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// Computes the Excel serial date+time for the current local wall clock.
// Used by the NOW / TODAY tests to derive an expected value from the same
// clock the impl reads. Returns a pair (date_serial, tod_fraction) so the
// test can assert the two halves independently.
struct LocalNowExpect {
  double date_serial;
  double tod_fraction;
};

LocalNowExpect ExpectedLocalNow() {
  using std::chrono::system_clock;
  const std::time_t tt = system_clock::to_time_t(system_clock::now());
  std::tm local{};
#if defined(_WIN32)
  localtime_s(&local, &tt);
#else
  localtime_r(&tt, &local);
#endif
  const int y = local.tm_year + 1900;
  const int m = local.tm_mon + 1;
  const int d = local.tm_mday;
  // Compute the Excel serial by delegating to the engine's own DATE impl
  // via a formula so any future epoch adjustment stays in one place.
  std::string src = "=DATE(" + std::to_string(y) + "," + std::to_string(m) + "," + std::to_string(d) + ")";
  const Value v = EvalSource(src);
  EXPECT_TRUE(v.is_number());
  const double date_serial = v.is_number() ? v.as_number() : 0.0;
  const double tod = (static_cast<double>(local.tm_hour) * 3600.0 + static_cast<double>(local.tm_min) * 60.0 +
                      static_cast<double>(local.tm_sec)) /
                     86400.0;
  return LocalNowExpect{date_serial, tod};
}

// ---------------------------------------------------------------------------
// Registry pin -- catches accidental drops / renames during refactors.
// ---------------------------------------------------------------------------

TEST(BuiltinsMiscRegistry, AllNamesRegistered) {
  // A mix of eager registry entries and lazy-dispatch entries: NOW /
  // TODAY / DAYS360 / XOR live in the FunctionRegistry; AREAS /
  // NETWORKDAYS.INTL / WORKDAY.INTL live in the lazy dispatch table.
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("XOR"), nullptr);
  EXPECT_NE(reg.lookup("DAYS360"), nullptr);
  EXPECT_NE(reg.lookup("NOW"), nullptr);
  EXPECT_NE(reg.lookup("TODAY"), nullptr);
  // Lazy forms: resolve via evaluation rather than a direct registry lookup.
  // The names must at least parse and produce a non-#NAME? result on valid
  // input; a plain `AREAS(A1)` should evaluate to 1.
  const Value areas = EvalSource("=AREAS(A1)");
  ASSERT_TRUE(areas.is_number());
  EXPECT_EQ(areas.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// XOR
// ---------------------------------------------------------------------------

TEST(BuiltinsXor, TrueFalseIsTrue) {
  const Value v = EvalSource("=XOR(TRUE,FALSE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsXor, TwoTrueIsFalse) {
  const Value v = EvalSource("=XOR(TRUE,TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsXor, TwoFalseIsFalse) {
  const Value v = EvalSource("=XOR(FALSE,FALSE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsXor, ThreeTrueIsTrue) {
  // Odd number of TRUE -> TRUE.
  const Value v = EvalSource("=XOR(TRUE,TRUE,TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsXor, FourTrueIsFalse) {
  const Value v = EvalSource("=XOR(TRUE,TRUE,TRUE,TRUE)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsXor, NumericCoerce) {
  // 1 -> TRUE, 0 -> FALSE; XOR(1,0) -> TRUE.
  const Value v = EvalSource("=XOR(1,0)");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
}

TEST(BuiltinsXor, SingleArgumentPassesThrough) {
  // XOR(TRUE) -> TRUE, XOR(FALSE) -> FALSE. The `min_arity = 1` guard
  // means zero-argument calls are rejected at dispatch time.
  const Value a = EvalSource("=XOR(TRUE)");
  ASSERT_TRUE(a.is_boolean());
  EXPECT_TRUE(a.as_boolean());
  const Value b = EvalSource("=XOR(FALSE)");
  ASSERT_TRUE(b.is_boolean());
  EXPECT_FALSE(b.as_boolean());
}

TEST(BuiltinsXor, InlineArrayOddCountTrue) {
  // {TRUE,FALSE,TRUE} has two TRUE values -> FALSE.
  const Value v = EvalSource("=XOR({TRUE,FALSE,TRUE})");
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
}

TEST(BuiltinsXor, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=XOR(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsXor, ErrorPropagates) {
  const Value v = EvalSource("=XOR(#REF!,TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsXor, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=XOR()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// DAYS360
// ---------------------------------------------------------------------------

TEST(BuiltinsDays360, FullYearIs360) {
  const Value v = EvalSource("=DAYS360(DATE(2024,1,1),DATE(2025,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 360.0);
}

TEST(BuiltinsDays360, Jan31ToFeb28UsMethod) {
  // US/NASD: start day 31 becomes 30; end day 28 stays. Result:
  // 360*0 + 30*(2-1) + (28-30) = 28. (No step 1 trigger because end < 31.)
  const Value v = EvalSource("=DAYS360(DATE(2024,1,31),DATE(2024,2,28))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 28.0);
}

TEST(BuiltinsDays360, Jan31ToFeb28EuropeanMethod) {
  // European: both days capped at 30, so start 31->30, end 28 unchanged.
  // 360*0 + 30*1 + (28-30) = 28. Same as US here because start's 31 is
  // the only difference, and both methods reduce it.
  const Value v = EvalSource("=DAYS360(DATE(2024,1,31),DATE(2024,2,28),TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 28.0);
}

TEST(BuiltinsDays360, Jan31ToMar31EuropeanMethod) {
  // European: both 31s become 30, so 2 months * 30 = 60.
  const Value v = EvalSource("=DAYS360(DATE(2024,1,31),DATE(2024,3,31),TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(BuiltinsDays360, Jan31ToMar31UsMethod) {
  // US: start 31 + end 31 -> start_is_last=true, step 2 fires (end day
  // becomes 30), then start becomes 30. Result: 360*0 + 30*(3-1) +
  // (30-30) = 60.
  const Value v = EvalSource("=DAYS360(DATE(2024,1,31),DATE(2024,3,31))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 60.0);
}

TEST(BuiltinsDays360, End31StartUnder30UsRolls) {
  // US: start day < 30, end day = 31 -> end becomes 1, end month +=1.
  // DATE(2024,2,15) -> DATE(2024,3,31): start=(2024,2,15), end after
  // rule 1 = (2024,4,1). Result: 360*0 + 30*(4-2) + (1-15) = 46.
  const Value v = EvalSource("=DAYS360(DATE(2024,2,15),DATE(2024,3,31))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 46.0);
}

TEST(BuiltinsDays360, End31StartUnder30EuropeanNoRoll) {
  // European: end 31 becomes 30, no month-roll. Result:
  // 360*0 + 30*(3-2) + (30-15) = 45.
  const Value v = EvalSource("=DAYS360(DATE(2024,2,15),DATE(2024,3,31),TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45.0);
}

TEST(BuiltinsDays360, ReversedInputsReturnNegative) {
  // DAYS360 accepts any order; (end < start) produces a negative count.
  const Value v = EvalSource("=DAYS360(DATE(2025,1,1),DATE(2024,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -360.0);
}

TEST(BuiltinsDays360, NegativeSerialIsNum) {
  // Negative coerced serial -> #NUM! (ymd_from_serial rejects pre-epoch).
  const Value v = EvalSource("=DAYS360(-1,DATE(2024,1,1))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsDays360, ErrorPropagates) {
  const Value v = EvalSource("=DAYS360(\"abc\",DATE(2024,1,1))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// NETWORKDAYS.INTL
// ---------------------------------------------------------------------------

TEST(BuiltinsNetworkdaysIntl, DefaultMatchesNetworkdays) {
  // Default weekend = 1 (Sat+Sun). Should match NETWORKDAYS exactly.
  const Value a = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,12))");
  const Value b = EvalSource("=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,12))");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
  EXPECT_EQ(a.as_number(), 10.0);
}

TEST(BuiltinsNetworkdaysIntl, ExplicitWeekend1) {
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,14),1)");
  // Mon Jan 1 .. Sun Jan 14 with Sat+Sun as weekend -> 10 weekdays.
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsNetworkdaysIntl, StringMaskMatchesWeekend1) {
  // "0000011" = Sat+Sun weekend (positions 5,6). Should match default.
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,14),\"0000011\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsNetworkdaysIntl, SundayOnlyWeekend) {
  // Selector 11 = Sunday only. Mon Jan 1 .. Sun Jan 7 = 7 days, minus 1
  // Sunday -> 6 working days.
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsNetworkdaysIntl, AllWeekendMaskRejected) {
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,14),\"1111111\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsNetworkdaysIntl, InvalidSelectorIsNum) {
  // 8 is not a valid paired-weekend code; 11..17 reserved for singles.
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,14),8)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsNetworkdaysIntl, InlineArrayHolidays) {
  // Jan 1 + Jan 2 as holidays, default Sat+Sun weekend. Jan 1..Jan 5
  // has 5 weekdays; minus 2 -> 3.
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1,{45292,45293})");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsNetworkdaysIntl, RangeHolidays) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(45292.0));  // 2024-01-01
  wb.sheet(0).set_cell_value(1, 0, Value::number(45293.0));  // 2024-01-02
  const Value v = EvalSourceIn("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,5),1,A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsNetworkdaysIntl, ReversedRangeNegates) {
  const Value v = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,12),DATE(2024,1,1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -10.0);
}

TEST(BuiltinsNetworkdaysIntl, ErrorPropagates) {
  const Value v = EvalSource("=NETWORKDAYS.INTL(\"abc\",DATE(2024,1,5))");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// WORKDAY.INTL
// ---------------------------------------------------------------------------

TEST(BuiltinsWorkdayIntl, DefaultMatchesWorkday) {
  const Value a = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5)");
  const Value b = EvalSource("=WORKDAY(DATE(2024,1,1),5)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsWorkdayIntl, ForwardWithExplicitWeekend1) {
  // Mon Jan 1 + 5 working days (Sat+Sun weekend) -> Mon Jan 8.
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5,1)");
  const Value exp = EvalSource("=DATE(2024,1,8)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(BuiltinsWorkdayIntl, StringMaskSameAsSelector1) {
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5,\"0000011\")");
  const Value exp = EvalSource("=DATE(2024,1,8)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(BuiltinsWorkdayIntl, SundayOnlyWeekend) {
  // Single-day weekend (Sun only). Mon Jan 1 + 5 working days, skipping
  // only Sunday Jan 7 -> Sat Jan 6 is a working day, so:
  //   Day 1: Tue Jan 2, Day 2: Wed Jan 3, Day 3: Thu Jan 4,
  //   Day 4: Fri Jan 5, Day 5: Sat Jan 6. Result: Sat Jan 6.
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5,11)");
  const Value exp = EvalSource("=DATE(2024,1,6)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(BuiltinsWorkdayIntl, WithHolidays) {
  // Same setup as WORKDAY's HolidaysExtendForward test.
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5,1,{45293,45294})");
  const Value exp = EvalSource("=DATE(2024,1,10)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(BuiltinsWorkdayIntl, ZeroDaysReturnsStart) {
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,3),0,1)");
  const Value exp = EvalSource("=DATE(2024,1,3)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(BuiltinsWorkdayIntl, AllWeekendStringRejected) {
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5,\"1111111\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWorkdayIntl, InvalidSelectorIsNum) {
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,1),5,8)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsWorkdayIntl, BackwardDays) {
  // Mon Jan 15 - 5 working days (Sat+Sun weekend) -> Mon Jan 8.
  const Value v = EvalSource("=WORKDAY.INTL(DATE(2024,1,15),-5,1)");
  const Value exp = EvalSource("=DATE(2024,1,8)");
  ASSERT_TRUE(v.is_number());
  ASSERT_TRUE(exp.is_number());
  EXPECT_EQ(v.as_number(), exp.as_number());
}

TEST(BuiltinsWorkdayIntl, ErrorPropagates) {
  const Value v = EvalSource("=WORKDAY.INTL(\"abc\",5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Weekday-rotation cross-check: 2024-01-01 must be a Monday. If this
// assertion ever fails, the Mon=0 rotation inside `is_weekend_masked`
// has drifted and every INTL variant will return wrong counts.
// ---------------------------------------------------------------------------

TEST(BuiltinsWorkdayIntl, Jan012024IsMonday) {
  // Using selector 12 = "Mon only" weekend: 2024-01-01 is a Monday, so
  // a same-day count should be 0; a 2024-01-02 interval should have 1.
  const Value a = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,1),12)");
  ASSERT_TRUE(a.is_number());
  EXPECT_EQ(a.as_number(), 0.0);
  const Value b = EvalSource("=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,2),12)");
  ASSERT_TRUE(b.is_number());
  EXPECT_EQ(b.as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// NOW / TODAY
//
// These are time-dependent; compare against a freshly-read local wall
// clock with a generous tolerance to stay robust under CI scheduling
// jitter.
// ---------------------------------------------------------------------------

TEST(BuiltinsNow, WithinFiveSecondsOfWallClock) {
  const LocalNowExpect exp = ExpectedLocalNow();
  const double expected_serial = exp.date_serial + exp.tod_fraction;
  const Value v = EvalSource("=NOW()");
  ASSERT_TRUE(v.is_number());
  // 5-second tolerance in serial-day units.
  EXPECT_NEAR(v.as_number(), expected_serial, 5.0 / 86400.0);
}

TEST(BuiltinsNow, ZeroArgsArity) {
  // NOW takes no arguments; any arg is an arity violation.
  const Value v = EvalSource("=NOW(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsToday, IsFloorOfNow) {
  // floor(NOW()) must equal TODAY() whenever both were read in the same
  // calendar second; use a 1-day tolerance as a safety net for the
  // extremely unlikely midnight-crossing race.
  const LocalNowExpect exp = ExpectedLocalNow();
  const Value v = EvalSource("=TODAY()");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), exp.date_serial);
  // And the floor relationship, reading NOW again within the same test.
  const Value now = EvalSource("=NOW()");
  ASSERT_TRUE(now.is_number());
  EXPECT_EQ(std::floor(now.as_number()), v.as_number());
}

TEST(BuiltinsToday, ZeroArgsArity) {
  const Value v = EvalSource("=TODAY(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// AREAS
// ---------------------------------------------------------------------------

TEST(BuiltinsAreas, SingleCellIsOne) {
  const Value v = EvalSource("=AREAS(A1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsAreas, RangeIsOne) {
  const Value v = EvalSource("=AREAS(A1:B2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsAreas, LiteralArgumentIsValueError) {
  // A numeric literal is not a reference.
  const Value v = EvalSource("=AREAS(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAreas, ArithmeticIsValueError) {
  const Value v = EvalSource("=AREAS(1+2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAreas, TooFewArgsIsValue) {
  const Value v = EvalSource("=AREAS()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsAreas, TooManyArgsIsValue) {
  const Value v = EvalSource("=AREAS(A1,B2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// Excel's space-as-intersection operator: overlapping rectangles count as a
// single area. `A1:C3 B2:D4` overlaps at `B2:C3`, so AREAS reports 1.
TEST(BuiltinsAreas, IntersectingRangesIsOneArea) {
  const Value v = EvalSource("=AREAS(A1:C3 B2:D4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

// Disjoint operands intersect to nothing -> Excel surfaces `#NULL!`.
TEST(BuiltinsAreas, DisjointRangesIsNullError) {
  const Value v = EvalSource("=AREAS(A1:B2 C3:D4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Null);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
