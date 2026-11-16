//
// A coupon schedule whose previous coupon date falls before the date system's first day: the bond functions answer
// #NUM!, COUPPCD answers 0, and the discount functions, which have no schedule, are unaffected.

#include <string_view>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace {

using formulon::test::EvalSource;

ErrorCode Error(std::string_view src) {
  const Value v = EvalSource(src);
  EXPECT_TRUE(v.is_error()) << src;
  return v.is_error() ? v.as_error() : ErrorCode::Null;
}

double Number(std::string_view src) {
  const Value v = EvalSource(src);
  EXPECT_TRUE(v.is_number()) << src;
  return v.is_number() ? v.as_number() : 0.0;
}

TEST(CouponCalendarFloor, ACouponDateBefore1900IsNum) {
  EXPECT_EQ(Error("=PRICE(0,DATE(2017,11,15),0.0575,0.065,100,2,0)"), ErrorCode::Num);
  EXPECT_EQ(Error("=PRICE(1,DATE(1900,6,15),0.0575,0.065,100,2,0)"), ErrorCode::Num);
  EXPECT_EQ(Error("=PRICE(DATE(1900,3,1),DATE(1900,6,15),0.0575,0.065,100,2,0)"), ErrorCode::Num);
  EXPECT_EQ(Error("=YIELD(0,DATE(2016,11,15),0.0575,95.04287,100,2,0)"), ErrorCode::Num);
  EXPECT_EQ(Error("=DURATION(0,DATE(2016,1,1),0.08,0.09,2,1)"), ErrorCode::Num);
  EXPECT_EQ(Error("=MDURATION(0,DATE(2016,1,1),0.08,0.09,2,1)"), ErrorCode::Num);
  EXPECT_EQ(Error("=COUPNUM(1,DATE(1900,6,15),2,0)"), ErrorCode::Num);
  EXPECT_EQ(Error("=ACCRINT(0,DATE(2024,3,31),DATE(2024,9,15),0.05,1000,1,0,TRUE)"), ErrorCode::Num);
}

TEST(CouponCalendarFloor, ACouponDateOnTheFirstDayIsKept) {
  EXPECT_NEAR(Number("=PRICE(1,DATE(1900,7,1),0.0575,0.065,100,2,0)"), 99.6368039, 1e-7);
}

TEST(CouponCalendarFloor, CouppcdAnswersZero) {
  EXPECT_EQ(Number("=COUPPCD(1,DATE(1900,6,15),2,0)"), 0.0);
  EXPECT_EQ(Number("=COUPPCD(0,DATE(2016,1,1),2,1)"), 0.0);
}

TEST(CouponCalendarFloor, AFunctionWithoutASchedule) {
  EXPECT_NEAR(Number("=PRICEDISC(0,DATE(2018,3,1),0.0525,100,2)"), -529.41667, 1e-5);
}

}  // namespace
}  // namespace formulon
