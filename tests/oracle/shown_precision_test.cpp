//
// Unit tests for `shown_half_unit`, the precision a number's displayed text
// carries, which bounds the `numeric_text` comparison from below.

#include "tests/oracle/shown_precision.h"

#include "gtest/gtest.h"

namespace formulon::tests::oracle {
namespace {

TEST(ShownHalfUnit, DecimalTextGivesHalfItsLastDigit) {
  EXPECT_DOUBLE_EQ(shown_half_unit("-471.7808"), 5e-5);
  EXPECT_DOUBLE_EQ(shown_half_unit("1.3151"), 5e-5);
  EXPECT_DOUBLE_EQ(shown_half_unit("+.5"), 0.05);
  EXPECT_DOUBLE_EQ(shown_half_unit(" 3.25 "), 0.005);
  EXPECT_DOUBLE_EQ(shown_half_unit("1,5"), 0.05);
}

TEST(ShownHalfUnit, IntegerTextGivesHalfAUnit) {
  EXPECT_DOUBLE_EQ(shown_half_unit("2100"), 0.5);
  EXPECT_DOUBLE_EQ(shown_half_unit("-7"), 0.5);
}

TEST(ShownHalfUnit, ExponentShiftsTheLastDigit) {
  EXPECT_DOUBLE_EQ(shown_half_unit("1.22514845490862E-16"), 5e-31);
  EXPECT_DOUBLE_EQ(shown_half_unit("1e3"), 500.0);
  EXPECT_DOUBLE_EQ(shown_half_unit("2.5E+10"), 5e8);
}

TEST(ShownHalfUnit, NonNumericTextGivesZero) {
  for (const char* text : {"", "-", ".", "abc", "1.2.3", "1e", "1e+", "12abc", "#N/A", "1 2"}) {
    EXPECT_EQ(shown_half_unit(text), 0.0) << text;
  }
}

}  // namespace
}  // namespace formulon::tests::oracle
