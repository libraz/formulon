// HOUR / MINUTE / SECOND bound their serial by the workbook's own last day,
// like YEAR: 9999-12-31 is serial 2957003 in the 1904 date system.

#include "builtins_datetime_test_helpers.h"

namespace formulon {
namespace eval {
namespace {

TEST(TimeParts1904, LastDayOfTheEpochEvaluates) {
  EXPECT_EQ(EvalSource1904("=HOUR(2957003.5)").as_number(), 12.0);
  EXPECT_EQ(EvalSource1904("=MINUTE(2957003.75)").as_number(), 0.0);
  EXPECT_EQ(EvalSource1904("=SECOND(2957003.5)").as_number(), 0.0);
}

TEST(TimeParts1904, SerialPastTheEpochLastDayIsNum) {
  for (const char* formula : {"=HOUR(2957004.5)", "=MINUTE(2957004.5)", "=SECOND(2957004.5)", "=HOUR(2958465.5)"}) {
    const Value v = EvalSource1904(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << formula;
  }
}

TEST(TimeParts1904, The1900BoundIsUnchanged) {
  EXPECT_EQ(EvalSource("=HOUR(2958465.5)").as_number(), 12.0);
  const Value past = EvalSource("=HOUR(2958466.5)");
  ASSERT_TRUE(past.is_error());
  EXPECT_EQ(past.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
