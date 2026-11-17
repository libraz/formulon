//
// WEEKNUM lifts element by element over an array argument, as ISOWEEKNUM does, but a range argument is #VALUE!
// where ISOWEEKNUM lifts it.

#include <string_view>

#include "gtest/gtest.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using formulon::test::EvalSourceIn;

// A1:A3 hold 45000, 45010 and 45020.
Value Eval(std::string_view src) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(45000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(45010.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(45020.0));
  return EvalSourceIn(src, wb, wb.sheet(0));
}

double Number(std::string_view src) {
  const Value v = Eval(src);
  EXPECT_TRUE(v.is_number()) << src;
  return v.is_number() ? v.as_number() : 0.0;
}

TEST(WeeknumArrayLift, AnArrayLifts) {
  EXPECT_EQ(Number("=COLUMNS(WEEKNUM({45000,45010}))"), 2.0);
  EXPECT_EQ(Number("=INDEX(WEEKNUM({45000,45010}),1)"), 11.0);
  EXPECT_EQ(Number("=INDEX(WEEKNUM({45000,45010}),2)"), 12.0);
  EXPECT_EQ(Number("=INDEX(WEEKNUM({366,366},2),2)"), Number("=WEEKNUM(366,2)"));
  EXPECT_EQ(Number("=INDEX(ISOWEEKNUM({45000,45010}),2)"), 12.0);
}

TEST(WeeknumArrayLift, ARangeIsValueWhereIsoweeknumLifts) {
  for (const char* src : {"=WEEKNUM(A1:A3)", "=SUM(WEEKNUM(A1:A3))", "=SUM(WEEKNUM(A1:A3,2))"}) {
    const Value v = Eval(src);
    ASSERT_TRUE(v.is_error()) << src;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << src;
  }
  EXPECT_EQ(Number("=SUM(ISOWEEKNUM(A1:A3))"), 37.0);
}

}  // namespace
}  // namespace formulon
