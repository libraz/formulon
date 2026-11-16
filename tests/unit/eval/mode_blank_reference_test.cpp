//
// MODE, MODE.SNGL and MODE.MULT answer #VALUE! for a direct reference to a blank cell in any position, while a blank
// cell inside a range and a text cell are skipped; MEDIAN skips the blank reference.

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

// C1 holds "x", D1 holds 2; B2 and B3 are blank.
Value Eval(std::string_view src) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 2, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 3, Value::number(2.0));
  return EvalSourceIn(src, wb, wb.sheet(0));
}

TEST(ModeBlankReference, ABlankCellReferenceIsValue) {
  for (const char* src :
       {"=MODE(B2,2,2)", "=MODE(D1,B2,2)", "=MODE(2,2,B2)", "=MODE.SNGL(B2,2,2)", "=MODE.MULT(B2,2,2)"}) {
    const Value v = Eval(src);
    ASSERT_TRUE(v.is_error()) << src;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << src;
  }
}

TEST(ModeBlankReference, ARangeOrATextCellIsSkipped) {
  for (const char* src : {"=MODE(B2:B3,2,2)", "=MODE(C1,2,2)", "=MODE(C1:C2,2,2)"}) {
    const Value v = Eval(src);
    ASSERT_TRUE(v.is_number()) << src;
    EXPECT_EQ(v.as_number(), 2.0) << src;
  }
}

TEST(ModeBlankReference, MedianSkipsTheBlankReference) {
  const Value v = Eval("=MEDIAN(B2,2,3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.5);
}

}  // namespace
}  // namespace formulon
