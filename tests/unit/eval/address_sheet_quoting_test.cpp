// ADDRESS quotes its sheet_text by the same rule as an A1 sheet qualifier in
// formula display, in both reference styles (Mac Excel 365 ja-JP).

#include <string>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;

void ExpectText(const std::string& formula, const std::string& expected) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_text()) << formula;
  EXPECT_EQ(v.as_text(), expected) << formula;
}

TEST(AddressSheetQuoting, QuotesCellAndR1C1AndBoolShapes) {
  ExpectText("=ADDRESS(1,1,1,TRUE,\"A1\")", "'A1'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"A01\")", "'A01'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"XFD1048576\")", "'XFD1048576'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"R1C1\")", "'R1C1'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"rc\")", "'rc'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"TRUE\")", "'TRUE'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"1abc\")", "'1abc'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"a-b\")", "'a-b'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"a'b\")", "'a''b'!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"\xE5\x85\xA8\xE3\x80\x80\xE8\xA7\x92\")",
             "'\xE5\x85\xA8\xE3\x80\x80\xE8\xA7\x92'!$A$1");  // 全　角
}

TEST(AddressSheetQuoting, LeavesBareNamesUnquoted) {
  ExpectText("=ADDRESS(1,1,1,TRUE,\"a.b\")", "a.b!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"a_b\")", "a_b!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"XFE1\")", "XFE1!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"A0\")", "A0!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"A0001048577\")", "A0001048577!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"\")", "!$A$1");
  ExpectText("=ADDRESS(1,1,1,TRUE,\"\xE6\x97\xA5\xE6\x9C\xAC\")", "\xE6\x97\xA5\xE6\x9C\xAC!$A$1");  // 日本
}

TEST(AddressSheetQuoting, R1C1StyleKeepsTheA1Rule) {
  ExpectText("=ADDRESS(1,1,1,FALSE,\"A01\")", "'A01'!R1C1");
  ExpectText("=ADDRESS(1,1,1,FALSE,\"A1\")", "'A1'!R1C1");
  ExpectText("=ADDRESS(1,1,1,FALSE,\"a.b\")", "a.b!R1C1");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
