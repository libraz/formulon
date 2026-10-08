// Reference builtin tests grouped by reference constructor.

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "builtins_references_test_helpers.h"

namespace formulon {
namespace eval {
namespace {
using namespace references_test_helpers;

TEST(BuiltinsAddress, DefaultAbsoluteA1) {
  const Value v = EvalSource("=ADDRESS(1,1)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$A$1");
}

TEST(BuiltinsAddress, AbsoluteRowRelativeColumn) {
  // abs_num = 2 -> A$1 (row absolute, column relative)
  const Value v = EvalSource("=ADDRESS(1,1,2)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "A$1");
}

TEST(BuiltinsAddress, RelativeRowAbsoluteColumn) {
  // abs_num = 3 -> $A1
  const Value v = EvalSource("=ADDRESS(1,1,3)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$A1");
}

TEST(BuiltinsAddress, BothRelative) {
  // abs_num = 4 -> A1
  const Value v = EvalSource("=ADDRESS(1,1,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "A1");
}

TEST(BuiltinsAddress, InvalidAbsNumIsValueError) {
  EXPECT_TRUE(EvalSource("=ADDRESS(1,1,0)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,1,5)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,1,-1)").is_error());
}

TEST(BuiltinsAddress, ColumnLettersBoundary) {
  // 26 -> Z
  const Value v26 = EvalSource("=ADDRESS(1,26,4)");
  ASSERT_TRUE(v26.is_text());
  EXPECT_EQ(std::string(v26.as_text()), "Z1");
  // 27 -> AA
  const Value v27 = EvalSource("=ADDRESS(1,27,4)");
  ASSERT_TRUE(v27.is_text());
  EXPECT_EQ(std::string(v27.as_text()), "AA1");
  // 702 -> ZZ
  const Value v702 = EvalSource("=ADDRESS(1,702,4)");
  ASSERT_TRUE(v702.is_text());
  EXPECT_EQ(std::string(v702.as_text()), "ZZ1");
  // 703 -> AAA
  const Value v703 = EvalSource("=ADDRESS(1,703,4)");
  ASSERT_TRUE(v703.is_text());
  EXPECT_EQ(std::string(v703.as_text()), "AAA1");
  // 16384 -> XFD (Excel's last column)
  const Value vMax = EvalSource("=ADDRESS(1,16384,4)");
  ASSERT_TRUE(vMax.is_text());
  EXPECT_EQ(std::string(vMax.as_text()), "XFD1");
}

TEST(BuiltinsAddress, RowAndColumnOutOfRange) {
  EXPECT_TRUE(EvalSource("=ADDRESS(0,1)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1048577,1)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,0)").is_error());
  EXPECT_TRUE(EvalSource("=ADDRESS(1,16385)").is_error());
}

TEST(BuiltinsAddress, R1C1AbsoluteStyle) {
  // a1 = FALSE, abs_num = 1 -> R1C1
  const Value v = EvalSource("=ADDRESS(1,1,1,FALSE)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "R1C1");
}

TEST(BuiltinsAddress, R1C1MixedStyles) {
  // abs_num = 2 -> R1C[1]
  const Value v2 = EvalSource("=ADDRESS(1,1,2,FALSE)");
  ASSERT_TRUE(v2.is_text());
  EXPECT_EQ(std::string(v2.as_text()), "R1C[1]");
  // abs_num = 3 -> R[1]C1
  const Value v3 = EvalSource("=ADDRESS(1,1,3,FALSE)");
  ASSERT_TRUE(v3.is_text());
  EXPECT_EQ(std::string(v3.as_text()), "R[1]C1");
  // abs_num = 4 -> R[1]C[1]
  const Value v4 = EvalSource("=ADDRESS(1,1,4,FALSE)");
  ASSERT_TRUE(v4.is_text());
  EXPECT_EQ(std::string(v4.as_text()), "R[1]C[1]");
}

TEST(BuiltinsAddress, SheetPrefixBare) {
  const Value v = EvalSource("=ADDRESS(1,1,1,TRUE,\"Sheet1\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "Sheet1!$A$1");
}

TEST(BuiltinsAddress, SheetPrefixWithSpaceQuoted) {
  const Value v = EvalSource("=ADDRESS(1,1,1,TRUE,\"Sheet 1\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "'Sheet 1'!$A$1");
}

TEST(BuiltinsAddress, SheetPrefixWithApostropheEscaped) {
  const Value v = EvalSource("=ADDRESS(1,1,4,TRUE,\"O'Brien\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "'O''Brien'!A1");
}

TEST(BuiltinsAddress, SheetPrefixLeadingDigitQuoted) {
  const Value v = EvalSource("=ADDRESS(1,1,4,TRUE,\"2020\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "'2020'!A1");
}

TEST(BuiltinsAddress, SheetPrefixEmptyStillEmitsSeparator) {
  // Excel 365 emits the trailing `!` whenever the sheet_text arg is
  // supplied, even if the name is an empty string. Matches oracle.
  const Value v = EvalSource("=ADDRESS(1,1,4,TRUE,\"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "!A1");
}

TEST(BuiltinsAddress, LargeRowAndColumn) {
  const Value v = EvalSource("=ADDRESS(100,50,4)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "AX100");
}

TEST(BuiltinsAddress, AbsoluteWithSheet) {
  const Value v = EvalSource("=ADDRESS(10,5,1,TRUE,\"Data\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "Data!$E$10");
}

TEST(BuiltinsAddress, R1C1WithSheet) {
  const Value v = EvalSource("=ADDRESS(10,5,1,FALSE,\"Data\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "Data!R10C5");
}

TEST(BuiltinsAddress, ArityBelowMinIsError) {
  EXPECT_TRUE(EvalSource("=ADDRESS(1)").is_error());
}

TEST(BuiltinsAddress, ErrorArgumentPropagates) {
  EXPECT_TRUE(EvalSource("=ADDRESS(1/0,1)").is_error());
}

TEST(BuiltinsAddress, NumericCoercionOnRowCol) {
  // Fractional rows truncate toward zero (matches std::trunc).
  const Value v = EvalSource("=ADDRESS(3.7, 2.9)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(std::string(v.as_text()), "$B$3");
}

TEST(BuiltinsAddress, RowSnapsToNearInteger) {
  ExpectText(EvalSource("=ADDRESS(1.9999999,1)"), "$A$2", "=ADDRESS(1.9999999,1)");
  ExpectText(EvalSource("=ADDRESS(1.999999,1)"), "$A$1", "=ADDRESS(1.999999,1)");
}

TEST(BuiltinsAddress, ColumnSnapsToNearInteger) {
  ExpectText(EvalSource("=ADDRESS(1,1.9999999)"), "$B$1", "=ADDRESS(1,1.9999999)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
