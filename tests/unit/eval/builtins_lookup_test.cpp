// Lookup builtin tests grouped by function.

#include <cstdint>
#include <string>
#include <string_view>

#include "builtins_lookup_test_helpers.h"
#include "eval/builtins.h"
#include "util/test_eval_helpers.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;
using namespace lookup_test_helpers;

TEST(BuiltinsMatch, ExactNumericFirstHit) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(20.0));  // duplicate, should be ignored
  const Value v = EvalSourceIn("=MATCH(20, A1:A4, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactFullwidthDigitsFollowHostFolding) {
  struct ProfileCase {
    ExcelProfile profile;
    bool folds_fullwidth_digits;
  };
  const ProfileCase cases[] = {
      {mac_365_ja_jp_profile(), true},
      {mac_365_en_us_profile(), true},
      {win_365_ja_jp_profile(), false},
      {win_365_en_us_profile(), false},
  };

  for (const ProfileCase& test_case : cases) {
    SCOPED_TRACE(excel_profile_id(test_case.profile));
    Workbook wb = Workbook::create();
    wb.set_excel_profile(test_case.profile);
    const Value v = EvalSourceIn("=MATCH(\"１２３\",{\"123\";\"456\"},0)", wb, wb.sheet(0));
    if (test_case.folds_fullwidth_digits) {
      ASSERT_TRUE(v.is_number());
      EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
    } else {
      ASSERT_TRUE(v.is_error());
      EXPECT_EQ(v.as_error(), ErrorCode::NA);
    }
  }
}

TEST(BuiltinsMatch, ExactMatchArrayLiteralUsesSharedRangeMaterialization) {
  const Value v = EvalSource("=MATCH(20,{10;20;30},0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, FractionalMatchTypeTruncatesTowardZeroLikeXmatch) {
  // Every integer argument in the lookup family narrows through the shared
  // truncate-toward-zero helper, so `-0.5` is mode 0 (exact) for MATCH just
  // as it is for XMATCH — never mode -1 via a floor.
  const Value match = EvalSource("=MATCH(2,{1;2;3},-0.5)");
  ASSERT_TRUE(match.is_number()) << "a fractional match_type must not flip MATCH into descending mode";
  EXPECT_DOUBLE_EQ(match.as_number(), 2.0);

  const Value xmatch = EvalSource("=XMATCH(2,{1;2;3},-0.5)");
  ASSERT_TRUE(xmatch.is_number());
  EXPECT_DOUBLE_EQ(xmatch.as_number(), match.as_number());
}

TEST(BuiltinsMatch, ArrayLookupValueExactPreservesShapeAndErrors) {
  const Value v = EvalSource("=MATCH({20;99;#DIV/0!},{10;20;30},0)");
  ExpectArrayShape(v, 3U, 1U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_number());
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 2.0);
  ASSERT_TRUE(cells[1].is_error());
  EXPECT_EQ(cells[1].as_error(), ErrorCode::NA);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMatch, ArrayLookupValueNestedSequenceMapsOnce) {
  const Value v = EvalSource("=MATCH(SEQUENCE(1,3,25,-10),SEQUENCE(3,1,10,10),1)");
  ExpectArrayShape(v, 1U, 3U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_number());
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 2.0);
  ASSERT_TRUE(cells[1].is_number());
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 1.0);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, ArrayLookupValueModeTwoUsesAscendingCoercion) {
  const Value v = EvalSource("=MATCH({20;#DIV/0!},{10;20;30},2)");
  ExpectArrayShape(v, 2U, 1U);
  const Value* cells = v.as_array_cells();
  ASSERT_TRUE(cells[0].is_number());
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 2.0);
  ASSERT_TRUE(cells[1].is_error());
  EXPECT_EQ(cells[1].as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMatch, ExactZeroDoesNotMatchBlankCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  // A2 remains empty. MATCH exact mode excludes blank array cells from a
  // numeric lookup; it must not coerce the blank to zero.
  const Value v = EvalSourceIn("=MATCH(0, A1:A2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, ExactZeroSkipsBlankBeforeRealZero) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::number(0.0));
  // A1 remains empty; only the literal zero in A2 can match.
  const Value v = EvalSourceIn("=MATCH(0, A1:A2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactTextCaseInsensitive) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("apple"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("CHERRY"));
  const Value v = EvalSourceIn("=MATCH(\"banana\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactWildcardStar) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("Apricot"));
  const Value v = EvalSourceIn("=MATCH(\"A*\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactWildcardQuestion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("cab"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("ab"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("abc"));
  // "?b" matches exactly one byte before 'b' -> "ab" doesn't match (no
  // leading byte) but "cb" would; we want "?b" to match a 2-byte string
  // whose 2nd char is 'b'. Only "ab" is length 2 with b at index 1, which
  // matches "?b" (? eats "a"). "cab" is 3 bytes so no match.
  const Value v = EvalSourceIn("=MATCH(\"?b\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactWildcardEscape) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("foo"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("*"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("bar"));
  // "~*" is the literal asterisk.
  const Value v = EvalSourceIn("=MATCH(\"~*\", A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, ExactNoMatchIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=MATCH(99, A1:A2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, AscendingExactHit) {
  // Ascending array, target exactly matches a cell.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(5, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, AscendingBetweenCells) {
  // Target falls between cells: return the position whose value is
  // largest but still <= target.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(7, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, AscendingAllGreaterIsNa) {
  // Every cell is strictly greater than the target -> #N/A.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  const Value v = EvalSourceIn("=MATCH(1, A1:A3, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, DescendingSmallestGreaterOrEqual) {
  // Descending array: return largest position whose value is >= target.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(15, A1:A3, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, DescendingAllLessIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=MATCH(10, A1:A3, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, TwoDimensionalRangeIsNa) {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(static_cast<double>(r * 2 + c)));
    }
  }
  const Value v = EvalSourceIn("=MATCH(1, A1:B2, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, OmittedMatchTypeDefaultsToAscending) {
  // Default match_type is 1; omitting it should behave identically.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=MATCH(7, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, LookupValueErrorPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=MATCH(#DIV/0!, A1:A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMatch, SingleCellRefAsLookupArray) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value hit = EvalSourceIn("=MATCH(42, A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(hit.is_number());
  EXPECT_DOUBLE_EQ(hit.as_number(), 1.0);
  const Value miss = EvalSourceIn("=MATCH(99, A1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(miss.is_error());
  EXPECT_EQ(miss.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, CrossSheetLookupArray) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::text("alpha"));
  wb.sheet(1).set_cell_value(1, 0, Value::text("beta"));
  wb.sheet(1).set_cell_value(2, 0, Value::text("gamma"));
  const Value v = EvalSourceIn("=MATCH(\"beta\", Data!A1:A3, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMatch, InvalidMatchTypeIsNa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=MATCH(1, A1:A1, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMatch, WrongArityIsValueError) {
  EXPECT_EQ(EvalSource("=MATCH(1)").as_error(), ErrorCode::Value);
}

TEST(BuiltinsLookup, VectorFormExactMatch) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("one"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("two"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 1, Value::text("three"));
  const Value v = EvalSourceIn("=LOOKUP(3, A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "three");
}

TEST(BuiltinsLookup, VectorFormApproximateBetween) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::text("c"));
  // 4 falls between 3 and 5; the last cell with lookup <= 4 is row 2 -> "b".
  const Value v = EvalSourceIn("=LOOKUP(4, A1:A3, B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsLookup, VectorFormBelowFirstIsNA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("y"));
  const Value v = EvalSourceIn("=LOOKUP(1, A1:A2, B1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsLookup, VectorFormAboveAllClampsToLastMatch) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("b"));
  const Value v = EvalSourceIn("=LOOKUP(99, A1:A2, B1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "b");
}

TEST(BuiltinsLookup, ErrorInLookupValuePropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("a"));
  const Value v = EvalSourceIn("=LOOKUP(1/0, A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsLookup, BlankLookupValueIsNA) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("a"));
  // A99 is blank; LOOKUP of blank surfaces `#N/A` even when the vector is
  // numeric and could otherwise compare blank as 0.
  const Value v = EvalSourceIn("=LOOKUP(A99, A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsLookup, TwoArgFormReturnsLookupVectorCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  // No result-vector arg: the matched lookup cell itself is the result.
  const Value v = EvalSourceIn("=LOOKUP(2.5, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
