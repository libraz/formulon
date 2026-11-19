//
// Pictures in formulas: every engine-evaluated row of the Excel-measured
// picture semantics, the CalcOrigin a formula result is saved with, and the
// pinned identity and sort rules. Expected values are the goldens in
// tests/fixtures/excel/cell_pictures/picture_formula_semantics.json
// (mac-365-ja_JP); the books there are rebuilt here cell by cell.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "eval/cell_evaluator.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/text_format/display_text.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "parser/parser.h"
#include "picture_table.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint8_t kPlaced = 5;
constexpr std::uint8_t kFormula = kFormulaResultPictureOrigin;
constexpr std::string_view kBlueAlt = "Blue circle on a blue field";

::testing::AssertionResult IsPicture(const Value& v, PicturePayload p, std::uint8_t origin) {
  if (!v.is_picture()) {
    return ::testing::AssertionFailure() << "kind " << static_cast<int>(v.kind()) << ", not a picture";
  }
  const PicturePayload got = v.as_picture();
  if (got.picture_id != p.picture_id || got.media_id != p.media_id || v.picture_origin() != origin) {
    return ::testing::AssertionFailure() << "picture {" << got.picture_id << ", " << got.media_id << "} origin "
                                         << static_cast<int>(v.picture_origin()) << ", want {" << p.picture_id << ", "
                                         << p.media_id << "} origin " << static_cast<int>(origin);
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult IsError(const Value& v, ErrorCode code) {
  if (!v.is_error() || v.as_error() != code) {
    return ::testing::AssertionFailure() << "kind " << static_cast<int>(v.kind()) << ", want error "
                                         << static_cast<int>(code);
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult IsNumber(const Value& v, double d) {
  if (!v.is_number() || v.as_number() != d) {
    return ::testing::AssertionFailure() << "kind " << static_cast<int>(v.kind()) << ", want number " << d;
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult IsBool(const Value& v, bool b) {
  if (!v.is_boolean() || v.as_boolean() != b) {
    return ::testing::AssertionFailure() << "kind " << static_cast<int>(v.kind()) << ", want " << b;
  }
  return ::testing::AssertionSuccess();
}

::testing::AssertionResult IsText(const Value& v, std::string_view s) {
  if (!v.is_text() || v.as_text() != s) {
    return ::testing::AssertionFailure() << "kind " << static_cast<int>(v.kind()) << ", want text " << s;
  }
  return ::testing::AssertionSuccess();
}

// Row and column of an A1 address such as "E2".
std::pair<std::uint32_t, std::uint32_t> Cell(std::string_view a1) {
  std::uint32_t col = 0;
  std::size_t i = 0;
  for (; i < a1.size() && a1[i] >= 'A' && a1[i] <= 'Z'; ++i) {
    col = col * 26U + static_cast<std::uint32_t>(a1[i] - 'A' + 1);
  }
  std::uint32_t row = 0;
  for (; i < a1.size(); ++i) {
    row = row * 10U + static_cast<std::uint32_t>(a1[i] - '0');
  }
  return {row - 1U, col - 1U};
}

class PictureValue : public ::testing::Test {
 protected:
  void SetUp() override { wb_.set_excel_profile(mac_365_ja_jp_profile()); }

  // Adds a media at ascending picture rank `rank`; the ranks at and above it move up.
  std::uint32_t AddMedia(std::size_t rank, std::uint64_t size, std::uint64_t digest) {
    return wb_.picture_table().add_media_at(rank, MediaKey{size, digest});
  }

  PicturePayload AddPicture(std::uint32_t media_id, std::optional<std::string> alt = std::nullopt) {
    return PicturePayload{wb_.picture_table().add_entry(media_id, std::move(alt)), media_id};
  }

  void Put(std::string_view a1, Value v) {
    const auto [row, col] = Cell(a1);
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_value(0, row, col, v))) << a1;
  }

  void PutFormula(std::string_view a1, std::string_view formula) {
    const auto [row, col] = Cell(a1);
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, row, col, std::string(formula)))) << a1;
  }

  // The main book: E2 = E4 = image1, E6 = image2, E8 = image2 with an alt, E3/E5/E7 blank, A1 = 5, A2 = "x".
  // image2 sorts before image1.
  void BuildMain() {
    image1_media_ = AddMedia(0, 463U, 0xea3c443bdbd00ff7ULL);
    image2_media_ = AddMedia(0, 1804U, 0x523d3d5e8b3f742dULL);
    image1_ = AddPicture(image1_media_);
    image2_ = AddPicture(image2_media_);
    image2_alt_ = AddPicture(image2_media_, std::string(kBlueAlt));
    Put("A1", Value::number(5.0));
    Put("A2", Value::text("x"));
    Put("E2", Value::picture(image1_, kPlaced));
    Put("E4", Value::picture(image1_, kPlaced));
    Put("E6", Value::picture(image2_, kPlaced));
    Put("E8", Value::picture(image2_alt_, kPlaced));
  }

  // The identity book: the main book plus P2 (image1's bytes, a second entry), P3 (image1's pixels in other bytes),
  // and image1 with alts "Picture" (P4, P6) and "PICTURE" (P5).
  void BuildIdentity() {
    BuildMain();
    const std::uint32_t reencoded = AddMedia(2, 376U, 0x178f562bd11bd27bULL);
    Put("P2", Value::picture(AddPicture(image1_media_), kPlaced));
    Put("P3", Value::picture(AddPicture(reencoded), kPlaced));
    p4_ = AddPicture(image1_media_, std::string("Picture"));
    Put("P4", Value::picture(p4_, kPlaced));
    Put("P5", Value::picture(AddPicture(image1_media_, std::string("PICTURE")), kPlaced));
    Put("P6", Value::picture(AddPicture(image1_media_, std::string("Picture")), kPlaced));
  }

  // The sort book: media B < C < A ascending; Q2 = C, Q3 = A, Q4 = B, Q5 = A alt "b", Q6 = A alt "a", Q7 blank.
  void BuildSort() {
    const std::uint32_t a = AddMedia(0, 211U, 0x2068963cbc5634b0ULL);
    const std::uint32_t c = AddMedia(0, 133U, 0x6250cf25cd87760dULL);
    const std::uint32_t b = AddMedia(0, 73U, 0x983fd423ae2b659cULL);
    sk_a_ = AddPicture(a);
    sk_b_ = AddPicture(b);
    sk_c_ = AddPicture(c);
    sk_a_alt_b_ = AddPicture(a, std::string("b"));
    sk_a_alt_a_ = AddPicture(a, std::string("a"));
    Put("Q2", Value::picture(sk_c_, kPlaced));
    Put("Q3", Value::picture(sk_a_, kPlaced));
    Put("Q4", Value::picture(sk_b_, kPlaced));
    Put("Q5", Value::picture(sk_a_alt_b_, kPlaced));
    Put("Q6", Value::picture(sk_a_alt_a_, kPlaced));
  }

  // Enters `formula` in a fresh column of row 1, recalculates and returns the anchor's value.
  Value Eval(std::string_view formula) {
    anchor_col_ += 4U;
    EXPECT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, 0U, anchor_col_, std::string(formula)))) << formula;
    EXPECT_TRUE(static_cast<bool>(wb_.recalc(default_registry()))) << formula;
    return Member(0U);
  }

  // A member of the last result's spill, `row` down and `col` across from the anchor.
  Value Member(std::uint32_t row, std::uint32_t col = 0U) const {
    return wb_.sheet(0).resolve_cell_value(row, anchor_col_ + col);
  }

  Workbook wb_ = Workbook::create();
  std::uint32_t anchor_col_ = 30U;
  std::uint32_t image1_media_ = 0;
  std::uint32_t image2_media_ = 0;
  PicturePayload image1_{};
  PicturePayload image2_{};
  PicturePayload image2_alt_{};
  PicturePayload p4_{};
  PicturePayload sk_a_{};
  PicturePayload sk_b_{};
  PicturePayload sk_c_{};
  PicturePayload sk_a_alt_b_{};
  PicturePayload sk_a_alt_a_{};
};

// --- Reads and information functions -------------------------------------

TEST_F(PictureValue, CellReadIsThePicture) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=E2"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=E8"), image2_alt_, kFormula));
}

TEST_F(PictureValue, TypeIs128) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=TYPE(E2)"), 128.0));
  EXPECT_TRUE(IsNumber(Eval("=TYPE(E8)"), 128.0));
}

TEST_F(PictureValue, IsTextIsNumberIsErrorAreFalse) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=ISTEXT(E2)"), false));
  EXPECT_TRUE(IsBool(Eval("=ISTEXT(E8)"), false));
  EXPECT_TRUE(IsBool(Eval("=ISNUMBER(E2)"), false));
  EXPECT_TRUE(IsBool(Eval("=ISERROR(E2)"), false));
}

TEST_F(PictureValue, IsErrIsLogicalIsNaAreFalse) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=ISERR(E2)"), false));
  EXPECT_TRUE(IsBool(Eval("=ISLOGICAL(E2)"), false));
  EXPECT_TRUE(IsBool(Eval("=ISNA(E2)"), false));
}

TEST_F(PictureValue, IsNonTextIsTrue) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=ISNONTEXT(E2)"), true));
}

TEST_F(PictureValue, IsBlankIsFalse) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=ISBLANK(E2)"), false));
}

TEST_F(PictureValue, IsRefIsTrue) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=ISREF(E2)"), true));
}

TEST_F(PictureValue, ErrorTypeIsNA) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=ERROR.TYPE(E2)"), ErrorCode::NA));
}

TEST_F(PictureValue, LenAndConcatenationAreValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=LEN(E2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=E2&\"\""), ErrorCode::Value));
}

TEST_F(PictureValue, CellTypeIsV) {
  BuildMain();
  EXPECT_TRUE(IsText(Eval("=CELL(\"type\",E2)"), "v"));
}

TEST_F(PictureValue, CountaCountsAndCountSkips) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=COUNTA(E2:E8)"), 4.0));
  EXPECT_TRUE(IsNumber(Eval("=COUNT(E2:E8)"), 0.0));
}

TEST_F(PictureValue, CountblankRowsIsformula) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=COUNTBLANK(E2:E4)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(E2)"), 1.0));
  EXPECT_TRUE(IsBool(Eval("=ISFORMULA(E2)"), false));
}

TEST_F(PictureValue, DisplayTextIsEmpty) {
  const text_format::DisplayText shown = text_format::format_value_for_display(
      Value::picture(PicturePayload{0U, 0U}, kPlaced), "General", false, mac_365_ja_jp_profile());
  EXPECT_EQ(shown.text, "");
}

// --- Comparison ----------------------------------------------------------

TEST_F(PictureValue, EqualsSamePictureIsTrue) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=E2=E4"), true));
}

TEST_F(PictureValue, EqualsDifferentImagesIsFalse) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=E2=E6"), false));
}

TEST_F(PictureValue, EqualsTextOrNumberIsFalse) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=E2=\"x\""), false));
  EXPECT_TRUE(IsBool(Eval("=E2=5"), false));
}

TEST_F(PictureValue, OrderingAgainstNumberIsValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=E2>5"), ErrorCode::Value));
}

TEST_F(PictureValue, NotEqualsNegatesEquals) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=E2<>E4"), false));
}

TEST_F(PictureValue, OrderingAgainstAnyOperandIsValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=E2<E6"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=E2>TRUE"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=E2>E3"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=E2>\"zzz\""), ErrorCode::Value));
}

TEST_F(PictureValue, EqualityAcrossAltAndOrigin) {
  BuildMain();
  PutFormula("G2", "=E2");
  EXPECT_TRUE(IsBool(Eval("=E6=E8"), false));
  EXPECT_TRUE(IsBool(Eval("=E2=G2"), true));
}

TEST_F(PictureValue, IdenticalBytesInSecondPartAreEqual) {
  BuildMain();
  Put("E10", Value::picture(AddPicture(image1_media_), kPlaced));
  EXPECT_TRUE(IsBool(Eval("=E2=E10"), true));
}

// --- Identity pairs: `=`, COUNTIF, MATCH, UNIQUE and XLOOKUP agree --------

TEST_F(PictureValue, IdentitySameBytesInAnotherEntry) {
  BuildIdentity();
  EXPECT_TRUE(IsBool(Eval("=E2=P2"), true));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(P2,E2)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=MATCH(E2,P2,0)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(UNIQUE(VSTACK(E2,P2)))"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=XLOOKUP(E2,P2,1,0)"), 1.0));
}

TEST_F(PictureValue, IdentityDifferentImages) {
  BuildIdentity();
  EXPECT_TRUE(IsBool(Eval("=E2=E6"), false));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(E6,E2)"), 0.0));
  EXPECT_TRUE(IsError(Eval("=MATCH(E2,E6,0)"), ErrorCode::NA));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(UNIQUE(VSTACK(E2,E6)))"), 2.0));
  EXPECT_TRUE(IsNumber(Eval("=XLOOKUP(E2,E6,1,0)"), 0.0));
}

TEST_F(PictureValue, IdentitySamePixelsInDifferentBytes) {
  BuildIdentity();
  EXPECT_TRUE(IsBool(Eval("=E2=P3"), false));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(P3,E2)"), 0.0));
  EXPECT_TRUE(IsError(Eval("=MATCH(E2,P3,0)"), ErrorCode::NA));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(UNIQUE(VSTACK(E2,P3)))"), 2.0));
  EXPECT_TRUE(IsNumber(Eval("=XLOOKUP(E2,P3,1,0)"), 0.0));
}

// Pinned: the alt folds case as the profile's text `=` does; measured on ASCII only.
TEST_F(PictureValue, AltCaseFoldsAtTextEquality) {
  BuildIdentity();
  EXPECT_TRUE(IsBool(Eval("=P4=P5"), true));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(P5,P4)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=MATCH(P4,P5,0)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(UNIQUE(VSTACK(P4,P5)))"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=XLOOKUP(P4,P5,1,0)"), 1.0));
}

TEST_F(PictureValue, IdentitySameAltInAnotherEntry) {
  BuildIdentity();
  EXPECT_TRUE(IsBool(Eval("=P4=P6"), true));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(P6,P4)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=MATCH(P4,P6,0)"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(UNIQUE(VSTACK(P4,P6)))"), 1.0));
  EXPECT_TRUE(IsNumber(Eval("=XLOOKUP(P4,P6,1,0)"), 1.0));
}

TEST_F(PictureValue, IdentityAltAgainstNoAlt) {
  BuildIdentity();
  EXPECT_TRUE(IsBool(Eval("=E2=P4"), false));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(P4,E2)"), 0.0));
  EXPECT_TRUE(IsError(Eval("=MATCH(E2,P4,0)"), ErrorCode::NA));
  EXPECT_TRUE(IsNumber(Eval("=ROWS(UNIQUE(VSTACK(E2,P4)))"), 2.0));
  EXPECT_TRUE(IsNumber(Eval("=XLOOKUP(E2,P4,1,0)"), 0.0));
}

// --- Arithmetic, aggregates and coercion ---------------------------------

TEST_F(PictureValue, ArithmeticIsValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=E2+1"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=-E2"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=SUM(E2,1)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=--E2"), ErrorCode::Value));
}

TEST_F(PictureValue, IsNumberOfArithmeticIsFalse) {
  BuildMain();
  EXPECT_TRUE(IsBool(Eval("=ISNUMBER(E2+0)"), false));
}

TEST_F(PictureValue, LiftedRangeArithmeticErrsPerPicture) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=E2:E4+1"), ErrorCode::Value));
  EXPECT_TRUE(IsNumber(Member(1U), 1.0));
  EXPECT_TRUE(IsError(Member(2U), ErrorCode::Value));
}

TEST_F(PictureValue, SumOverRangeIsValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=SUM(E2:E4)"), ErrorCode::Value));
}

TEST_F(PictureValue, AggregatesOverRange) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=MAX(E2:E4)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=AVERAGE(E2:E4)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=MAXA(E2:E4)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=AVERAGEA(E2:E4)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=PRODUCT(E2:E4)"), ErrorCode::Value));
  EXPECT_TRUE(IsNumber(Eval("=AGGREGATE(4,6,E2:E4)"), 0.0));
  EXPECT_TRUE(IsError(Eval("=SUBTOTAL(9,E2:E4)"), ErrorCode::Value));
}

TEST_F(PictureValue, LogicalOverPictureIsValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=IF(E2,1,2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=NOT(E2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=AND(E2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=AND(E2:E4,TRUE)"), ErrorCode::Value));
}

TEST_F(PictureValue, ValueMmultImrealAreValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=VALUE(E2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=MMULT(E2,1)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=IMREAL(E2)"), ErrorCode::Value));
}

TEST_F(PictureValue, TAndNAreValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=T(E2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=N(E2)"), ErrorCode::Value));
}

TEST_F(PictureValue, ExactIsCalcAndTextIsValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=EXACT(E2,E4)"), ErrorCode::Calc));
  EXPECT_TRUE(IsError(Eval("=TEXT(E2,\"0\")"), ErrorCode::Value));
}

TEST_F(PictureValue, ConcatAndTextjoinAreValueError) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=CONCAT(E2)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=TEXTJOIN(\",\",TRUE,E2:E4)"), ErrorCode::Value));
  EXPECT_TRUE(IsError(Eval("=CONCAT(E2:E4)"), ErrorCode::Value));
}

// --- Criteria and lookups ------------------------------------------------

TEST_F(PictureValue, CountifPictureCriterion) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(E2:E6,E2)"), 2.0));
}

TEST_F(PictureValue, CountifOperatorCriteria) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(E2:E6,\"<>\")"), 3.0));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(E2:E6,\"*\")"), 0.0));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(E2:E6,\">0\")"), 0.0));
  EXPECT_TRUE(IsNumber(Eval("=COUNTIF(E2:E6,\"#VALUE!\")"), 0.0));
}

TEST_F(PictureValue, MatchExactApproximateAndVlookup) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=MATCH(E2,E2:E6,0)"), 1.0));
  EXPECT_TRUE(IsError(Eval("=MATCH(5,E2:E6,1)"), ErrorCode::NA));
  EXPECT_TRUE(IsPicture(Eval("=VLOOKUP(E2,E2:E6,1,FALSE)"), image1_, kFormula));
}

// Pinned: an approximate lookup treats a picture as not comparable.
TEST_F(PictureValue, ApproximateLookupOfAPictureIsNA) {
  BuildMain();
  EXPECT_TRUE(IsError(Eval("=MATCH(E2,E2:E6,1)"), ErrorCode::NA));
  EXPECT_TRUE(IsError(Eval("=MATCH(E2,E2:E6,-1)"), ErrorCode::NA));
  EXPECT_TRUE(IsError(Eval("=XLOOKUP(E2,E2:E6,E2:E6,,1)"), ErrorCode::NA));
  EXPECT_TRUE(IsError(Eval("=XLOOKUP(E2,E2:E6,E2:E6,,-1)"), ErrorCode::NA));
}

// Estimate: SWITCH matches with the same identity as `=`.
TEST_F(PictureValue, SwitchUsesPictureIdentity) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=SWITCH(E2,E6,1,E4,2,3)"), 2.0));
  EXPECT_TRUE(IsNumber(Eval("=SWITCH(E8,E6,1,9)"), 9.0));
}

// --- Carried values ------------------------------------------------------

TEST_F(PictureValue, CarriedThroughXlookupIfAndIndex) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=XLOOKUP(1,{1},E2)"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=IF(TRUE,E2)"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=INDEX(E2:E6,3)"), image1_, kFormula));
  EXPECT_TRUE(IsNumber(Eval("=TYPE(XLOOKUP(1,{1},E2))"), 128.0));
}

TEST_F(PictureValue, CarriedThroughVstackChooseAndLet) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=INDEX(VSTACK(E2,E6),2)"), image2_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=CHOOSE(2,1,E2)"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=LET(x,E2,x)"), image1_, kFormula));
}

TEST_F(PictureValue, CellContentsIsThePicture) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=CELL(\"contents\",E2)"), image1_, kFormula));
}

TEST_F(PictureValue, SpillOfRangeHoldingPictures) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=E2:E4"), image1_, kFormula));
  EXPECT_TRUE(IsNumber(Member(1U), 0.0));
  EXPECT_TRUE(IsPicture(Member(2U), image1_, kFormula));
}

// --- Sorting, UNIQUE and GROUPBY -----------------------------------------

TEST_F(PictureValue, SortAcrossKindsInARange) {
  BuildMain();
  Put("E3", Value::number(7.0));
  Put("E5", Value::text("x"));
  EXPECT_TRUE(IsNumber(Eval("=SORT(E2:E6)"), 7.0));
  EXPECT_TRUE(IsText(Member(1U), "x"));
  EXPECT_TRUE(IsPicture(Member(2U), image2_, kFormula));
  EXPECT_TRUE(IsPicture(Member(3U), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Member(4U), image1_, kFormula));
}

TEST_F(PictureValue, SortAcrossKindsInAnArray) {
  BuildMain();
  EXPECT_TRUE(IsNumber(Eval("=SORT(VSTACK(E2,TRUE,\"x\",7))"), 7.0));
  EXPECT_TRUE(IsText(Member(1U), "x"));
  EXPECT_TRUE(IsPicture(Member(2U), image1_, kFormula));
  EXPECT_TRUE(IsBool(Member(3U), true));
  EXPECT_TRUE(IsNumber(Eval("=SORT(VSTACK(E2,7))"), 7.0));
  EXPECT_TRUE(IsPicture(Member(1U), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=SORT(VSTACK(E2,TRUE))"), image1_, kFormula));
  EXPECT_TRUE(IsBool(Member(1U), true));
}

TEST_F(PictureValue, SortAmongPicturesByMediaBytes) {
  BuildSort();
  EXPECT_TRUE(IsPicture(Eval("=SORT(Q2:Q4)"), sk_b_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), sk_c_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), sk_a_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=SORT(Q2:Q4,,-1)"), sk_a_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), sk_c_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), sk_b_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=SORTBY(Q2:Q4,Q2:Q4)"), sk_b_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), sk_c_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), sk_a_, kFormula));
}

// Pinned: two present alts order at the text sort collation; measured on `a` / `b` only.
TEST_F(PictureValue, SortAltByCollation) {
  BuildSort();
  EXPECT_TRUE(IsPicture(Eval("=SORT(Q5:Q6)"), sk_a_alt_a_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), sk_a_alt_b_, kFormula));
}

// Pinned: an error sorts after a picture, as after a logical.
TEST_F(PictureValue, SortPutsErrorAfterPicture) {
  BuildMain();
  Put("E3", Value::error(ErrorCode::NA));
  EXPECT_TRUE(IsPicture(Eval("=SORT(E2:E3)"), image1_, kFormula));
  EXPECT_TRUE(IsError(Member(1U), ErrorCode::NA));
}

// Pinned: media bytes order before alt presence (image2 without an alt before image1 with one).
TEST_F(PictureValue, SortBytesBeforeAltPresence) {
  BuildIdentity();
  EXPECT_TRUE(IsPicture(Eval("=SORT(VSTACK(P4,E6))"), image2_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), p4_, kFormula));
}

// Pinned: a blank sorts last, and descending is the exact reverse across the alt tiers.
TEST_F(PictureValue, SortBlankAndDescending) {
  BuildSort();
  EXPECT_TRUE(IsPicture(Eval("=SORT(Q2:Q7)"), sk_b_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), sk_c_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), sk_a_alt_a_, kFormula));
  EXPECT_TRUE(IsPicture(Member(3U), sk_a_alt_b_, kFormula));
  EXPECT_TRUE(IsPicture(Member(4U), sk_a_, kFormula));
  EXPECT_TRUE(IsNumber(Member(5U), 0.0));
  EXPECT_TRUE(IsPicture(Eval("=SORT(Q2:Q6,,-1)"), sk_a_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), sk_a_alt_b_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), sk_a_alt_a_, kFormula));
  EXPECT_TRUE(IsPicture(Member(3U), sk_c_, kFormula));
  EXPECT_TRUE(IsPicture(Member(4U), sk_b_, kFormula));
}

TEST_F(PictureValue, UniqueKeepsFirstOfEachClass) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=UNIQUE(E2:E6)"), image1_, kFormula));
  EXPECT_TRUE(IsNumber(Member(1U), 0.0));
  EXPECT_TRUE(IsPicture(Member(2U), image2_, kFormula));
}

TEST_F(PictureValue, GroupbyKeysOnePerClassInSortOrder) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=GROUPBY(E2:E6,{1;2;3;4;5},SUM)"), image2_, kFormula));
  EXPECT_TRUE(IsNumber(Member(0U, 1U), 5.0));
  EXPECT_TRUE(IsPicture(Member(1U), image1_, kFormula));
  EXPECT_TRUE(IsNumber(Member(1U, 1U), 4.0));
  EXPECT_TRUE(IsNumber(Member(2U), 0.0));
  EXPECT_TRUE(IsNumber(Member(2U, 1U), 6.0));
  EXPECT_TRUE(IsText(Member(3U), "合計"));
  EXPECT_TRUE(IsNumber(Member(3U, 1U), 15.0));
}

TEST_F(PictureValue, GroupbyKeysSplitOnAlt) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=GROUPBY(E2:E8,SEQUENCE(7),SUM)"), image2_alt_, kFormula));
  EXPECT_TRUE(IsNumber(Member(0U, 1U), 7.0));
  EXPECT_TRUE(IsPicture(Member(1U), image2_, kFormula));
  EXPECT_TRUE(IsNumber(Member(1U, 1U), 5.0));
  EXPECT_TRUE(IsPicture(Member(2U), image1_, kFormula));
  EXPECT_TRUE(IsNumber(Member(2U, 1U), 4.0));
  EXPECT_TRUE(IsNumber(Member(3U), 0.0));
  EXPECT_TRUE(IsNumber(Member(3U, 1U), 12.0));
  EXPECT_TRUE(IsText(Member(4U), "合計"));
  EXPECT_TRUE(IsNumber(Member(4U, 1U), 28.0));
}

// --- Formula-result origin -----------------------------------------------

TEST_F(PictureValue, OriginOfReferenceRoot) {
  BuildMain();
  PutFormula("G2", "=E2");
  PutFormula("G7", "=E8");
  ASSERT_TRUE(static_cast<bool>(wb_.recalc(default_registry())));
  EXPECT_TRUE(IsPicture(wb_.sheet(0).resolve_cell_value(1U, 6U), image1_, kFormula));
  EXPECT_TRUE(IsPicture(wb_.sheet(0).resolve_cell_value(6U, 6U), image2_alt_, kFormula));
}

TEST_F(PictureValue, OriginOfLookupAndSelectionCalls) {
  BuildMain();
  EXPECT_EQ(Eval("=XLOOKUP(1,{1},E2)").picture_origin(), kFormula);
  EXPECT_EQ(Eval("=VLOOKUP(E2,E2:E6,1,FALSE)").picture_origin(), kFormula);
  EXPECT_EQ(Eval("=CHOOSE(2,1,E2)").picture_origin(), kFormula);
  EXPECT_EQ(Eval("=LET(x,E2,x)").picture_origin(), kFormula);
  EXPECT_EQ(Eval("=INDEX(VSTACK(E2,E6),2)").picture_origin(), kFormula);
  EXPECT_EQ(Eval("=CELL(\"contents\",E2)").picture_origin(), kFormula);
}

TEST_F(PictureValue, OriginOfIfKeepsTheAlt) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=IF(TRUE,E2)"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=IF(TRUE,E8)"), image2_alt_, kFormula));
}

TEST_F(PictureValue, OriginOfReferenceIndex) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=INDEX(E2:E6,3)"), image1_, kFormula));
}

TEST_F(PictureValue, OriginOfRangeSpillMembers) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=E2:E6"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Member(4U), image2_, kFormula));
}

TEST_F(PictureValue, OriginOfSortAndUniqueMembers) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=SORT(E2:E6)"), image2_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=UNIQUE(E2:E6)"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Member(2U), image2_, kFormula));
}

TEST_F(PictureValue, OriginOfSortOverVstackAndGroupbyKeys) {
  BuildMain();
  EXPECT_TRUE(IsPicture(Eval("=SORT(VSTACK(E2,TRUE))"), image1_, kFormula));
  EXPECT_TRUE(IsPicture(Eval("=GROUPBY(E2:E6,{1;2;3;4;5},SUM)"), image2_, kFormula));
  EXPECT_TRUE(IsPicture(Member(1U), image1_, kFormula));
}

TEST_F(PictureValue, StampCopiesAnArrayHoldingAPicture) {
  Arena parse_arena;
  Arena arena;
  parser::Parser p("=E2:E4", parse_arena);
  const parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const PicturePayload pic{7U, 3U};
  const Value cells[3] = {Value::picture(pic, kPlaced), Value::number(1.0), Value::picture(pic, kPlaced)};
  const ArrayValue array{3U, 1U, cells};
  const Value stamped = stamp_picture_origins(*root, Value::array(&array), arena);
  ASSERT_TRUE(stamped.is_array());
  EXPECT_NE(stamped.as_array(), &array);
  const Value* out = stamped.as_array_cells();
  EXPECT_TRUE(IsPicture(out[0], pic, kFormula));
  EXPECT_TRUE(IsNumber(out[1], 1.0));
  EXPECT_TRUE(IsPicture(out[2], pic, kFormula));
  EXPECT_TRUE(IsPicture(cells[0], pic, kPlaced));
}

TEST_F(PictureValue, StampLeavesOtherValuesUntouched) {
  Arena parse_arena;
  Arena arena;
  parser::Parser p("=A1:A2", parse_arena);
  const parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const Value cells[2] = {Value::number(1.0), Value::boolean(true)};
  const ArrayValue array{2U, 1U, cells};
  EXPECT_EQ(stamp_picture_origins(*root, Value::array(&array), arena).as_array(), &array);
  EXPECT_TRUE(IsNumber(stamp_picture_origins(*root, Value::number(2.0), arena), 2.0));
  EXPECT_TRUE(IsPicture(stamp_picture_origins(*root, Value::picture(PicturePayload{1U, 0U}, kPlaced), arena),
                        PicturePayload{1U, 0U}, kFormula));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
