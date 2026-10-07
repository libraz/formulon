//
// ROWS / COLUMNS and SUMPRODUCT over a whole column or row read the declared
// size of the array without expanding it.
//
// Sheet L: A1:A5 = "k1","k2","k3",(blank),"k5"; B1:B5 = 10..50; column C is
// empty. Formulas are evaluated as if on sheet H. Expected values are the Mac
// Excel 365 ja-JP measurements recorded in
// tests/oracle/cases/whole_axis_spill.yaml.

#include <cstddef>
#include <cstdint>
#include <string>

#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint32_t kH = 0U;  // Sheet index of H.
constexpr std::uint32_t kL = 1U;  // Sheet index of L.
constexpr std::uint32_t kHCol = 7U;

// A dense whole column cannot come in under this; the evaluations measured
// here stay far below it.
constexpr std::size_t kArenaCeilingBytes = 1U << 20;

class WholeAxisReduce : public ::testing::Test {
 protected:
  WholeAxisReduce() : wb_(Workbook::create()) {
    EXPECT_TRUE(static_cast<bool>(wb_.rename_sheet(0U, "H")));
    wb_.add_sheet("L");
    const char* const keys[] = {"k1", "k2", "k3", nullptr, "k5"};
    for (std::uint32_t r = 0; r < 5U; ++r) {
      if (keys[r] != nullptr) {
        EXPECT_TRUE(static_cast<bool>(wb_.set_cell_text(kL, r, 0U, keys[r])));
      }
      EXPECT_TRUE(static_cast<bool>(wb_.set_cell_value(kL, r, 1U, Value::number(10.0 * (r + 1U)))));
    }
  }

  // Read-only evaluation of `formula` anchored at H1; `bytes_` records the
  // arena the evaluation used.
  Value Adhoc(const std::string& formula) {
    Arena arena;
    const Value v = evaluate_formula_text_array(wb_, wb_.sheet(kH), 0U, kHCol, formula, arena, default_registry());
    bytes_ = arena.bytes_used();
    return v;
  }

  double Number(const std::string& formula) {
    const Value v = Adhoc(formula);
    EXPECT_TRUE(v.is_number()) << formula;
    return v.is_number() ? v.as_number() : 0.0;
  }

  void ExpectError(const std::string& formula, ErrorCode code) {
    const Value v = Adhoc(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), code) << formula;
  }

  Workbook wb_;
  std::size_t bytes_ = 0;
};

TEST_F(WholeAxisReduce, RowsOfDerivedColumnIsDeclaredHeight) {
  EXPECT_EQ(Number("ROWS(L!B:B+0)"), 1048576.0);
  EXPECT_LT(bytes_, kArenaCeilingBytes);
}

TEST_F(WholeAxisReduce, RowsOfLetBoundColumnIsDeclaredHeight) {
  EXPECT_EQ(Number("LET(x,L!B:B,ROWS(x))"), 1048576.0);
  EXPECT_LT(bytes_, kArenaCeilingBytes);
}

TEST_F(WholeAxisReduce, RowsOfScalarFunctionOverColumn) {
  EXPECT_EQ(Number("ROWS(ABS(L!B:B))"), 1048576.0);
}

TEST_F(WholeAxisReduce, RowsOfWideDerivedSpanStaysUnderTheCeiling) {
  EXPECT_EQ(Number("ROWS(L!A:J+0)"), 1048576.0);
  EXPECT_EQ(Number("COLUMNS(L!A:J+0)"), 10.0);
  EXPECT_LT(bytes_, kArenaCeilingBytes);
}

TEST_F(WholeAxisReduce, ColumnsOfDerivedRowIsDeclaredWidth) {
  EXPECT_EQ(Number("COLUMNS(L!1:1+0)"), 16384.0);
  EXPECT_EQ(Number("ROWS(L!1:1+0)"), 1.0);
}

TEST_F(WholeAxisReduce, SumproductOfMaskedColumn) {
  EXPECT_EQ(Number("SUMPRODUCT((L!A:A=\"k2\")*L!B:B)"), 20.0);
  EXPECT_LT(bytes_, kArenaCeilingBytes);
}

TEST_F(WholeAxisReduce, SumproductCountsBlankTail) {
  EXPECT_EQ(Number("SUMPRODUCT((L!A:A=\"\")*1)"), 1048572.0);
  EXPECT_LT(bytes_, kArenaCeilingBytes);
}

TEST_F(WholeAxisReduce, SumproductOfTwoColumns) {
  EXPECT_EQ(Number("SUMPRODUCT(L!B:B,L!B:B)"), 5500.0);
}

TEST_F(WholeAxisReduce, SumproductColumnAgainstShortRangeIsValueError) {
  ExpectError("SUMPRODUCT(L!A:A,L!B1:B5)", ErrorCode::Value);
}

TEST_F(WholeAxisReduce, SumproductTailIsSummedRowByRow) {
  // 0.1 added 1048572 times in row order; a product x count would differ in
  // the last bits.
  EXPECT_EQ(Number("SUMPRODUCT((L!A:A=\"\")*0.1)-104857.2"), 1.6156118363142014e-06);
}

TEST_F(WholeAxisReduce, SumproductTailIsScannedForErrors) {
  // Every head cell is a number; only the repeated tail divides by zero.
  ExpectError("SUMPRODUCT(1/(L!B:B*1))", ErrorCode::Div0);
}

TEST_F(WholeAxisReduce, SumproductOfWholeRow) {
  EXPECT_EQ(Number("SUMPRODUCT((L!1:1=\"\")*1)"), 16382.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
