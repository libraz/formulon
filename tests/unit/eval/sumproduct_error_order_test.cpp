// SUMPRODUCT reports the first error a row-major walk of each argument
// reaches. Over a whole-row argument the repeated tail column of row r comes
// before the head cells of row r + 1.
//
// Sheet S: A1 = 1, A2 = "x". `1/S!1:2` holds 1 in A1, #VALUE! in A2 and
// #DIV/0! in every blank tail cell, so B1 (#DIV/0!) precedes A2 (#VALUE!).

#include <cstdint>
#include <string>

#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint32_t kH = 0U;  // Sheet index of H, where formulas are anchored.
constexpr std::uint32_t kS = 1U;  // Sheet index of S.

class SumproductErrorOrder : public ::testing::Test {
 protected:
  SumproductErrorOrder() : wb_(Workbook::create()) {
    EXPECT_TRUE(static_cast<bool>(wb_.rename_sheet(kH, "H")));
    wb_.add_sheet("S");
    EXPECT_TRUE(static_cast<bool>(wb_.set_cell_value(kS, 0U, 0U, Value::number(1.0))));
    EXPECT_TRUE(static_cast<bool>(wb_.set_cell_text(kS, 1U, 0U, "x")));
  }

  void ExpectError(const std::string& formula, ErrorCode code) {
    Arena arena;
    const Value v = evaluate_formula_text_array(wb_, wb_.sheet(kH), 0U, 7U, formula, arena, default_registry());
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), code) << formula;
  }

  Workbook wb_;
};

TEST_F(SumproductErrorOrder, WholeRowTailOfEarlierRowPrecedesLaterHead) {
  ExpectError("SUMPRODUCT(1/S!1:2)", ErrorCode::Div0);
}

TEST_F(SumproductErrorOrder, BoundedRowsWalkRowMajor) {
  ExpectError("SUMPRODUCT(1/S!A1:B2)", ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
