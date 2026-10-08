//
// Values Excel 365 gives a built-in function named as a value (`=TYPE(SUM)`)
// and the trim-reference operators `.:`, `:.`, `.:.`, each formula entered
// into a workbook and recalculated the way a cell is.

#include <cstdint>
#include <string>
#include <vector>

#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

class FunctionValueTrimRefEval : public ::testing::Test {
 protected:
  void SetUp() override {
    wb_.add_sheet("Sheet1");
    wb_.add_sheet("Sheet2");
  }

  // Recalculates `formula` in Sheet1!H20, away from every setup cell.
  Value Eval(const std::string& formula) {
    EXPECT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, 19, 7, formula))) << formula;
    EXPECT_TRUE(static_cast<bool>(wb_.recalc(eval::default_registry()))) << formula;
    const Cell* cell = wb_.sheet(0).cell_at(19, 7);
    return cell != nullptr ? cell->cached_value : Value::blank();
  }

  void ExpectNumber(const std::string& formula, double expected) {
    const Value v = Eval(formula);
    ASSERT_TRUE(v.is_number()) << formula << " -> " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), expected) << formula;
  }

  void ExpectError(const std::string& formula, ErrorCode expected) {
    const Value v = Eval(formula);
    ASSERT_TRUE(v.is_error()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), expected) << formula;
  }

  void ExpectBool(const std::string& formula, bool expected) {
    const Value v = Eval(formula);
    ASSERT_TRUE(v.is_boolean()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_boolean(), expected) << formula;
  }

  void Set(std::size_t sheet, std::uint32_t row, std::uint32_t col, double v) {
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_value(sheet, row, col, Value::number(v))));
  }

  Workbook wb_ = Workbook::create_empty();
};

TEST_F(FunctionValueTrimRefEval, BareFunctionNameIsAFunctionValue) {
  ExpectError("=SUM", ErrorCode::Calc);
  ExpectError("=sum", ErrorCode::Calc);
  ExpectError("=UNICODE", ErrorCode::Calc);
  ExpectError("=IF(TRUE,SUM)", ErrorCode::Calc);
  ExpectError("=ABS+1", ErrorCode::Value);
  ExpectError("=SUM(1,2)+SUM", ErrorCode::Value);
  ExpectError("=SUM({1,2})&ABS", ErrorCode::Value);
  ExpectError("=ABS(SUM)", ErrorCode::Value);
  ExpectNumber("=TYPE(SUM)", 128.0);
  ExpectNumber("=TYPE(sum)", 128.0);
  ExpectNumber("=TYPE(XLOOKUP)", 128.0);
  ExpectNumber("=TYPE(LAMBDA(x,x))", 128.0);
  ExpectBool("=ISERROR(SUM)", false);
  ExpectBool("=ISERROR(LAMBDA(x,x))", false);
  ExpectBool("=ISREF(SUM)", false);
  ExpectBool("=ISOMITTED(SUM)", false);
  ExpectError("=ERROR.TYPE(SUM)", ErrorCode::NA);
  ExpectError("=INDEX(ABS,1)", ErrorCode::Calc);
  ExpectError("=FOOBARBAZ", ErrorCode::Name);
}

TEST_F(FunctionValueTrimRefEval, FunctionValueIsCallableWhereverItFlows) {
  ExpectNumber("=LET(f,ABS,f(-2))", 2.0);
  ExpectNumber("=CHOOSE(1,SUM,ABS)(5)", 5.0);
  ExpectNumber("=IF(TRUE,ABS,SUM)(-3)", 3.0);
  ExpectNumber("=REDUCE(0,{1,2},SUM)", 3.0);
  ExpectNumber("=LET(f,MAX,f(4,9,2))", 9.0);
  const Value mapped = Eval("=SUM(MAP({-1,-2},ABS))");
  ASSERT_TRUE(mapped.is_number()) << mapped.debug_to_string();
  EXPECT_DOUBLE_EQ(mapped.as_number(), 3.0);
}

TEST_F(FunctionValueTrimRefEval, BindingsAndDefinedNamesOfTheSameSpellingWin) {
  ExpectNumber("=LET(SUM,5,TYPE(SUM))", 1.0);
  wb_.set_defined_names({DefinedName{"ABS", "7", -1, false, ""}});
  ExpectNumber("=ABS", 7.0);
  ExpectNumber("=TYPE(ABS)", 1.0);
}

TEST_F(FunctionValueTrimRefEval, TrimReferencesOverBoundedRanges) {
  Set(0, 0, 0, 1.0);
  Set(0, 1, 0, 2.0);
  Set(0, 2, 0, 3.0);
  ExpectNumber("=SUM(A1:.A10)", 6.0);
  ExpectNumber("=ROWS(A1:.A10)", 3.0);
  ExpectNumber("=ROWS(TRIMRANGE(A1:A10))", 3.0);
  ExpectNumber("=ROWS(A:.A)", 3.0);
  ExpectNumber("=ROWS($A$1:.$A$10)", 3.0);
  ExpectNumber("=ROWS(Sheet1!A1:.A10)", 3.0);
}

TEST_F(FunctionValueTrimRefEval, LeadingAndBothEdges) {
  Set(0, 1, 0, 1.0);
  Set(0, 2, 0, 2.0);
  ExpectNumber("=ROWS(A1.:.A10)", 2.0);
  ExpectNumber("=ROWS(A1.:A10)", 9.0);
  ExpectNumber("=SUM(A1.:.A10)", 3.0);
  ExpectNumber("=ROWS(A.:A)", 1048575.0);
  ExpectNumber("=ROWS(A.:.A)", 2.0);
  ExpectNumber("=ROWS(TRIMRANGE(A1:A10,2))", 3.0);
  ExpectNumber("=ROWS(TRIMRANGE(A1:A10,0))", 10.0);
}

TEST_F(FunctionValueTrimRefEval, InteriorGapsAndEmptyRanges) {
  ExpectError("=ROWS(A1:.A10)", ErrorCode::Ref);
  ExpectError("=TRIMRANGE(A1:A10)", ErrorCode::Ref);
  ExpectError("=ROWS(A:.A)", ErrorCode::Ref);
  Set(0, 0, 0, 1.0);
  Set(0, 4, 0, 2.0);
  ExpectNumber("=ROWS(A1:.A10)", 5.0);
}

TEST_F(FunctionValueTrimRefEval, TwoDimensionalAndWholeAxis) {
  Set(0, 0, 0, 1.0);
  Set(0, 1, 1, 2.0);
  const Value shape = Eval("=ROWS(A1:.C10)&\"x\"&COLUMNS(A1:.C10)");
  ASSERT_TRUE(shape.is_text()) << shape.debug_to_string();
  EXPECT_EQ(shape.as_text(), "2x2");
  Set(0, 0, 2, 5.0);
  ExpectNumber("=COLUMNS(1:.1)", 3.0);
  Set(0, 3, 1, 4.0);
  ExpectNumber("=ROWS(A:.B)", 4.0);
  ExpectNumber("=COLUMNS(A1:.E1)", 3.0);
}

TEST_F(FunctionValueTrimRefEval, FormulaInsideItsOwnTrimmedRangeIsNoCircularRead) {
  // Measured with the formula in Z1: row 1's last used column is Z itself.
  Set(0, 0, 0, 1.0);
  Set(0, 0, 2, 2.0);
  const auto eval_z1 = [&](const std::string& formula) {
    EXPECT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, 0, 25, formula))) << formula;
    EXPECT_TRUE(static_cast<bool>(wb_.recalc(eval::default_registry()))) << formula;
    return wb_.sheet(0).cell_at(0, 25)->cached_value;
  };
  const Value columns = eval_z1("=COLUMNS(1:.1)");
  ASSERT_TRUE(columns.is_number()) << columns.debug_to_string();
  EXPECT_DOUBLE_EQ(columns.as_number(), 26.0);
  const Value rows = eval_z1("=ROWS(A1.:.Z3)");
  ASSERT_TRUE(rows.is_number()) << rows.debug_to_string();
  EXPECT_DOUBLE_EQ(rows.as_number(), 1.0);
  // Reading the trimmed cells' values does read the formula's own cell, so
  // it follows the cycle policy of the same range written out.
  const Value through_trim = eval_z1("=SUM(1:.1)");
  const Value written_out = eval_z1("=SUM(A1:Z1)");
  EXPECT_EQ(through_trim.debug_to_string(), written_out.debug_to_string());
}

TEST_F(FunctionValueTrimRefEval, TrimRangeOverAReferenceIsAReference) {
  Set(0, 1, 0, 1.0);
  Set(0, 2, 0, 2.0);
  ExpectBool("=ISREF(A1:.A10)", true);
  ExpectBool("=ISREF(TRIMRANGE(A1:A10))", true);
  ExpectNumber("=ROW(A1.:A10)", 2.0);
  ExpectNumber("=ROWS(TRIMRANGE({1;2;3}))", 3.0);
}

TEST_F(FunctionValueTrimRefEval, FormulaReturningEmptyTextIsNotBlank) {
  Set(0, 0, 0, 1.0);
  ASSERT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, 1, 0, "=\"\"")));
  ExpectNumber("=ROWS(A1:.A10)", 2.0);
}

TEST_F(FunctionValueTrimRefEval, SheetQualifiedTrimReference) {
  Set(1, 0, 0, 1.0);
  Set(1, 1, 0, 2.0);
  ExpectNumber("=ROWS(Sheet2!A1:.A10)", 2.0);
}

TEST_F(FunctionValueTrimRefEval, TrimReferenceFollowsEditsToItsRange) {
  Set(0, 0, 0, 1.0);
  ExpectNumber("=ROWS(A1:.A10)", 1.0);
  Set(0, 5, 0, 2.0);
  ASSERT_TRUE(static_cast<bool>(wb_.recalc(eval::default_registry())));
  const Value v = wb_.sheet(0).cell_at(19, 7)->cached_value;
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

}  // namespace
}  // namespace formulon
