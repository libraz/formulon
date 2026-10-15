// Tests for NUMBERVALUE's separator pre-pass and the VALUE-stage grammar it
// feeds: grouping, signs, currency, digit scripts.

#include <cfloat>
#include <string>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

using formulon::test::EvalSourceIn;

class NumberValueRules : public ::testing::Test {
 protected:
  void SetUp() override {
    workbook_ = Workbook::create();
    workbook_.set_excel_profile(mac_365_ja_jp_profile());
  }

  Value eval(const std::string& formula) { return EvalSourceIn(formula, workbook_, workbook_.sheet(0)); }

  void expect_number(const std::string& formula, double want) {
    const Value v = eval(formula);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), want) << formula;
  }

  void expect_value_error(const std::string& formula) {
    const Value v = eval(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }

  Workbook workbook_ = Workbook::create();
};

TEST_F(NumberValueRules, WhitespaceAndLength) {
  expect_number("=NUMBERVALUE(\"1 2\" & CHAR(9) & \"3\")", 123.0);
  expect_number("=NUMBERVALUE(\"1 . 5\")", 1.5);
  expect_value_error("=NUMBERVALUE(\"1　　2\")");
  expect_number("=NUMBERVALUE(\"1\" & REPT(\" \", 253) & \"1\")", 11.0);
  expect_value_error("=NUMBERVALUE(\"1\" & REPT(\" \", 254) & \"1\")");
}

TEST_F(NumberValueRules, SeparatorsComeFromTheFirstUnit) {
  expect_number("=NUMBERVALUE(\"1,23\")", 123.0);
  expect_value_error("=NUMBERVALUE(\"1.5,000\")");
  expect_number("=NUMBERVALUE(\"1.5\", \",\")", 15.0);
  expect_number("=NUMBERVALUE(\"1.234.567,5\", \",\")", 1234567.5);
  expect_value_error("=NUMBERVALUE(\"1,234,5\", \",\")");
  expect_number("=NUMBERVALUE(\"2,5\", 1)", 25.0);
  expect_number("=NUMBERVALUE(\"1T5\", TRUE)", 1.5);
  expect_value_error("=NUMBERVALUE(\"15\", \".\", \".\")");
  expect_value_error("=NUMBERVALUE(\"15\", \"\")");
  expect_number("=NUMBERVALUE(\"105\", \"0\")", 1.5);
  expect_number("=NUMBERVALUE(\"15%\", \"%\")", 15.0);
  expect_number("=NUMBERVALUE(\"1,500\", \".\", \";\")", 1500.0);
  expect_number("=NUMBERVALUE(\"1.500\", \"x\", \";\")", 1.5);
}

TEST_F(NumberValueRules, TrailingPercentDividesAfterTheRangeCheck) {
  expect_number("=NUMBERVALUE(\"7%%\")", 7.0 / 100.0 / 100.0);
  expect_number("=NUMBERVALUE(\"%\")", 0.0);
  expect_value_error("=NUMBERVALUE(\"1e309%\")");
  expect_number("=NUMBERVALUE(\"1e-307%\")", 0.0);
  expect_number("=NUMBERVALUE(\"2.3e-306%\")", 2.3e-306 / 100.0);
}

TEST_F(NumberValueRules, NonTextFirstArgument) {
  expect_value_error("=NUMBERVALUE(TRUE)");
  expect_number("=NUMBERVALUE(1/3)", 0.333333333333333);
  expect_number("=NUMBERVALUE(\"\")", 0.0);
}

TEST_F(NumberValueRules, DateFallbackSeesThePrePassedText) {
  expect_number("=NUMBERVALUE(\"1 : 30\")", 1.5 / 24.0);
  expect_number("=NUMBERVALUE(\"1:30\", \":\")", 1.3);
}

TEST_F(NumberValueRules, ValueStageGrouping) {
  expect_number("=VALUE(\"1,2345\")", 12345.0);
  expect_number("=VALUE(\"1,234,5678\")", 12345678.0);
  expect_number("=VALUE(\"12345,678\")", 12345678.0);
  expect_value_error("=VALUE(\"0,123\")");
  expect_value_error("=VALUE(\"1,23\")");
  expect_value_error("=VALUE(\"1e1,0\")");
  expect_number("=VALUE(\"１，２３４５\")", 12345.0);
  expect_value_error("=VALUE(\"１，５\")");
}

TEST_F(NumberValueRules, ValueStageSignCurrencyParens) {
  expect_number("=VALUE(\"- 5\")", -5.0);
  expect_number("=VALUE(\"-　5\")", -5.0);
  expect_value_error("=VALUE(\"-\" & CHAR(9) & \"5\")");
  expect_number("=VALUE(\"¥(1)\")", -1.0);
  expect_number("=VALUE(\"€(5)\")", -5.0);
  expect_number("=VALUE(\"(5)€\")", -5.0);
  expect_number("=VALUE(\"(5€)\")", -5.0);
  expect_value_error("=VALUE(\"¥1%\")");
  expect_value_error("=VALUE(\"(¥1)%\")");
  expect_value_error("=VALUE(\"5€%\")");
  expect_number("=VALUE(\"＄５\")", 5.0);
  expect_value_error("=VALUE(\"1　2\")");
}

TEST_F(NumberValueRules, ValueStageDigitScripts) {
  expect_number("=VALUE(\"١٢٣\")", 123.0);
  expect_number("=VALUE(\"१.५\")", 1.5);
  expect_number("=VALUE(\"௧௨\")", 12.0);
  expect_number("=VALUE(\"፩፪\")", 12.0);
  expect_value_error("=VALUE(\"①\")");
  expect_number("=VALUE(\"١:٣٠\")", 1.5 / 24.0);
}

}  // namespace
}  // namespace formulon::eval
