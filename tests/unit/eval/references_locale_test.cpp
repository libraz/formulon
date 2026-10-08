// ADDRESS, INDIRECT(a1=FALSE) and CELL follow the locale's R1C1 letters and CELL codes.

#include <string>
#include <string_view>

#include "builtins_references_test_helpers.h"
#include "eval/a1_parse.h"
#include "eval/eval_profile_scope.h"
#include "excel_profile.h"

namespace formulon {
namespace eval {
namespace {
using namespace references_test_helpers;

ExcelProfile mac(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

Value eval_as(ExcelLocale locale, std::string_view src, const Workbook& wb) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx = test::workbook_context(wb, wb.sheet(0), state).with_excel_profile(mac(locale));
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

std::string text_as(ExcelLocale locale, std::string_view src, const Workbook& wb) {
  const Value v = eval_as(locale, src, wb);
  EXPECT_TRUE(v.is_text()) << src;
  return v.is_text() ? std::string(v.as_text()) : std::string();
}

TEST(ReferencesLocale, AddressR1C1Letters) {
  const Workbook wb = Workbook::create();
  EXPECT_EQ(text_as(ExcelLocale::kDeDE, "=ADDRESS(2,3,1,FALSE)", wb), "Z2S3");
  EXPECT_EQ(text_as(ExcelLocale::kDeDE, "=ADDRESS(2,3,4,FALSE)", wb), "Z(2)S(3)");
  EXPECT_EQ(text_as(ExcelLocale::kFrFR, "=ADDRESS(2,3,1,FALSE)", wb), "L2C3");
  EXPECT_EQ(text_as(ExcelLocale::kFrFR, "=ADDRESS(2,3,4,FALSE)", wb), "L(2)C(3)");
  EXPECT_EQ(text_as(ExcelLocale::kJaJP, "=ADDRESS(2,3,4,FALSE)", wb), "R[2]C[3]");
}

TEST(ReferencesLocale, IndirectAcceptsOnlyLocaleLetters) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 2, Value::number(7.0));  // C2
  const auto number_at = [&](ExcelLocale locale, std::string_view src) {
    const Value v = eval_as(locale, src, wb);
    return v.is_number() ? v.as_number() : -1.0;
  };
  EXPECT_DOUBLE_EQ(number_at(ExcelLocale::kDeDE, "=INDIRECT(\"Z2S3\",FALSE)"), 7.0);
  EXPECT_TRUE(eval_as(ExcelLocale::kDeDE, "=INDIRECT(\"R2C3\",FALSE)", wb).is_error());
  EXPECT_TRUE(eval_as(ExcelLocale::kDeDE, "=INDIRECT(\"L2C3\",FALSE)", wb).is_error());
  EXPECT_DOUBLE_EQ(number_at(ExcelLocale::kFrFR, "=INDIRECT(\"L2C3\",FALSE)"), 7.0);
  EXPECT_TRUE(eval_as(ExcelLocale::kFrFR, "=INDIRECT(\"R2C3\",FALSE)", wb).is_error());
  EXPECT_DOUBLE_EQ(number_at(ExcelLocale::kJaJP, "=INDIRECT(\"R2C3\",FALSE)"), 7.0);
}

TEST(ReferencesLocale, ParserUsesLocaleBrackets) {
  const ScopedEvalProfile scope(mac(ExcelLocale::kDeDE));
  const refs_internal::R1C1Base base{true, 5, 5};
  const refs_internal::A1Parse ok = refs_internal::parse_r1c1_ref("Z(-1)S(1)", base);
  ASSERT_TRUE(ok.valid);
  EXPECT_EQ(ok.row, 4U);
  EXPECT_EQ(ok.col, 6U);
  EXPECT_FALSE(refs_internal::parse_r1c1_ref("Z[-1]S[1]", base).valid);
}

TEST(ReferencesLocale, CellTypeAndFormatCodes) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  EXPECT_EQ(text_as(ExcelLocale::kDeDE, "=CELL(\"type\",A1)", wb), "w");
  EXPECT_EQ(text_as(ExcelLocale::kDeDE, "=CELL(\"type\",A2)", wb), "l");
  EXPECT_EQ(text_as(ExcelLocale::kDeDE, "=CELL(\"type\",A3)", wb), "b");
  EXPECT_EQ(text_as(ExcelLocale::kDeDE, "=CELL(\"format\",A1)", wb), "S");
  EXPECT_EQ(text_as(ExcelLocale::kFrFR, "=CELL(\"type\",A1)", wb), "v");
  EXPECT_EQ(text_as(ExcelLocale::kFrFR, "=CELL(\"type\",A3)", wb), "i");
  EXPECT_EQ(text_as(ExcelLocale::kJaJP, "=CELL(\"type\",A1)", wb), "v");
  EXPECT_EQ(text_as(ExcelLocale::kJaJP, "=CELL(\"format\",A1)", wb), "G");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
