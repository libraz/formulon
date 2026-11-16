//
// A text prefix criterion in a database-function criteria range matches Text
// records only, never a blank, number or boolean record. The zero-length
// string (constant `""` or formula `=""`) is the empty prefix and matches every
// Text record, zero-length text included. Mac Excel 365 ja-JP measured. A blank criterion cell, `"="` and `"<>"` keep
// their own rules, and COUNTIF's `""` criterion (blank plus zero-length text) is a different rule.
//
// Layout: database A8:B11 (header k / v), keys [special, 7, "x"] in A9:A11,
// values [1, 2, 4] in B9:B11, criteria B14:B15 (header k, criterion in B15),
// COUNTIF fixture D9:D11 = [blank, "", 5]. The prefix cases use their own
// database D1:D5 = [k, TRUE, FALSE, 1E+20, "FALx"] with the criterion in F2.

#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx = test::workbook_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

enum class Special { Blank, EmptyText, Zero, False };

enum class Crit { Blank, EmptyText, Equals, NotEquals };

Workbook MakeWorkbook(Special special, Crit crit) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(7, 0, "k");
  s.set_cell_text(7, 1, "v");
  switch (special) {
    case Special::Blank:
      break;
    case Special::EmptyText:
      s.set_cell_text(8, 0, "");
      break;
    case Special::Zero:
      s.set_cell_value(8, 0, Value::number(0.0));
      break;
    case Special::False:
      s.set_cell_value(8, 0, Value::boolean(false));
      break;
  }
  s.set_cell_value(9, 0, Value::number(7.0));
  s.set_cell_text(10, 0, "x");
  s.set_cell_value(8, 1, Value::number(1.0));
  s.set_cell_value(9, 1, Value::number(2.0));
  s.set_cell_value(10, 1, Value::number(4.0));

  s.set_cell_text(13, 1, "k");
  switch (crit) {
    case Crit::Blank:
      break;
    case Crit::EmptyText:
      s.set_cell_text(14, 1, "");
      break;
    case Crit::Equals:
      s.set_cell_text(14, 1, "=");
      break;
    case Crit::NotEquals:
      s.set_cell_text(14, 1, "<>");
      break;
  }

  s.set_cell_text(9, 3, "");
  s.set_cell_value(10, 3, Value::number(5.0));
  return wb;
}

Value Eval(std::string_view src, Special special, Crit crit) {
  const Workbook wb = MakeWorkbook(special, crit);
  return EvalSourceIn(src, wb, wb.sheet(0));
}

void ExpectNumber(const Value& v, double expected) {
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_DOUBLE_EQ(v.as_number(), expected);
}

void ExpectError(const Value& v, ErrorCode expected) {
  ASSERT_TRUE(v.is_error()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), expected);
}

constexpr std::string_view kDsum = "=DSUM($A$8:$B$11,\"v\",$B$14:$B$15)";

TEST(DFuncEmptyCriterion, DsumMatchesTextRecordsOnly) {
  ExpectNumber(Eval(kDsum, Special::Blank, Crit::EmptyText), 4.0);
  ExpectNumber(Eval(kDsum, Special::EmptyText, Crit::EmptyText), 5.0);
  ExpectNumber(Eval(kDsum, Special::Zero, Crit::EmptyText), 4.0);
  ExpectNumber(Eval(kDsum, Special::False, Crit::EmptyText), 4.0);
}

TEST(DFuncEmptyCriterion, DgetCountsTextRecordsOnly) {
  constexpr std::string_view kDget = "=DGET($A$8:$B$11,\"v\",$B$14:$B$15)";
  ExpectNumber(Eval(kDget, Special::Blank, Crit::EmptyText), 4.0);
  ExpectError(Eval(kDget, Special::EmptyText, Crit::EmptyText), ErrorCode::Num);
  ExpectNumber(Eval(kDget, Special::Zero, Crit::EmptyText), 4.0);
}

TEST(DFuncEmptyCriterion, DcountaCountsZeroLengthTextKey) {
  constexpr std::string_view kDcounta = "=DCOUNTA($A$8:$B$11,\"k\",$B$14:$B$15)";
  ExpectNumber(Eval(kDcounta, Special::EmptyText, Crit::EmptyText), 2.0);
  ExpectNumber(Eval(kDcounta, Special::Blank, Crit::EmptyText), 1.0);
}

TEST(DFuncEmptyCriterion, DaverageSkipsNonTextRecords) {
  constexpr std::string_view kDaverage = "=DAVERAGE($A$8:$B$11,\"v\",$B$14:$B$15)";
  ExpectNumber(Eval(kDaverage, Special::Zero, Crit::EmptyText), 4.0);
  ExpectNumber(Eval(kDaverage, Special::EmptyText, Crit::EmptyText), 2.5);
}

TEST(DFuncEmptyCriterion, BlankCriterionCellMatchesEveryRecord) {
  ExpectNumber(Eval(kDsum, Special::Blank, Crit::Blank), 7.0);
}

TEST(DFuncEmptyCriterion, BareEqualsMatchesOnlyTheBlankRecord) {
  ExpectNumber(Eval(kDsum, Special::Blank, Crit::Equals), 1.0);
  ExpectNumber(Eval(kDsum, Special::EmptyText, Crit::Equals), 0.0);
}

TEST(DFuncEmptyCriterion, BareNotEqualsMatchesEveryNonBlankRecord) {
  ExpectNumber(Eval(kDsum, Special::Blank, Crit::NotEquals), 6.0);
  ExpectNumber(Eval(kDsum, Special::EmptyText, Crit::NotEquals), 7.0);
}

TEST(DFuncEmptyCriterion, CountIfEmptyCriterionStillMatchesBlankAndEmptyText) {
  ExpectNumber(Eval("=COUNTIF($D$9:$D$11,\"\")", Special::Blank, Crit::Blank), 2.0);
}

double CountaWithPrefix(std::string_view prefix) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(0, 3, "k");
  s.set_cell_value(1, 3, Value::boolean(true));
  s.set_cell_value(2, 3, Value::boolean(false));
  s.set_cell_value(3, 3, Value::number(1e20));
  s.set_cell_text(4, 3, "FALx");
  s.set_cell_text(0, 5, "k");
  s.set_cell_text(1, 5, prefix);
  const Value v = EvalSourceIn("=DCOUNTA(D1:D5,\"k\",F1:F2)", wb, wb.sheet(0));
  EXPECT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  return v.is_number() ? v.as_number() : -1.0;
}

TEST(DFuncEmptyCriterion, TextPrefixNeverMatchesBooleanOrNumberRecords) {
  EXPECT_DOUBLE_EQ(CountaWithPrefix("FAL"), 1.0);
  EXPECT_DOUBLE_EQ(CountaWithPrefix("TR"), 0.0);
  EXPECT_DOUBLE_EQ(CountaWithPrefix("1E"), 0.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
