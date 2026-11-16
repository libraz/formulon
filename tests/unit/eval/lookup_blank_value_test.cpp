// Pins how the lookup family types a blank or empty-text lookup_value against blank, "", 0 and "0" keys,
// measured in Mac Excel 365: VLOOKUP / HLOOKUP / MATCH / LOOKUP read a blank lookup_value as 0, XLOOKUP / XMATCH
// keep it blank and sort it above every other kind, and no lookup value but a blank one finds a blank cell.

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
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

// The lookup_value forms: a blank cell, an omitted argument, a "" literal, 0, and a cell holding "".
enum LookupForm : std::uint8_t { kBlankCell, kOmitted, kEmptyLiteral, kZero, kEmptyTextCell };

// The first key of each 3-key table, in table order: blank, "", 0, "0".
constexpr std::size_t kKeyKinds = 4;

std::string Replace(std::string text, std::string_view token, std::string_view with) {
  for (std::size_t at = text.find(token); at != std::string::npos; at = text.find(token, at + with.size())) {
    text.replace(at, token.size(), with);
  }
  return text;
}

std::string Col(std::uint32_t c) {
  return std::string(1, static_cast<char>('A' + c));
}

// Table k: keys [kind k, 7, "x"] over values [20, 10, 30], down columns 3k..3k+1 at rows 1-3 and across rows
// 5+3k..6+3k at columns A:C. Z1 stays blank; Z2 holds "".
Workbook SeedTables() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  const std::array<Value, kKeyKinds> first = {Value::blank(), Value::text(""), Value::number(0.0), Value::text("0")};
  const std::array<Value, 3> values = {Value::number(20.0), Value::number(10.0), Value::number(30.0)};
  for (std::uint32_t k = 0; k < kKeyKinds; ++k) {
    const std::array<Value, 3> keys = {first[k], Value::number(7.0), Value::text("x")};
    for (std::uint32_t i = 0; i < 3U; ++i) {
      if (!keys[i].is_blank()) {
        s.set_cell_value(i, 3U * k, keys[i]);
        s.set_cell_value(4U + 3U * k, i, keys[i]);
      }
      s.set_cell_value(i, 3U * k + 1U, values[i]);
      s.set_cell_value(5U + 3U * k, i, values[i]);
    }
  }
  s.set_cell_value(1U, 25U, Value::text(""));
  return wb;
}

std::string Instantiate(std::string_view tmpl, LookupForm form, std::uint32_t k) {
  static constexpr std::array<std::string_view, 5> kLookup = {"Z1", "", "\"\"", "0", "Z2"};
  const std::string key = Col(3U * k);
  const std::string value = Col(3U * k + 1U);
  const std::string hrow_key = std::to_string(5U + 3U * k);
  const std::string hrow_value = std::to_string(6U + 3U * k);
  std::string out(tmpl);
  out = Replace(out, "{L}", kLookup[form]);
  out = Replace(out, "{T}", key + "1:" + value + "3");
  out = Replace(out, "{K}", key + "1:" + key + "3");
  out = Replace(out, "{V}", value + "1:" + value + "3");
  out = Replace(out, "{H}", "A" + hrow_key + ":C" + hrow_value);
  return out;
}

std::string Render(const Value& v) {
  if (v.is_error()) {
    return v.as_error() == ErrorCode::NA ? "#N/A" : v.debug_to_string();
  }
  if (v.is_number()) {
    const double n = v.as_number();
    return n == static_cast<double>(static_cast<std::int64_t>(n)) ? std::to_string(static_cast<std::int64_t>(n))
                                                                  : v.debug_to_string();
  }
  return v.debug_to_string();
}

std::string Eval(std::string_view src, const Workbook& wb) {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr) {
    return "parse failure";
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, wb.sheet(0), state);
  return Render(evaluate(*root, eval_arena, default_registry(), ctx));
}

struct Row {
  const char* formula;
  LookupForm form;
  std::array<const char*, kKeyKinds> expected;
};

constexpr Row kRows[] = {
    {"=VLOOKUP({L},{T},2,FALSE)", kBlankCell, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=VLOOKUP({L},{T},2,FALSE)", kOmitted, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=VLOOKUP({L},{T},2,FALSE)", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=VLOOKUP({L},{T},2,FALSE)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=VLOOKUP({L},{T},2,FALSE)", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=VLOOKUP({L},{T},2,TRUE)", kBlankCell, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=VLOOKUP({L},{T},2,TRUE)", kOmitted, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=VLOOKUP({L},{T},2,TRUE)", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=VLOOKUP({L},{T},2,TRUE)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=VLOOKUP({L},{T},2,TRUE)", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=HLOOKUP({L},{H},2,FALSE)", kBlankCell, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=HLOOKUP({L},{H},2,FALSE)", kOmitted, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=HLOOKUP({L},{H},2,FALSE)", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=HLOOKUP({L},{H},2,FALSE)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=HLOOKUP({L},{H},2,FALSE)", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=HLOOKUP({L},{H},2,TRUE)", kBlankCell, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=HLOOKUP({L},{H},2,TRUE)", kOmitted, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=HLOOKUP({L},{H},2,TRUE)", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=HLOOKUP({L},{H},2,TRUE)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=HLOOKUP({L},{H},2,TRUE)", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=MATCH({L},{K},0)", kBlankCell, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=MATCH({L},{K},0)", kOmitted, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=MATCH({L},{K},0)", kEmptyLiteral, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=MATCH({L},{K},0)", kZero, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=MATCH({L},{K},0)", kEmptyTextCell, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=MATCH({L},{K},1)", kBlankCell, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=MATCH({L},{K},1)", kOmitted, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=MATCH({L},{K},1)", kEmptyLiteral, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=MATCH({L},{K},1)", kZero, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=MATCH({L},{K},1)", kEmptyTextCell, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=MATCH({L},{K},-1)", kBlankCell, {"2", "2", "1", "2"}},
    {"=MATCH({L},{K},-1)", kOmitted, {"2", "2", "1", "2"}},
    {"=MATCH({L},{K},-1)", kEmptyLiteral, {"3", "1", "3", "3"}},
    {"=MATCH({L},{K},-1)", kZero, {"2", "2", "1", "2"}},
    {"=MATCH({L},{K},-1)", kEmptyTextCell, {"3", "1", "3", "3"}},
    {"=XLOOKUP({L},{K},{V},,0)", kBlankCell, {"20", "#N/A", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,0)", kOmitted, {"20", "#N/A", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,0)", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,0)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,0)", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,-1)", kBlankCell, {"20", "30", "30", "30"}},
    {"=XLOOKUP({L},{K},{V},,-1)", kOmitted, {"20", "30", "30", "30"}},
    {"=XLOOKUP({L},{K},{V},,-1)", kEmptyLiteral, {"10", "20", "10", "10"}},
    {"=XLOOKUP({L},{K},{V},,-1)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,-1)", kEmptyTextCell, {"10", "20", "10", "10"}},
    {"=XLOOKUP({L},{K},{V},,1)", kBlankCell, {"20", "#N/A", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,1)", kOmitted, {"20", "#N/A", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,1)", kEmptyLiteral, {"30", "20", "30", "20"}},
    {"=XLOOKUP({L},{K},{V},,1)", kZero, {"10", "10", "20", "10"}},
    {"=XLOOKUP({L},{K},{V},,1)", kEmptyTextCell, {"30", "20", "30", "20"}},
    {"=XLOOKUP({L},{K},{V},,2)", kBlankCell, {"20", "#N/A", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,2)", kOmitted, {"20", "#N/A", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,2)", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,2)", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=XLOOKUP({L},{K},{V},,2)", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},0)", kBlankCell, {"1", "#N/A", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},0)", kOmitted, {"1", "#N/A", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},0)", kEmptyLiteral, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},0)", kZero, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=XMATCH({L},{K},0)", kEmptyTextCell, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},-1)", kBlankCell, {"1", "3", "3", "3"}},
    {"=XMATCH({L},{K},-1)", kOmitted, {"1", "3", "3", "3"}},
    {"=XMATCH({L},{K},-1)", kEmptyLiteral, {"2", "1", "2", "2"}},
    {"=XMATCH({L},{K},-1)", kZero, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=XMATCH({L},{K},-1)", kEmptyTextCell, {"2", "1", "2", "2"}},
    {"=XMATCH({L},{K},1)", kBlankCell, {"1", "#N/A", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},1)", kOmitted, {"1", "#N/A", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},1)", kEmptyLiteral, {"3", "1", "3", "1"}},
    {"=XMATCH({L},{K},1)", kZero, {"2", "2", "1", "2"}},
    {"=XMATCH({L},{K},1)", kEmptyTextCell, {"3", "1", "3", "1"}},
    {"=XMATCH({L},{K},2)", kBlankCell, {"1", "#N/A", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},2)", kOmitted, {"1", "#N/A", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},2)", kEmptyLiteral, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=XMATCH({L},{K},2)", kZero, {"#N/A", "#N/A", "1", "#N/A"}},
    {"=XMATCH({L},{K},2)", kEmptyTextCell, {"#N/A", "1", "#N/A", "#N/A"}},
    {"=LOOKUP({L},{K},{V})", kBlankCell, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=LOOKUP({L},{K},{V})", kOmitted, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=LOOKUP({L},{K},{V})", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=LOOKUP({L},{K},{V})", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=LOOKUP({L},{K},{V})", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=INDEX({V},MATCH({L},{K},0))", kBlankCell, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=INDEX({V},MATCH({L},{K},0))", kOmitted, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=INDEX({V},MATCH({L},{K},0))", kEmptyLiteral, {"#N/A", "20", "#N/A", "#N/A"}},
    {"=INDEX({V},MATCH({L},{K},0))", kZero, {"#N/A", "#N/A", "20", "#N/A"}},
    {"=INDEX({V},MATCH({L},{K},0))", kEmptyTextCell, {"#N/A", "20", "#N/A", "#N/A"}},
};

TEST(LookupBlankValue, EveryFamilyTypesBlankAndEmptyLookupsAsExcelDoes) {
  const Workbook wb = SeedTables();
  for (const Row& row : kRows) {
    for (std::uint32_t k = 0; k < kKeyKinds; ++k) {
      const std::string formula = Instantiate(row.formula, row.form, k);
      EXPECT_EQ(Eval(formula, wb), row.expected[k]) << formula;
    }
  }
}

// XLOOKUP / XMATCH sort a blank above booleans and errors too.
TEST(LookupBlankValue, XmatchSortsBlankAboveEveryOtherKind) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::boolean(true));
  s.set_cell_value(1U, 0U, Value::boolean(false));
  s.set_cell_value(2U, 0U, Value::error(ErrorCode::NA));
  s.set_cell_value(3U, 0U, Value::error(ErrorCode::Div0));
  s.set_cell_value(0U, 2U, Value::boolean(true));
  s.set_cell_value(1U, 2U, Value::text("x"));
  s.set_cell_value(2U, 2U, Value::number(7.0));
  for (const char* key : {"A1", "A2", "A3", "A4"}) {
    EXPECT_EQ(Eval(std::string("=XMATCH(Z1,") + key + ",-1)", wb), "1") << key;
    EXPECT_EQ(Eval(std::string("=XMATCH(Z1,") + key + ",1)", wb), "#N/A") << key;
  }
  for (const char* lookup : {"TRUE", "FALSE"}) {
    EXPECT_EQ(Eval(std::string("=XMATCH(") + lookup + ",Z1,-1)", wb), "#N/A") << lookup;
    EXPECT_EQ(Eval(std::string("=XMATCH(") + lookup + ",Z1,1)", wb), "1") << lookup;
  }
  EXPECT_EQ(Eval("=XMATCH(Z1,C1:C3,-1)", wb), "1");
  EXPECT_EQ(Eval("=XMATCH(Z1,C1:C3,1)", wb), "#N/A");
}

// A number above every numeric key takes the first blank key as the next larger one, from a range or an array.
TEST(LookupBlankValue, XlookupNextLargerTakesTheFirstBlankKey) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  for (std::uint32_t c = 0; c < 5U; ++c) {
    s.set_cell_value(0U, c + 1U, Value::number(10.0 * (c + 1U)));
  }
  s.set_cell_value(1U, 1U, Value::number(0.1));
  s.set_cell_value(1U, 3U, Value::number(0.2));
  s.set_cell_value(1U, 4U, Value::number(0.3));
  s.set_cell_value(1U, 5U, Value::number(0.4));
  for (std::uint32_t c = 0; c < 5U; ++c) {
    s.set_cell_value(2U, c + 1U, Value::number(0.1 * (c + 1U)));
  }
  EXPECT_EQ(Eval("=XLOOKUP(0.5,B2:F2,B1:F1,\"\",1)", wb), "20");
  EXPECT_EQ(Eval("=XLOOKUP(0.5,FILTER(B2:F2,TRUE),B1:F1,\"\",1)", wb), "20");
  EXPECT_EQ(Eval("=XLOOKUP(0.5,B3:F3,B1:F1,\"\",1)", wb), "50");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
