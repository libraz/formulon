#include "cf/validation_eval.h"

#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

enum class Ty : std::uint8_t { kNone, kWhole, kDecimal, kList, kDate, kTime, kTextLength, kCustom };
enum class Op : std::uint8_t {
  kBetween,
  kNotBetween,
  kEqual,
  kNotEqual,
  kGreaterThan,
  kLessThan,
  kGreaterThanOrEqual,
  kLessThanOrEqual,
  kNone
};
enum class In : std::uint8_t { kNum, kNumtext, kText, kBlank, kDate, kDatetime };
enum class Src : std::uint8_t { kNone, kInline, kRange };
enum class ListSrc : std::uint8_t { kInline, kRange, kInlineNum };

Value Text(const char* s) {
  return Value::text(s);
}
Value Blank() {
  return Value::blank();
}

Value InputValue(In in) {
  switch (in) {
    case In::kNum:
      return Value::number(5.0);
    case In::kNumtext:
      return Text("5");
    case In::kText:
      return Text("abcde");
    case In::kBlank:
      return Blank();
    case In::kDate:
      return Value::number(45306.0);
    case In::kDatetime:
      return Value::number(45306.5);
  }
  return Blank();
}

/// A one-sheet workbook with the list source `H1:H3` = a, b, c and a single rule on A1.
struct Fixture {
  Workbook wb = Workbook::create_empty();
  std::size_t sheet_index = 0;

  Fixture() {
    sheet_index = wb.add_sheet("S");
    wb.set_cell_text(sheet_index, 0, 7, "a");
    wb.set_cell_text(sheet_index, 1, 7, "b");
    wb.set_cell_text(sheet_index, 2, 7, "c");
  }

  Sheet& sheet() { return wb.sheet(sheet_index); }

  void AddRule(Ty ty, Op op, bool allow_blank, std::string f1, std::string f2 = "", std::uint32_t row = 0,
               std::uint32_t col = 0) {
    DataValidation dv;
    dv.ranges.push_back(MergeRange{row, col, row, col});
    dv.type = static_cast<std::uint8_t>(ty);
    dv.op = op == Op::kNone ? 0 : static_cast<std::uint8_t>(op);
    dv.allow_blank = allow_blank;
    dv.formula1 = std::move(f1);
    dv.formula2 = std::move(f2);
    sheet().mutable_validations().push_back(std::move(dv));
  }

  void Store(const Value& v, std::uint32_t row = 0, std::uint32_t col = 0) {
    if (v.is_text()) {
      wb.set_cell_text(sheet_index, row, col, std::string(v.as_text()));
    } else if (!v.is_blank()) {
      wb.set_cell_value(sheet_index, row, col, v);
    }
  }

  ValidationOutcome Check(const Value& v, std::uint32_t row = 0, std::uint32_t col = 0) {
    auto r = validate_value(wb, sheet(), row, col, v);
    EXPECT_TRUE(static_cast<bool>(r));
    return r.value();
  }
};

std::string BoundOf(Ty ty, int which) {
  switch (ty) {
    case Ty::kDate:
      return which == 1 ? "45292" : "45657";
    case Ty::kTime:
      return which == 1 ? "0.25" : "0.75";
    case Ty::kTextLength:
    case Ty::kWhole:
    case Ty::kDecimal:
      if (which == 1) {
        return "3";
      }
      return "8";
    default:
      return "3";
  }
}

bool IsRange(Op op) {
  return op == Op::kBetween || op == Op::kNotBetween;
}

/// Bounds the probes used: 3/8 (whole, decimal, textLength), 45292/45657 (date), 0.25/0.75 (time).
/// Single-bound operators carry only formula1: 3, 45292 and 0.25 respectively.
bool CheckMatrixCell(Ty ty, Op op, bool allow_blank, In in) {
  Fixture f;
  f.AddRule(ty, op, allow_blank, BoundOf(ty, 1), IsRange(op) ? BoundOf(ty, 2) : "");
  const Value v = InputValue(in);
  f.Store(v);
  return f.Check(v).valid;
}

// Matrix of results (Mac Excel `Validation.Value`), one row per type/operator/allowBlank;
// `expect` lists validity for num, numtext, text, blank, date, datetime.
struct MatrixRow {
  Ty ty;
  Op op;
  bool allow_blank;
  const char* expect;
};

const MatrixRow kMatrix[] = {
    {Ty::kWhole, Op::kBetween, true, "100100"},
    {Ty::kWhole, Op::kBetween, false, "100000"},
    {Ty::kWhole, Op::kNotBetween, true, "000110"},
    {Ty::kWhole, Op::kNotBetween, false, "000110"},
    {Ty::kWhole, Op::kEqual, true, "000100"},
    {Ty::kWhole, Op::kEqual, false, "000000"},
    {Ty::kWhole, Op::kNotEqual, true, "100110"},
    {Ty::kWhole, Op::kNotEqual, false, "100110"},
    {Ty::kWhole, Op::kGreaterThan, true, "100110"},
    {Ty::kWhole, Op::kGreaterThan, false, "100010"},
    {Ty::kWhole, Op::kLessThan, true, "000100"},
    {Ty::kWhole, Op::kLessThan, false, "000100"},
    {Ty::kWhole, Op::kGreaterThanOrEqual, true, "100110"},
    {Ty::kWhole, Op::kGreaterThanOrEqual, false, "100010"},
    {Ty::kWhole, Op::kLessThanOrEqual, true, "000100"},
    {Ty::kWhole, Op::kLessThanOrEqual, false, "000100"},
    {Ty::kDecimal, Op::kBetween, true, "100100"},
    {Ty::kDecimal, Op::kBetween, false, "100000"},
    {Ty::kDecimal, Op::kNotBetween, true, "000111"},
    {Ty::kDecimal, Op::kNotBetween, false, "000111"},
    {Ty::kDecimal, Op::kEqual, true, "000100"},
    {Ty::kDecimal, Op::kEqual, false, "000000"},
    {Ty::kDecimal, Op::kNotEqual, true, "100111"},
    {Ty::kDecimal, Op::kNotEqual, false, "100111"},
    {Ty::kDecimal, Op::kGreaterThan, true, "100111"},
    {Ty::kDecimal, Op::kGreaterThan, false, "100011"},
    {Ty::kDecimal, Op::kLessThan, true, "000100"},
    {Ty::kDecimal, Op::kLessThan, false, "000100"},
    {Ty::kDecimal, Op::kGreaterThanOrEqual, true, "100111"},
    {Ty::kDecimal, Op::kGreaterThanOrEqual, false, "100011"},
    {Ty::kDecimal, Op::kLessThanOrEqual, true, "000100"},
    {Ty::kDecimal, Op::kLessThanOrEqual, false, "000100"},
    {Ty::kDate, Op::kBetween, true, "000111"},
    {Ty::kDate, Op::kBetween, false, "000011"},
    {Ty::kDate, Op::kNotBetween, true, "100100"},
    {Ty::kDate, Op::kNotBetween, false, "100100"},
    {Ty::kDate, Op::kEqual, true, "000100"},
    {Ty::kDate, Op::kEqual, false, "000000"},
    {Ty::kDate, Op::kNotEqual, true, "100111"},
    {Ty::kDate, Op::kNotEqual, false, "100111"},
    {Ty::kDate, Op::kGreaterThan, true, "000111"},
    {Ty::kDate, Op::kGreaterThan, false, "000011"},
    {Ty::kDate, Op::kLessThan, true, "100100"},
    {Ty::kDate, Op::kLessThan, false, "100100"},
    {Ty::kDate, Op::kGreaterThanOrEqual, true, "000111"},
    {Ty::kDate, Op::kGreaterThanOrEqual, false, "000011"},
    {Ty::kDate, Op::kLessThanOrEqual, true, "100100"},
    {Ty::kDate, Op::kLessThanOrEqual, false, "100100"},
    {Ty::kTime, Op::kBetween, true, "000100"},
    {Ty::kTime, Op::kBetween, false, "000000"},
    {Ty::kTime, Op::kNotBetween, true, "100111"},
    {Ty::kTime, Op::kNotBetween, false, "100111"},
    {Ty::kTime, Op::kEqual, true, "000100"},
    {Ty::kTime, Op::kEqual, false, "000000"},
    {Ty::kTime, Op::kNotEqual, true, "100111"},
    {Ty::kTime, Op::kNotEqual, false, "100111"},
    {Ty::kTime, Op::kGreaterThan, true, "100111"},
    {Ty::kTime, Op::kGreaterThan, false, "100011"},
    {Ty::kTime, Op::kLessThan, true, "000100"},
    {Ty::kTime, Op::kLessThan, false, "000100"},
    {Ty::kTime, Op::kGreaterThanOrEqual, true, "100111"},
    {Ty::kTime, Op::kGreaterThanOrEqual, false, "100011"},
    {Ty::kTime, Op::kLessThanOrEqual, true, "000100"},
    {Ty::kTime, Op::kLessThanOrEqual, false, "000100"},
    {Ty::kTextLength, Op::kBetween, true, "001111"},
    {Ty::kTextLength, Op::kBetween, false, "001011"},
    {Ty::kTextLength, Op::kNotBetween, true, "110100"},
    {Ty::kTextLength, Op::kNotBetween, false, "110100"},
    {Ty::kTextLength, Op::kEqual, true, "000100"},
    {Ty::kTextLength, Op::kEqual, false, "000000"},
    {Ty::kTextLength, Op::kNotEqual, true, "111111"},
    {Ty::kTextLength, Op::kNotEqual, false, "111111"},
    {Ty::kTextLength, Op::kGreaterThan, true, "001111"},
    {Ty::kTextLength, Op::kGreaterThan, false, "001011"},
    {Ty::kTextLength, Op::kLessThan, true, "110100"},
    {Ty::kTextLength, Op::kLessThan, false, "110100"},
    {Ty::kTextLength, Op::kGreaterThanOrEqual, true, "001111"},
    {Ty::kTextLength, Op::kGreaterThanOrEqual, false, "001011"},
    {Ty::kTextLength, Op::kLessThanOrEqual, true, "110100"},
    {Ty::kTextLength, Op::kLessThanOrEqual, false, "110100"},
};

constexpr In kInputs[] = {In::kNum, In::kNumtext, In::kText, In::kBlank, In::kDate, In::kDatetime};

struct ListRow {
  ListSrc src;
  bool allow_blank;
  Value input;
  bool valid;
};

const ListRow kListRows[] = {
    {ListSrc::kInline, true, Text("a"), true},
    {ListSrc::kInline, true, Text("A"), false},
    {ListSrc::kInline, true, Text("B"), false},
    {ListSrc::kInline, true, Text("1"), false},
    {ListSrc::kInline, true, Text("d"), false},
    {ListSrc::kInline, true, Blank(), true},
    {ListSrc::kInline, true, Value::number(1.0), false},
    {ListSrc::kInline, false, Text("a"), true},
    {ListSrc::kInline, false, Text("A"), false},
    {ListSrc::kInline, false, Text("B"), false},
    {ListSrc::kInline, false, Text("1"), false},
    {ListSrc::kInline, false, Text("d"), false},
    {ListSrc::kInline, false, Blank(), false},
    {ListSrc::kInline, false, Value::number(1.0), false},
    {ListSrc::kRange, true, Text("a"), true},
    {ListSrc::kRange, true, Text("A"), true},
    {ListSrc::kRange, true, Text("B"), true},
    {ListSrc::kRange, true, Text("1"), false},
    {ListSrc::kRange, true, Text("d"), false},
    {ListSrc::kRange, true, Blank(), true},
    {ListSrc::kRange, true, Value::number(1.0), false},
    {ListSrc::kRange, false, Text("a"), true},
    {ListSrc::kRange, false, Text("A"), true},
    {ListSrc::kRange, false, Text("B"), true},
    {ListSrc::kRange, false, Text("1"), false},
    {ListSrc::kRange, false, Text("d"), false},
    {ListSrc::kRange, false, Blank(), false},
    {ListSrc::kRange, false, Value::number(1.0), false},
    {ListSrc::kInlineNum, true, Text("a"), false},
    {ListSrc::kInlineNum, true, Text("A"), false},
    {ListSrc::kInlineNum, true, Text("B"), false},
    {ListSrc::kInlineNum, true, Value::number(1.0), true},
    {ListSrc::kInlineNum, true, Text("d"), false},
    {ListSrc::kInlineNum, true, Blank(), true},
    {ListSrc::kInlineNum, true, Value::number(1.0), true},
    {ListSrc::kInlineNum, false, Text("a"), false},
    {ListSrc::kInlineNum, false, Text("A"), false},
    {ListSrc::kInlineNum, false, Text("B"), false},
    {ListSrc::kInlineNum, false, Value::number(1.0), true},
    {ListSrc::kInlineNum, false, Text("d"), false},
    {ListSrc::kInlineNum, false, Blank(), false},
    {ListSrc::kInlineNum, false, Value::number(1.0), true},
};

struct CustomRow {
  const char* formula;
  bool allow_blank;
  In in;
  bool valid;
};

const CustomRow kCustomRows[] = {
    {"ISNUMBER(A1)", true, In::kNum, true},
    {"ISNUMBER(A1)", true, In::kNumtext, false},
    {"ISNUMBER(A1)", true, In::kText, false},
    {"ISNUMBER(A1)", true, In::kBlank, true},
    {"ISNUMBER(A1)", true, In::kDate, true},
    {"ISNUMBER(A1)", true, In::kDatetime, true},
    {"ISNUMBER(A1)", false, In::kNum, true},
    {"ISNUMBER(A1)", false, In::kNumtext, false},
    {"ISNUMBER(A1)", false, In::kText, false},
    {"ISNUMBER(A1)", false, In::kBlank, false},
    {"ISNUMBER(A1)", false, In::kDate, true},
    {"ISNUMBER(A1)", false, In::kDatetime, true},
    {"A1>3", true, In::kNum, true},
    {"A1>3", true, In::kNumtext, true},
    {"A1>3", true, In::kText, true},
    {"A1>3", true, In::kBlank, true},
    {"A1>3", true, In::kDate, true},
    {"A1>3", true, In::kDatetime, true},
    {"A1>3", false, In::kNum, true},
    {"A1>3", false, In::kNumtext, true},
    {"A1>3", false, In::kText, true},
    {"A1>3", false, In::kBlank, false},
    {"A1>3", false, In::kDate, true},
    {"A1>3", false, In::kDatetime, true},
    {"A1", true, In::kNum, true},
    {"A1", true, In::kNumtext, false},
    {"A1", true, In::kText, false},
    {"A1", true, In::kBlank, true},
    {"A1", true, In::kDate, true},
    {"A1", true, In::kDatetime, true},
    {"A1", false, In::kNum, true},
    {"A1", false, In::kNumtext, false},
    {"A1", false, In::kText, false},
    {"A1", false, In::kBlank, false},
    {"A1", false, In::kDate, true},
    {"A1", false, In::kDatetime, true},
    {"LEN(A1)>2", true, In::kNum, false},
    {"LEN(A1)>2", true, In::kNumtext, false},
    {"LEN(A1)>2", true, In::kText, true},
    {"LEN(A1)>2", true, In::kBlank, true},
    {"LEN(A1)>2", true, In::kDate, true},
    {"LEN(A1)>2", true, In::kDatetime, true},
    {"LEN(A1)>2", false, In::kNum, false},
    {"LEN(A1)>2", false, In::kNumtext, false},
    {"LEN(A1)>2", false, In::kText, true},
    {"LEN(A1)>2", false, In::kBlank, false},
    {"LEN(A1)>2", false, In::kDate, true},
    {"LEN(A1)>2", false, In::kDatetime, true},
    {"TRUE", true, In::kNum, true},
    {"TRUE", true, In::kNumtext, true},
    {"TRUE", true, In::kText, true},
    {"TRUE", true, In::kBlank, true},
    {"TRUE", true, In::kDate, true},
    {"TRUE", true, In::kDatetime, true},
    {"TRUE", false, In::kNum, true},
    {"TRUE", false, In::kNumtext, true},
    {"TRUE", false, In::kText, true},
    {"TRUE", false, In::kBlank, true},
    {"TRUE", false, In::kDate, true},
    {"TRUE", false, In::kDatetime, true},
    {"FALSE", true, In::kNum, false},
    {"FALSE", true, In::kNumtext, false},
    {"FALSE", true, In::kText, false},
    {"FALSE", true, In::kBlank, true},
    {"FALSE", true, In::kDate, false},
    {"FALSE", true, In::kDatetime, false},
    {"FALSE", false, In::kNum, false},
    {"FALSE", false, In::kNumtext, false},
    {"FALSE", false, In::kText, false},
    {"FALSE", false, In::kBlank, false},
    {"FALSE", false, In::kDate, false},
    {"FALSE", false, In::kDatetime, false},
    {"1", true, In::kNum, true},
    {"1", true, In::kNumtext, true},
    {"1", true, In::kText, true},
    {"1", true, In::kBlank, true},
    {"1", true, In::kDate, true},
    {"1", true, In::kDatetime, true},
    {"1", false, In::kNum, true},
    {"1", false, In::kNumtext, true},
    {"1", false, In::kText, true},
    {"1", false, In::kBlank, true},
    {"1", false, In::kDate, true},
    {"1", false, In::kDatetime, true},
    {"0", true, In::kNum, false},
    {"0", true, In::kNumtext, false},
    {"0", true, In::kText, false},
    {"0", true, In::kBlank, true},
    {"0", true, In::kDate, false},
    {"0", true, In::kDatetime, false},
    {"0", false, In::kNum, false},
    {"0", false, In::kNumtext, false},
    {"0", false, In::kText, false},
    {"0", false, In::kBlank, false},
    {"0", false, In::kDate, false},
    {"0", false, In::kDatetime, false},
    {"\"x\"", true, In::kNum, false},
    {"\"x\"", true, In::kNumtext, false},
    {"\"x\"", true, In::kText, false},
    {"\"x\"", true, In::kBlank, true},
    {"\"x\"", true, In::kDate, false},
    {"\"x\"", true, In::kDatetime, false},
    {"\"x\"", false, In::kNum, false},
    {"\"x\"", false, In::kNumtext, false},
    {"\"x\"", false, In::kText, false},
    {"\"x\"", false, In::kBlank, false},
    {"\"x\"", false, In::kDate, false},
    {"\"x\"", false, In::kDatetime, false},
    {"A1=\"abcde\"", true, In::kNum, false},
    {"A1=\"abcde\"", true, In::kNumtext, false},
    {"A1=\"abcde\"", true, In::kText, true},
    {"A1=\"abcde\"", true, In::kBlank, true},
    {"A1=\"abcde\"", true, In::kDate, false},
    {"A1=\"abcde\"", true, In::kDatetime, false},
    {"A1=\"abcde\"", false, In::kNum, false},
    {"A1=\"abcde\"", false, In::kNumtext, false},
    {"A1=\"abcde\"", false, In::kText, true},
    {"A1=\"abcde\"", false, In::kBlank, false},
    {"A1=\"abcde\"", false, In::kDate, false},
    {"A1=\"abcde\"", false, In::kDatetime, false},
};

const char* ListFormula(ListSrc src) {
  switch (src) {
    case ListSrc::kInline:
      return "\"a,b,c\"";
    case ListSrc::kRange:
      return "$H$1:$H$3";
    case ListSrc::kInlineNum:
      return "\"1,2,3\"";
  }
  return "";
}

// Pairwise model (greedy all-pairs over valid combinations):
//   type(8): none, whole, decimal, date, time, textLength, list, custom
//   operator(8): between, notBetween, equal, notEqual, greaterThan, lessThan,
//                greaterThanOrEqual, lessThanOrEqual (only for whole/decimal/date/time/textLength)
//   allowBlank(2): true, false
//   input kind(6): number 5, numeric text "5", text "abcde", blank, date 45306, datetime 45306.5
//   list source(2): inline "a,b,c", range H1:H3 = a,b,c (only for list)
// Combinations are 528 valid tuples; the table below is 72 rows covering every pair of
// parameter values that co-occur. Expected values come from the matrix: typed rows
// from the matrix table, list rows are invalid unless blank is allowed (none of the six
// inputs is a list member), custom rows use ISNUMBER(A1), type none is always valid.
struct PairRow {
  Ty ty;
  Op op;
  bool allow_blank;
  In in;
  Src src;
  bool valid;
};

const PairRow kPairs[] = {
    {Ty::kNone, Op::kNone, false, In::kNum, Src::kNone, true},
    {Ty::kNone, Op::kNone, false, In::kText, Src::kNone, true},
    {Ty::kNone, Op::kNone, false, In::kBlank, Src::kNone, true},
    {Ty::kNone, Op::kNone, false, In::kDate, Src::kNone, true},
    {Ty::kNone, Op::kNone, false, In::kDatetime, Src::kNone, true},
    {Ty::kNone, Op::kNone, true, In::kNumtext, Src::kNone, true},
    {Ty::kWhole, Op::kBetween, false, In::kNum, Src::kNone, true},
    {Ty::kWhole, Op::kBetween, false, In::kDatetime, Src::kNone, false},
    {Ty::kWhole, Op::kNotBetween, false, In::kDate, Src::kNone, true},
    {Ty::kWhole, Op::kNotBetween, true, In::kNumtext, Src::kNone, false},
    {Ty::kWhole, Op::kEqual, false, In::kBlank, Src::kNone, false},
    {Ty::kWhole, Op::kEqual, false, In::kDatetime, Src::kNone, false},
    {Ty::kWhole, Op::kNotEqual, false, In::kText, Src::kNone, false},
    {Ty::kWhole, Op::kNotEqual, false, In::kDatetime, Src::kNone, false},
    {Ty::kWhole, Op::kGreaterThan, false, In::kBlank, Src::kNone, false},
    {Ty::kWhole, Op::kGreaterThan, true, In::kDatetime, Src::kNone, false},
    {Ty::kWhole, Op::kLessThan, false, In::kText, Src::kNone, false},
    {Ty::kWhole, Op::kLessThan, false, In::kBlank, Src::kNone, true},
    {Ty::kWhole, Op::kGreaterThanOrEqual, false, In::kNumtext, Src::kNone, false},
    {Ty::kWhole, Op::kGreaterThanOrEqual, false, In::kBlank, Src::kNone, false},
    {Ty::kWhole, Op::kLessThanOrEqual, false, In::kText, Src::kNone, false},
    {Ty::kWhole, Op::kLessThanOrEqual, false, In::kDate, Src::kNone, false},
    {Ty::kDecimal, Op::kBetween, false, In::kDate, Src::kNone, false},
    {Ty::kDecimal, Op::kNotBetween, false, In::kDatetime, Src::kNone, true},
    {Ty::kDecimal, Op::kEqual, false, In::kNumtext, Src::kNone, false},
    {Ty::kDecimal, Op::kNotEqual, true, In::kNum, Src::kNone, true},
    {Ty::kDecimal, Op::kGreaterThan, false, In::kNum, Src::kNone, true},
    {Ty::kDecimal, Op::kLessThan, false, In::kNum, Src::kNone, false},
    {Ty::kDecimal, Op::kGreaterThanOrEqual, true, In::kText, Src::kNone, false},
    {Ty::kDecimal, Op::kLessThanOrEqual, true, In::kBlank, Src::kNone, true},
    {Ty::kDate, Op::kBetween, true, In::kBlank, Src::kNone, true},
    {Ty::kDate, Op::kNotBetween, false, In::kText, Src::kNone, false},
    {Ty::kDate, Op::kEqual, false, In::kNum, Src::kNone, false},
    {Ty::kDate, Op::kNotEqual, false, In::kNumtext, Src::kNone, false},
    {Ty::kDate, Op::kGreaterThan, false, In::kNumtext, Src::kNone, false},
    {Ty::kDate, Op::kLessThan, false, In::kDate, Src::kNone, false},
    {Ty::kDate, Op::kGreaterThanOrEqual, false, In::kDatetime, Src::kNone, true},
    {Ty::kDate, Op::kLessThanOrEqual, false, In::kNum, Src::kNone, true},
    {Ty::kTime, Op::kBetween, false, In::kNumtext, Src::kNone, false},
    {Ty::kTime, Op::kNotBetween, false, In::kNum, Src::kNone, true},
    {Ty::kTime, Op::kEqual, true, In::kText, Src::kNone, false},
    {Ty::kTime, Op::kNotEqual, false, In::kBlank, Src::kNone, true},
    {Ty::kTime, Op::kGreaterThan, false, In::kText, Src::kNone, false},
    {Ty::kTime, Op::kLessThan, false, In::kNumtext, Src::kNone, false},
    {Ty::kTime, Op::kGreaterThanOrEqual, false, In::kDate, Src::kNone, true},
    {Ty::kTime, Op::kLessThanOrEqual, false, In::kDatetime, Src::kNone, false},
    {Ty::kTextLength, Op::kBetween, false, In::kText, Src::kNone, true},
    {Ty::kTextLength, Op::kNotBetween, false, In::kBlank, Src::kNone, true},
    {Ty::kTextLength, Op::kEqual, false, In::kDate, Src::kNone, false},
    {Ty::kTextLength, Op::kNotEqual, false, In::kDate, Src::kNone, true},
    {Ty::kTextLength, Op::kGreaterThan, false, In::kDate, Src::kNone, true},
    {Ty::kTextLength, Op::kLessThan, true, In::kDatetime, Src::kNone, false},
    {Ty::kTextLength, Op::kGreaterThanOrEqual, false, In::kNum, Src::kNone, false},
    {Ty::kTextLength, Op::kLessThanOrEqual, false, In::kNumtext, Src::kNone, true},
    {Ty::kList, Op::kNone, false, In::kNum, Src::kInline, false},
    {Ty::kList, Op::kNone, false, In::kNum, Src::kRange, false},
    {Ty::kList, Op::kNone, false, In::kNumtext, Src::kRange, false},
    {Ty::kList, Op::kNone, false, In::kText, Src::kInline, false},
    {Ty::kList, Op::kNone, false, In::kText, Src::kRange, false},
    {Ty::kList, Op::kNone, false, In::kBlank, Src::kInline, false},
    {Ty::kList, Op::kNone, false, In::kBlank, Src::kRange, false},
    {Ty::kList, Op::kNone, false, In::kDate, Src::kInline, false},
    {Ty::kList, Op::kNone, false, In::kDatetime, Src::kInline, false},
    {Ty::kList, Op::kNone, false, In::kDatetime, Src::kRange, false},
    {Ty::kList, Op::kNone, true, In::kNumtext, Src::kInline, false},
    {Ty::kList, Op::kNone, true, In::kDate, Src::kRange, false},
    {Ty::kCustom, Op::kNone, false, In::kNum, Src::kNone, true},
    {Ty::kCustom, Op::kNone, false, In::kText, Src::kNone, false},
    {Ty::kCustom, Op::kNone, false, In::kBlank, Src::kNone, false},
    {Ty::kCustom, Op::kNone, false, In::kDate, Src::kNone, true},
    {Ty::kCustom, Op::kNone, false, In::kDatetime, Src::kNone, true},
    {Ty::kCustom, Op::kNone, true, In::kNumtext, Src::kNone, false},
};

bool RunPair(const PairRow& p) {
  Fixture f;
  switch (p.ty) {
    case Ty::kList:
      f.AddRule(p.ty, Op::kNone, p.allow_blank, p.src == Src::kInline ? "\"a,b,c\"" : "$H$1:$H$3");
      break;
    case Ty::kCustom:
      f.AddRule(p.ty, Op::kNone, p.allow_blank, "ISNUMBER(A1)");
      break;
    case Ty::kNone:
      f.AddRule(p.ty, Op::kNone, p.allow_blank, "");
      break;
    default:
      f.AddRule(p.ty, p.op, p.allow_blank, BoundOf(p.ty, 1), IsRange(p.op) ? BoundOf(p.ty, 2) : "");
      break;
  }
  const Value v = InputValue(p.in);
  f.Store(v);
  return f.Check(v).valid;
}

TEST(ValidationEval, PairwiseTable) {
  for (const PairRow& p : kPairs) {
    EXPECT_EQ(RunPair(p), p.valid) << "type=" << static_cast<int>(p.ty) << " op=" << static_cast<int>(p.op)
                                   << " allowBlank=" << p.allow_blank << " input=" << static_cast<int>(p.in)
                                   << " src=" << static_cast<int>(p.src);
  }
}

TEST(ValidationEval, ProbeTable) {
  for (const MatrixRow& row : kMatrix) {
    for (int i = 0; i < 6; ++i) {
      EXPECT_EQ(CheckMatrixCell(row.ty, row.op, row.allow_blank, kInputs[i]), row.expect[i] == '1')
          << "type=" << static_cast<int>(row.ty) << " op=" << static_cast<int>(row.op)
          << " allowBlank=" << row.allow_blank << " input=" << i;
    }
  }
  int idx = 0;
  for (const ListRow& row : kListRows) {
    Fixture f;
    f.AddRule(Ty::kList, Op::kNone, row.allow_blank, ListFormula(row.src));
    f.Store(row.input);
    EXPECT_EQ(f.Check(row.input).valid, row.valid) << "list row " << idx;
    ++idx;
  }
  idx = 0;
  for (const CustomRow& row : kCustomRows) {
    Fixture f;
    f.AddRule(Ty::kCustom, Op::kNone, row.allow_blank, row.formula);
    const Value v = InputValue(row.in);
    f.Store(v);
    EXPECT_EQ(f.Check(v).valid, row.valid) << "custom row " << idx << " " << row.formula;
    ++idx;
  }
}

// textLength counts the value text (as LEN does), not the formatted display; numbers that
// General renders with fifteen significant digits count those digits.
TEST(ValidationEval, TextLengthUsesValueText) {
  struct Case {
    Value v;
    const char* len;
    bool valid;
  };
  const Case cases[] = {
      {Value::number(12345.6), "7", true},     {Value::number(12345.6), "8", false},
      {Value::number(5.0), "1", true},         {Value::number(5.0), "3", false},
      {Value::number(0.1), "3", true},         {Value::number(45306.0), "5", true},
      {Value::number(45306.0), "9", false},    {Value::boolean(true), "4", true},
      {Value::number(1e20), "5", true},        {Value::number(1.0 / 3.0), "17", true},
      {Value::number(1.0 / 3.0), "16", false}, {Value::number(1.0 / 3.0), "15", false},
  };
  for (const Case& c : cases) {
    Fixture f;
    f.AddRule(Ty::kTextLength, Op::kEqual, true, c.len);
    EXPECT_EQ(f.Check(c.v).valid, c.valid) << c.len;
  }
  // UTF-16 units: a supplementary-plane character counts twice.
  Fixture f;
  f.AddRule(Ty::kTextLength, Op::kEqual, true, "3");
  EXPECT_TRUE(f.Check(Text("a\xF0\x9F\x98\x80")).valid);
}

TEST(ValidationEval, WholeBoundsFromCellReferences) {
  Fixture f;
  f.wb.set_cell_value(f.sheet_index, 0, 9, Value::number(3.0));
  f.wb.set_cell_value(f.sheet_index, 1, 9, Value::number(8.0));
  f.AddRule(Ty::kWhole, Op::kBetween, true, "$J$1", "$J$2");
  EXPECT_TRUE(f.Check(Value::number(5.0)).valid);
  EXPECT_FALSE(f.Check(Value::number(9.0)).valid);
}

TEST(ValidationEval, WholeRejectsFractionAndAcceptsLargeBound) {
  Fixture f;
  f.AddRule(Ty::kWhole, Op::kGreaterThanOrEqual, true, "-100000000000");
  EXPECT_TRUE(f.Check(Value::number(5.0)).valid);
  EXPECT_FALSE(f.Check(Value::number(5.5)).valid);
  EXPECT_TRUE(f.Check(Value::number(-5.0)).valid);
  EXPECT_TRUE(f.Check(Value::number(1e10)).valid);
}

// Relative references in a rule are shifted from the sqref's top-left cell to the checked cell.
TEST(ValidationEval, RelativeFormulaShiftsFromSqrefAnchor) {
  Fixture f;
  DataValidation dv;
  dv.ranges.push_back(MergeRange{0, 0, 2, 0});  // A1:A3
  dv.type = static_cast<std::uint8_t>(Ty::kCustom);
  dv.formula1 = "B1>0";
  f.sheet().mutable_validations().push_back(dv);
  f.wb.set_cell_value(f.sheet_index, 0, 1, Value::number(1.0));
  f.wb.set_cell_value(f.sheet_index, 1, 1, Value::number(0.0));
  f.wb.set_cell_value(f.sheet_index, 2, 1, Value::number(7.0));
  EXPECT_TRUE(f.Check(Value::number(1.0), 0, 0).valid);
  EXPECT_FALSE(f.Check(Value::number(1.0), 1, 0).valid);
  EXPECT_TRUE(f.Check(Value::number(1.0), 2, 0).valid);
}

TEST(ValidationEval, CustomFormulaUsesProposedValueForScalarRangeAndFormulaReads) {
  // A scalar self-reference is evaluated against the candidate value. The stored cell is
  // deliberately different, and remains unchanged after both validation attempts.
  Fixture scalar;
  scalar.Store(Value::number(5.0));
  scalar.AddRule(Ty::kCustom, Op::kNone, true, "A1>3");
  EXPECT_FALSE(scalar.Check(Value::number(1.0)).valid);
  EXPECT_TRUE(scalar.Check(Value::number(5.0)).valid);
  EXPECT_DOUBLE_EQ(scalar.sheet().resolve_cell_value(0, 0).as_number(), 5.0);

  // Range materialization substitutes the candidate before aggregate evaluation.
  Fixture range;
  range.Store(Value::number(5.0));
  range.wb.set_cell_value(range.sheet_index, 0, 1, Value::number(7.0));
  range.AddRule(Ty::kCustom, Op::kNone, true, "SUM(A1:B1)>10");
  EXPECT_FALSE(range.Check(Value::number(1.0)).valid);
  EXPECT_TRUE(range.Check(Value::number(5.0)).valid);
  EXPECT_DOUBLE_EQ(range.sheet().resolve_cell_value(0, 0).as_number(), 5.0);

  Fixture relative;
  relative.wb.set_cell_value(relative.sheet_index, 1, 1, Value::number(5.0));  // B2
  DataValidation relative_rule;
  relative_rule.ranges.push_back(MergeRange{0, 0, 1, 1});  // A1:B2; anchor A1
  relative_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  relative_rule.formula1 = "A1>3";  // At B2 this shifts to B2.
  relative.sheet().mutable_validations().push_back(relative_rule);
  EXPECT_FALSE(relative.Check(Value::number(1.0), 1, 1).valid);
  EXPECT_TRUE(relative.Check(Value::number(5.0), 1, 1).valid);

  Fixture absolute;
  absolute.wb.set_cell_value(absolute.sheet_index, 1, 1, Value::number(5.0));  // B2
  absolute.AddRule(Ty::kCustom, Op::kNone, true, "$B$2>3", "", 1, 1);
  EXPECT_FALSE(absolute.Check(Value::number(1.0), 1, 1).valid);
  EXPECT_TRUE(absolute.Check(Value::number(5.0), 1, 1).valid);

  // Formula cells must be shadowed before formula-index recursion, rather than returning
  // their cached result. This distinguishes the overlay from a normal cached read.
  Fixture formula;
  ASSERT_TRUE(static_cast<bool>(formula.wb.set_cell_formula(formula.sheet_index, 0, 0, "=5")));
  ASSERT_TRUE(static_cast<bool>(formula.wb.recalc(eval::default_registry())));
  formula.AddRule(Ty::kCustom, Op::kNone, true, "SUM(A1:A1)>3");
  EXPECT_FALSE(formula.Check(Value::number(1.0)).valid);
  EXPECT_TRUE(formula.Check(Value::number(5.0)).valid);
  EXPECT_EQ(formula.sheet().cell_at(0, 0)->formula_text, "=5");

  // A qualified reference to another sheet is unaffected by the target-cell overlay.
  Fixture other;
  const std::size_t other_sheet = other.wb.add_sheet("T");
  other.wb.set_cell_value(other_sheet, 0, 0, Value::number(5.0));
  other.Store(Value::number(9.0));
  other.AddRule(Ty::kCustom, Op::kNone, true, "T!A1>3");
  EXPECT_TRUE(other.Check(Value::number(1.0)).valid);
  EXPECT_DOUBLE_EQ(other.sheet().resolve_cell_value(0, 0).as_number(), 9.0);
}

TEST(ValidationEval, CustomFormulaWholeAxisIncludesProposedValue) {
  // A populated head must not hide a proposed value below the stored extent of
  // the same whole column. The candidate is read only through validation and
  // the backing sheet remains unchanged.
  Fixture sum;
  ASSERT_TRUE(static_cast<bool>(sum.wb.set_cell_value(sum.sheet_index, 0, 0, Value::number(0.0))));
  sum.AddRule(Ty::kCustom, Op::kNone, true, "SUM($A:$A)>0", "", 99, 0);  // A100
  EXPECT_TRUE(sum.Check(Value::number(5.0), 99, 0).valid);
  EXPECT_FALSE(sum.Check(Value::number(0.0), 99, 0).valid);
  EXPECT_TRUE(sum.sheet().resolve_cell_value(99, 0).is_blank());
  EXPECT_DOUBLE_EQ(sum.sheet().resolve_cell_value(0, 0).as_number(), 0.0);

  Fixture countif;
  ASSERT_TRUE(static_cast<bool>(countif.wb.set_cell_value(countif.sheet_index, 0, 0, Value::number(0.0))));
  countif.AddRule(Ty::kCustom, Op::kNone, true, "COUNTIF($A:$A,5)>0", "", 99, 0);  // A100
  EXPECT_TRUE(countif.Check(Value::number(5.0), 99, 0).valid);
  EXPECT_FALSE(countif.Check(Value::number(4.0), 99, 0).valid);

  // With no populated cells at all, the overlay itself supplies the walk's
  // extent. This also guards the zero-length whole-axis fast path.
  Workbook empty = Workbook::create_empty();
  const std::size_t empty_index = empty.add_sheet("Empty");
  Sheet& empty_sheet = empty.sheet(empty_index);
  DataValidation empty_rule;
  empty_rule.ranges.push_back(MergeRange{99, 0, 99, 0});  // A100
  empty_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  empty_rule.allow_blank = true;
  empty_rule.formula1 = "SUM($A:$A)>0";
  empty_sheet.mutable_validations().push_back(empty_rule);
  auto empty_result = validate_value(empty, empty_sheet, 99, 0, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(empty_result));
  EXPECT_TRUE(empty_result.value().valid);

  // A proposal in an adjacent column must not extend or affect A:A.
  Workbook adjacent = Workbook::create_empty();
  const std::size_t adjacent_index = adjacent.add_sheet("Adjacent");
  Sheet& adjacent_sheet = adjacent.sheet(adjacent_index);
  ASSERT_TRUE(static_cast<bool>(adjacent.set_cell_value(adjacent_index, 0, 0, Value::number(0.0))));
  DataValidation adjacent_rule;
  adjacent_rule.ranges.push_back(MergeRange{99, 1, 99, 1});  // B100
  adjacent_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  adjacent_rule.allow_blank = true;
  adjacent_rule.formula1 = "SUM($A:$A)>0";
  adjacent_sheet.mutable_validations().push_back(adjacent_rule);
  auto adjacent_result = validate_value(adjacent, adjacent_sheet, 99, 1, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(adjacent_result));
  EXPECT_FALSE(adjacent_result.value().valid);

  // Whole-row walking has the symmetric issue on the bounded row axis.
  Workbook row = Workbook::create_empty();
  const std::size_t row_index = row.add_sheet("Row");
  Sheet& row_sheet = row.sheet(row_index);
  ASSERT_TRUE(static_cast<bool>(row.set_cell_value(row_index, 0, 0, Value::number(0.0))));
  DataValidation row_rule;
  row_rule.ranges.push_back(MergeRange{0, 9, 0, 9});  // J1
  row_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  row_rule.allow_blank = true;
  row_rule.formula1 = "SUM($1:$1)>0";
  row_sheet.mutable_validations().push_back(row_rule);
  auto row_result = validate_value(row, row_sheet, 0, 9, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(row_result));
  EXPECT_TRUE(row_result.value().valid);
}

TEST(ValidationEval, CustomFormulaThreeDimensionalReadsOverlayOnTargetSheet) {
  Workbook wb = Workbook::create_empty();
  const std::size_t first_index = wb.add_sheet("First");
  const std::size_t second_index = wb.add_sheet("Second");
  Sheet& first = wb.sheet(first_index);
  Sheet& second = wb.sheet(second_index);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(first_index, 0, 0, Value::number(0.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(second_index, 0, 0, Value::number(1.0))));

  DataValidation whole;
  whole.ranges.push_back(MergeRange{99, 0, 99, 0});  // A100
  whole.type = static_cast<std::uint8_t>(Ty::kCustom);
  whole.allow_blank = true;
  whole.formula1 = "SUM(First:Second!A:A)>5";
  first.mutable_validations().push_back(whole);

  DataValidation bounded;
  bounded.ranges.push_back(MergeRange{100, 0, 100, 0});  // A101
  bounded.type = static_cast<std::uint8_t>(Ty::kCustom);
  bounded.allow_blank = true;
  bounded.formula1 = "SUM(First:Second!A1:A101)>5";
  first.mutable_validations().push_back(bounded);

  auto whole_result = validate_value(wb, first, 99, 0, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(whole_result));
  EXPECT_TRUE(whole_result.value().valid);
  whole_result = validate_value(wb, first, 99, 0, Value::number(4.0));
  ASSERT_TRUE(static_cast<bool>(whole_result));
  EXPECT_FALSE(whole_result.value().valid);

  auto bounded_result = validate_value(wb, first, 100, 0, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(bounded_result));
  EXPECT_TRUE(bounded_result.value().valid);
  bounded_result = validate_value(wb, first, 100, 0, Value::number(4.0));
  ASSERT_TRUE(static_cast<bool>(bounded_result));
  EXPECT_FALSE(bounded_result.value().valid);

  EXPECT_TRUE(first.resolve_cell_value(99, 0).is_blank());
  EXPECT_TRUE(first.resolve_cell_value(100, 0).is_blank());
  EXPECT_DOUBLE_EQ(second.resolve_cell_value(0, 0).as_number(), 1.0);
}

TEST(ValidationEval, CustomFormulaSeesAProposalInsideASpillBreakIt) {
  // Committing a value at the anchor replaces the spilling formula, and one at
  // a phantom blocks the spill, so `A1#` is #REF! either way. A blank at a
  // phantom leaves the spill intact.
  Workbook wb = Workbook::create_empty();
  const std::size_t sheet_index = wb.add_sheet("Spill");
  Sheet& sheet = wb.sheet(sheet_index);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));

  DataValidation anchor_rule;
  anchor_rule.ranges.push_back(MergeRange{0, 0, 0, 0});  // A1
  anchor_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  anchor_rule.allow_blank = false;
  anchor_rule.formula1 = "ISERROR(SUM(A1#))";
  sheet.mutable_validations().push_back(anchor_rule);

  DataValidation phantom_rule;
  phantom_rule.ranges.push_back(MergeRange{1, 0, 1, 0});  // A2, a spill phantom
  phantom_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  phantom_rule.allow_blank = false;
  phantom_rule.formula1 = "ISERROR(SUM($A$1#))";
  sheet.mutable_validations().push_back(phantom_rule);

  auto result = validate_value(wb, sheet, 0U, 0U, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().valid);
  result = validate_value(wb, sheet, 0U, 0U, Value::blank());
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().valid);

  result = validate_value(wb, sheet, 1U, 0U, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().valid);
  result = validate_value(wb, sheet, 1U, 0U, Value::blank());
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_FALSE(result.value().valid);

  // The committed spill and its geometry are untouched.
  EXPECT_DOUBLE_EQ(sheet.resolve_cell_value(0U, 0U).as_number(), 1.0);
  EXPECT_DOUBLE_EQ(sheet.resolve_cell_value(1U, 0U).as_number(), 2.0);
  const SpillRegion* region = sheet.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 2U);
  EXPECT_EQ(region->cols, 1U);
}

TEST(ValidationEval, CustomFormulaTrimRangeUsesProposedContentExtent) {
  // A non-blank proposal beyond the stored tail extends the content extent;
  // a blank proposal at the same coordinate leaves only the stored head.
  Workbook tail = Workbook::create_empty();
  const std::size_t tail_index = tail.add_sheet("Tail");
  Sheet& tail_sheet = tail.sheet(tail_index);
  ASSERT_TRUE(static_cast<bool>(tail.set_cell_value(tail_index, 0, 0, Value::number(1.0))));
  DataValidation tail_rule;
  tail_rule.ranges.push_back(MergeRange{99, 0, 99, 0});  // A100
  tail_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  tail_rule.allow_blank = false;
  tail_rule.formula1 = "ROWS(TRIMRANGE($A$1:$A$100))=100";
  tail_sheet.mutable_validations().push_back(tail_rule);
  auto result = validate_value(tail, tail_sheet, 99U, 0U, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().valid);
  result = validate_value(tail, tail_sheet, 99U, 0U, Value::blank());
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_FALSE(result.value().valid);

  // The same merge must extend the leading edge when the stored value is at
  // the tail and the proposal supplies the first cell.
  Workbook head = Workbook::create_empty();
  const std::size_t head_index = head.add_sheet("Head");
  Sheet& head_sheet = head.sheet(head_index);
  ASSERT_TRUE(static_cast<bool>(head.set_cell_value(head_index, 99, 0, Value::number(1.0))));
  DataValidation head_rule;
  head_rule.ranges.push_back(MergeRange{0, 0, 0, 0});  // A1
  head_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  head_rule.allow_blank = false;
  head_rule.formula1 = "ROWS(TRIMRANGE($A$1:$A$100))=100";
  head_sheet.mutable_validations().push_back(head_rule);
  result = validate_value(head, head_sheet, 0U, 0U, Value::number(5.0));
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().valid);
  result = validate_value(head, head_sheet, 0U, 0U, Value::blank());
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_FALSE(result.value().valid);

  // Blank has no content and yields TRIMRANGE's empty-reference error;
  // empty text is content and therefore yields a one-cell reference.
  Workbook empty = Workbook::create_empty();
  const std::size_t empty_index = empty.add_sheet("Empty");
  Sheet& empty_sheet = empty.sheet(empty_index);
  DataValidation empty_rule;
  empty_rule.ranges.push_back(MergeRange{49, 0, 49, 0});  // A50
  empty_rule.type = static_cast<std::uint8_t>(Ty::kCustom);
  empty_rule.allow_blank = false;
  empty_rule.formula1 = "ISERROR(TRIMRANGE($A$1:$A$100))";
  empty_sheet.mutable_validations().push_back(empty_rule);
  result = validate_value(empty, empty_sheet, 49U, 0U, Value::blank());
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_TRUE(result.value().valid);
  result = validate_value(empty, empty_sheet, 49U, 0U, Value::text(""));
  ASSERT_TRUE(static_cast<bool>(result));
  EXPECT_FALSE(result.value().valid);
  EXPECT_TRUE(empty_sheet.resolve_cell_value(49U, 0U).is_blank());
}

TEST(ValidationEval, QualifiedBuiltinCustomFormulaDoesNotPass) {
  Fixture f;
  ASSERT_TRUE(static_cast<bool>(
      f.wb.set_defined_name_scoped("SUM", "LAMBDA(x,x+100)", static_cast<std::int32_t>(f.sheet_index))));
  f.AddRule(Ty::kCustom, Op::kNone, true, "S!SUM(7)");
  EXPECT_FALSE(f.Check(Value::number(0.0)).valid);
}

TEST(ValidationEval, NoRuleAndOutcomeFields) {
  Fixture f;
  f.AddRule(Ty::kWhole, Op::kEqual, true, "3", "", 4, 4);
  f.sheet().mutable_validations()[0].error_style = 2;
  ValidationOutcome none = f.Check(Value::number(1.0), 0, 0);
  EXPECT_FALSE(none.has_rule);
  EXPECT_TRUE(none.valid);
  ValidationOutcome hit = f.Check(Value::number(1.0), 4, 4);
  EXPECT_TRUE(hit.has_rule);
  EXPECT_FALSE(hit.valid);
  EXPECT_EQ(hit.rule_index, 0U);
  EXPECT_EQ(hit.error_style, 2U);
  auto oob = validate_value(f.wb, f.sheet(), Sheet::kMaxRows, 0, Value::number(1.0));
  ASSERT_FALSE(static_cast<bool>(oob));
  EXPECT_EQ(oob.error().code, FormulonErrorCode::kInvalidArgument);
}

// Overlapping sqrefs: R1 whole 1..3 on A1:A3, R2 whole 10..20 on A1:A2. The rule
// that comes first in file order is the only one applied, whichever order they are declared in.
TEST(ValidationEval, OverlapPicksOneRule) {
  for (const bool r1_first : {true, false}) {
    Fixture f;
    DataValidation r1;
    r1.ranges.push_back(MergeRange{0, 0, 2, 0});
    r1.type = static_cast<std::uint8_t>(Ty::kWhole);
    r1.formula1 = "1";
    r1.formula2 = "3";
    DataValidation r2;
    r2.ranges.push_back(MergeRange{0, 0, 1, 0});
    r2.type = static_cast<std::uint8_t>(Ty::kWhole);
    r2.formula1 = "10";
    r2.formula2 = "20";
    auto& rules = f.sheet().mutable_validations();
    if (r1_first) {
      rules.push_back(r1);
      rules.push_back(r2);
    } else {
      rules.push_back(r2);
      rules.push_back(r1);
    }
    const std::uint32_t r1_index = r1_first ? 0U : 1U;
    const std::uint32_t r2_index = r1_first ? 1U : 0U;
    const std::uint32_t first = 0U;
    for (const double v : {2.0, 15.0, 99.0}) {
      const ValidationOutcome o = f.Check(Value::number(v), 0, 0);
      EXPECT_TRUE(o.has_rule);
      EXPECT_EQ(o.rule_index, first);
      const bool in_first = r1_first ? (v >= 1 && v <= 3) : (v >= 10 && v <= 20);
      EXPECT_EQ(o.valid, in_first) << "r1_first=" << r1_first << " v=" << v;
    }
    // A3 is covered by R1 alone.
    const ValidationOutcome a3 = f.Check(Value::number(2.0), 2, 0);
    EXPECT_EQ(a3.rule_index, r1_index);
    EXPECT_TRUE(a3.valid);
    (void)r2_index;
  }
}

std::vector<std::uint8_t> ReadFile(const char* file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + file;
  FILE* fp = std::fopen(path.c_str(), "rb");
  if (fp == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(fp, 0, SEEK_END);
  const long size = std::ftell(fp);
  std::fseek(fp, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), fp) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(fp);
  return out;
}

bool CheckStored(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Sheet& s = wb.sheet(0);
  auto r = validate_value(wb, s, row, col, s.resolve_cell_value(row, col));
  EXPECT_TRUE(static_cast<bool>(r));
  EXPECT_TRUE(r.value().has_rule) << "row " << row << " col " << col;
  return r.value().valid;
}

// The Excel-saved matrix workbook: every stored cell is validated against its own rule and
// compared with the result for that row.
TEST(ValidationEval, FixtureMatrix) {
  const std::vector<std::uint8_t> bytes = ReadFile("data_validation_matrix.xlsx");
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(read));
  const Workbook& wb = read.value().workbook;

  // Rows 1-80: type/operator/allowBlank matrix, inputs in columns A..F.
  std::uint32_t row = 0;
  for (const MatrixRow& m : kMatrix) {
    for (std::uint32_t c = 0; c < 6; ++c) {
      EXPECT_EQ(CheckStored(wb, row, c), m.expect[c] == '1') << "matrix row " << row + 1 << " col " << c;
    }
    ++row;
  }
  // Rows 81-122: list probes, one cell each.
  for (const ListRow& l : kListRows) {
    EXPECT_EQ(CheckStored(wb, row, 0), l.valid) << "list row " << row + 1;
    ++row;
  }
  // Rows 123-242: custom probes, one cell each.
  for (const CustomRow& c : kCustomRows) {
    EXPECT_EQ(CheckStored(wb, row, 0), c.valid) << "custom row " << row + 1;
    ++row;
  }
  // Rows 243-261: textLength on formatted values, cell-reference bounds, large whole bound.
  const bool tail[] = {true, false, true,  false, true,  true, true,  false, true, true,
                       true, false, false, true,  false, true, false, true,  true};
  for (const bool expect : tail) {
    EXPECT_EQ(CheckStored(wb, row, 0), expect) << "tail row " << row + 1;
    ++row;
  }
  EXPECT_EQ(row, 261U);
}

// Inline lists compare the displayed text of the proposed value.
bool InlineValid(const char* list, const Value& v, const char* format = nullptr) {
  Fixture f;
  f.wb.set_excel_profile(ExcelProfile{ExcelHost::kMac365, ExcelLocale::kJaJP});
  if (format != nullptr) {
    StylesTable& st = f.wb.mutable_styles();
    if (st.cell_xfs.empty()) {
      st.cell_xfs.emplace_back();
    }
    st.num_fmt_strings.emplace_back(format);
    st.num_fmts.push_back(NumFmtRecord{200, static_cast<std::uint32_t>(st.num_fmt_strings.size() - 1U)});
    CellXf xf;
    xf.num_fmt_id = 200;
    xf.apply_number_format = true;
    st.cell_xfs.push_back(xf);
    f.Store(Value::number(0.0));
    EXPECT_TRUE(static_cast<bool>(
        f.wb.set_cell_xf_index(f.sheet_index, 0, 0, static_cast<std::uint32_t>(st.cell_xfs.size() - 1U))));
  }
  f.AddRule(Ty::kList, Op::kNone, false, std::string("\"") + list + "\"");
  return f.Check(v).valid;
}

TEST(ValidationEvalInline, TextIsSpaceTrimmedAndOtherwiseExact) {
  const char* kList =
      "Yes,No,\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3,\xEF\xBD\xB1\xEF\xBD\xB2\xEF\xBD\xB3,\xC3\x84"
      "b";
  EXPECT_TRUE(InlineValid(kList, Text("Yes")));
  EXPECT_TRUE(InlineValid(kList, Text("Yes ")));
  EXPECT_TRUE(InlineValid(kList, Text(" Yes")));
  EXPECT_TRUE(InlineValid(kList, Text("Yes  ")));
  EXPECT_FALSE(InlineValid(kList, Text("yes")));
  EXPECT_FALSE(InlineValid(kList, Text("Yes\t")));
  EXPECT_FALSE(InlineValid(kList, Text("Yes\xC2\xA0")));
  EXPECT_FALSE(InlineValid(kList, Text("Yes\xE3\x80\x80")));
  EXPECT_FALSE(InlineValid(kList, Text("Y es")));
  EXPECT_TRUE(InlineValid(kList, Text("\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3")));
  EXPECT_FALSE(InlineValid(kList, Text("abc")));
  EXPECT_TRUE(InlineValid(kList, Text("\xEF\xBD\xB1\xEF\xBD\xB2\xEF\xBD\xB3")));
  EXPECT_FALSE(InlineValid(kList, Text("\xE3\x82\xA2\xE3\x82\xA4\xE3\x82\xA6")));
  EXPECT_TRUE(InlineValid(kList, Text("\xC3\x84"
                                      "b")));
  EXPECT_FALSE(InlineValid(kList, Text("\xC3\xA4"
                                       "b")));
}

TEST(ValidationEvalInline, ItemKeepsTrailingSpaceDropsLeading) {
  EXPECT_TRUE(InlineValid("Yes, No", Text("No")));
  EXPECT_TRUE(InlineValid("Yes, No", Text(" No")));
  EXPECT_FALSE(InlineValid("Yes ,No", Text("Yes")));
}

TEST(ValidationEvalInline, NumericTextCoerces) {
  for (const char* t : {"1", "\xEF\xBC\x91", "01", " 1", "1.0", "1e0"}) {
    EXPECT_TRUE(InlineValid("1,2", Text(t))) << t;
  }
  EXPECT_FALSE(InlineValid("1,2", Text("3")));
}

TEST(ValidationEvalInline, NumbersCompareThroughDisplay) {
  EXPECT_TRUE(InlineValid("1,2", Value::number(1.0)));
  EXPECT_TRUE(InlineValid("1,2", Value::number(1.0000000001)));
  EXPECT_FALSE(InlineValid("1,2", Value::number(1.0000000001), "0.0000000000"));
  EXPECT_TRUE(InlineValid("2,3", Value::number(1.5), "0"));
  EXPECT_FALSE(InlineValid("1.5,3", Value::number(1.5), "0"));
  EXPECT_TRUE(InlineValid("50%,3", Value::number(0.5), "0%"));
  EXPECT_TRUE(InlineValid("0.5,3", Value::number(0.5), "0%"));
  EXPECT_TRUE(InlineValid("abc,3", Value::number(1.0), "\"abc\""));
  EXPECT_FALSE(InlineValid("1,3", Value::number(1.0), "\"abc\""));
  EXPECT_TRUE(InlineValid("1234,3", Value::number(1234.0), "#,##0"));
}

TEST(ValidationEvalInline, BooleanAndBlank) {
  EXPECT_TRUE(InlineValid("TRUE,FALSE", Value::boolean(true)));
  EXPECT_FALSE(InlineValid("TRUE,FALSE", Text("true")));
  EXPECT_FALSE(InlineValid("Yes,No", Blank()));
}

}  // namespace
}  // namespace formulon
