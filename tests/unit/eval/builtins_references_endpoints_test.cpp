// Reference builtin tests grouped by reference constructor.

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>

#include "builtins_references_test_helpers.h"

namespace formulon {
namespace eval {
namespace {
using namespace references_test_helpers;

Workbook NamedEndpointBook() {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  for (std::uint32_t r = 0; r < 4U; ++r) {
    for (std::uint32_t c = 0; c < 4U; ++c) {
      const double v = static_cast<double>((r + 1U) * 10U + (c + 1U));
      wb.sheet(0).set_cell_value(r, c, Value::number(v));
      wb.sheet(1).set_cell_value(r, c, Value::number(100.0 + v));
    }
  }
  const std::pair<const char*, const char*> names[] = {
      {"CellNm", "=Sheet1!$C$3"},
      {"RectNm", "=Sheet1!$C$3:$D$4"},
      {"Konst", "=5"},
      {"IdxNm", "=INDEX(Sheet1!$A$1:$D$4,3,3)"},
      {"OtherNm", "=Sheet2!$C$3"},
      {"ExprNm", "=Sheet1!$C$3+0"},
      {"TextNm", "=\"abc\""},
      {"OffNm", "=OFFSET(Sheet1!$A$1,2,2)"},
      {"ErrNm", "=#N/A"},
      {"NaNm", "=NA()"},
  };
  for (const auto& [name, formula] : names) {
    EXPECT_TRUE(static_cast<bool>(wb.set_defined_name(name, formula)));
  }
  return wb;
}

void ExpectError(const Value& v, ErrorCode expected, std::string_view formula) {
  ASSERT_TRUE(v.is_error()) << formula << " -> " << v.debug_to_string();
  EXPECT_EQ(v.as_error(), expected) << formula;
}

TEST(ReferenceCall, CallEndpointRangeSpillsAsAValue) {
  Workbook wb = ColumnOfTen();
  const Value v = EvalSourceIn("=A1:INDEX(A1:A10,3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 3.0);
  ExpectNumber(EvalSourceIn("=A2:OFFSET(A1,1,0)", wb, wb.sheet(0)), 2.0, "1x1 call-endpoint range");
}

TEST(ReferenceCall, IndexAsReferenceArgument) {
  Workbook wb = ColumnOfTen();
  const Value isref = EvalSourceIn("=ISREF(INDEX(A1:B2,1,1))", wb, wb.sheet(0));
  ASSERT_TRUE(isref.is_boolean());
  EXPECT_TRUE(isref.as_boolean());
  const Value literal = EvalSourceIn("=ISREF(INDEX({1,2},1))", wb, wb.sheet(0));
  ASSERT_TRUE(literal.is_boolean());
  EXPECT_FALSE(literal.as_boolean());
  ExpectNumber(EvalSourceIn("=ROW(INDEX(A1:A10,3))", wb, wb.sheet(0)), 3.0, "ROW(INDEX(A1:A10,3))");
  ExpectNumber(EvalSourceIn("=COLUMN(INDEX(A1:C3,2,3))", wb, wb.sheet(0)), 3.0, "COLUMN(INDEX(A1:C3,2,3))");
  ExpectText(EvalSourceIn("=CELL(\"address\",INDEX(A1:C3,2,3))", wb, wb.sheet(0)), "$C$2", "CELL of INDEX");
  // A zero index selects a whole column / row of the source.
  ExpectNumber(EvalSourceIn("=ROWS(INDEX(A1:C3,0,2))", wb, wb.sheet(0)), 3.0, "ROWS(INDEX(A1:C3,0,2))");
  ExpectNumber(EvalSourceIn("=COLUMNS(INDEX(A1:C3,2,0))", wb, wb.sheet(0)), 3.0, "COLUMNS(INDEX(A1:C3,2,0))");
  ExpectNumber(EvalSourceIn("=SUM(OFFSET(INDEX(A1:A10,2),1,0,2,1))", wb, wb.sheet(0)), 7.0, "OFFSET over INDEX");
}

TEST(ReferenceCall, IndexAreaNumSelectsFromUnion) {
  // A1:B2 = {1,2;3,4}, D1:E2 = {10,20;30,40}.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(0, 3, Value::number(10.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(20.0));
  wb.sheet(0).set_cell_value(1, 3, Value::number(30.0));
  wb.sheet(0).set_cell_value(1, 4, Value::number(40.0));
  ExpectNumber(EvalSourceIn("=INDEX((A1:B2,D1:E2),2,1,2)", wb, wb.sheet(0)), 30.0, "area 2");
  ExpectNumber(EvalSourceIn("=INDEX((A1:B2,D1:E2),2,1,1)", wb, wb.sheet(0)), 3.0, "area 1");
  ExpectNumber(EvalSourceIn("=INDEX((A1:B2,D1:E2),2,1)", wb, wb.sheet(0)), 3.0, "default area");
  ExpectNumber(EvalSourceIn("=INDEX(A1:B2,1,2,1)", wb, wb.sheet(0)), 2.0, "single area");
  ExpectText(EvalSourceIn("=CELL(\"address\",INDEX((A1:B2,D1:E2),1,2,2))", wb, wb.sheet(0)), "$E$1", "CELL of area 2");
  ExpectNumber(EvalSourceIn("=SUM(INDEX((A1:B2,D1:E2),0,2,2))", wb, wb.sheet(0)), 60.0, "whole column of area 2");
  ExpectNumber(EvalSourceIn("=AREAS(INDEX((A1:B2,D1:E2),1,1,2))", wb, wb.sheet(0)), 1.0, "AREAS of INDEX");
  for (const char* formula : {"=INDEX((A1:B2,D1:E2),1,1,3)", "=INDEX(A1:B2,1,1,2)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Ref) << formula;
  }
  for (const char* formula :
       {"=INDEX((A1:B2,D1:E2),1,1,0)", "=INDEX(A1:B2,1,1,0)", "=INDEX((A1:B2,D1:E2),1,1,-1)", "=INDEX(A1:B2,1,1,-1)"}) {
    const Value v = EvalSourceIn(formula, wb, wb.sheet(0));
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }
  // A fractional area_num truncates toward zero, like row_num / column_num.
  ExpectNumber(EvalSourceIn("=INDEX((A1:B2,D1:E2),2,1,1.9)", wb, wb.sheet(0)), 3.0, "area_num truncates");
  // area_num on a whole-area selection (row=0, col=0) picks that area's
  // own shape and top-left address, not the union's.
  ExpectNumber(EvalSourceIn("=ROWS(INDEX((A1:B2,D1:E2),0,0,2))", wb, wb.sheet(0)), 2.0, "ROWS of whole area 2");
  ExpectNumber(EvalSourceIn("=SUM(INDEX((A1:B2,D1:E2),0,0,2))", wb, wb.sheet(0)), 100.0, "SUM of whole area 2");
  ExpectText(EvalSourceIn("=CELL(\"address\",INDEX((A1:B2,D1:E2),0,0,2))", wb, wb.sheet(0)), "$D$1",
             "CELL of whole area 2");
}

TEST(ReferenceCall, XlookupAsReference) {
  // A1:A3 = a/b/c, B1:C3 = {1,10;2,20;3,30}.
  Workbook wb = Workbook::create();
  const char* keys[] = {"a", "b", "c"};
  for (std::uint32_t r = 0; r < 3; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::text(keys[r]));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 2, Value::number(static_cast<double>((r + 1) * 10)));
  }
  ExpectNumber(EvalSourceIn("=SUM(XLOOKUP(\"b\",A1:A3,B1:B3):B3)", wb, wb.sheet(0)), 5.0, "XLOOKUP endpoint");
  const Value isref = EvalSourceIn("=ISREF(XLOOKUP(\"b\",A1:A3,B1:B3))", wb, wb.sheet(0));
  ASSERT_TRUE(isref.is_boolean());
  EXPECT_TRUE(isref.as_boolean());
  ExpectText(EvalSourceIn("=CELL(\"address\",XLOOKUP(\"c\",A1:A3,B1:B3))", wb, wb.sheet(0)), "$B$3", "CELL of XLOOKUP");
  // A multi-column return_array yields the matched row.
  ExpectNumber(EvalSourceIn("=COLUMNS(XLOOKUP(\"a\",A1:A3,B1:C3))", wb, wb.sheet(0)), 2.0, "matched row");
  ExpectNumber(EvalSourceIn("=ROW(XLOOKUP(\"b\",A1:A3,B1:C3))", wb, wb.sheet(0)), 2.0, "ROW of matched row");
  // A miss returns if_not_found, which may itself be a reference.
  ExpectText(EvalSourceIn("=CELL(\"address\",XLOOKUP(\"z\",A1:A3,B1:B3,C1))", wb, wb.sheet(0)), "$C$1",
             "miss falls back to reference");
  const Value miss = EvalSourceIn("=SUM(A1:XLOOKUP(\"z\",A1:A3,B1:B3))", wb, wb.sheet(0));
  ASSERT_TRUE(miss.is_error());
  EXPECT_EQ(miss.as_error(), ErrorCode::NA);
}

TEST(ReferenceCall, IfsAndSwitchPassReferencesThrough) {
  Workbook wb = ColumnOfTen();
  ExpectNumber(EvalSourceIn("=ROWS(A1:SWITCH(2,1,A2,2,A4))", wb, wb.sheet(0)), 4.0, "SWITCH endpoint");
  ExpectNumber(EvalSourceIn("=SUM(A1:IFS(FALSE,A2,TRUE,A3))", wb, wb.sheet(0)), 6.0, "IFS endpoint");
  const Value isref = EvalSourceIn("=ISREF(IFS(FALSE,A1,TRUE,B2))", wb, wb.sheet(0));
  ASSERT_TRUE(isref.is_boolean());
  EXPECT_TRUE(isref.as_boolean());
  const Value scalar = EvalSourceIn("=ISREF(SWITCH(1,1,5))", wb, wb.sheet(0));
  ASSERT_TRUE(scalar.is_boolean());
  EXPECT_FALSE(scalar.as_boolean());
  ExpectNumber(EvalSourceIn("=AREAS(SWITCH(9,1,A1,B1:C2))", wb, wb.sheet(0)), 1.0, "AREAS of SWITCH default");
  // The value path is unchanged.
  ExpectNumber(EvalSourceIn("=SWITCH(2,1,A2,2,A4)", wb, wb.sheet(0)), 4.0, "SWITCH value");
  ExpectNumber(EvalSourceIn("=IFS(FALSE,A2,TRUE,A3)", wb, wb.sheet(0)), 3.0, "IFS value");
}

TEST(ReferenceCall, DefinedNameAsRangeEndpoint) {
  Workbook wb = NamedEndpointBook();
  const Sheet& s1 = wb.sheet(0);
  ExpectNumber(EvalSourceIn("=SUM(A1:CellNm)", wb, s1), 198.0, "SUM(A1:CellNm)");
  ExpectNumber(EvalSourceIn("=SUM(CellNm:A1)", wb, s1), 198.0, "SUM(CellNm:A1)");
  ExpectNumber(EvalSourceIn("=SUM(A1:RectNm)", wb, s1), 440.0, "SUM(A1:RectNm)");
  ExpectNumber(EvalSourceIn("=SUM(A1:IdxNm)", wb, s1), 198.0, "SUM(A1:IdxNm)");
  ExpectNumber(EvalSourceIn("=SUM(A1:OffNm)", wb, s1), 198.0, "SUM(A1:OffNm)");
  ExpectNumber(EvalSourceIn("=ROWS(A1:RectNm)", wb, s1), 4.0, "ROWS(A1:RectNm)");
  ExpectNumber(EvalSourceIn("=SUM(CellNm:RectNm)", wb, s1), 154.0, "SUM(CellNm:RectNm)");
  ExpectNumber(EvalSourceIn("=SUM(A1:CellNm:B1)", wb, s1), 198.0, "SUM(A1:CellNm:B1)");
  ExpectNumber(EvalSourceIn("=SUM(Sheet1!A1:CellNm)", wb, s1), 198.0, "SUM(Sheet1!A1:CellNm)");
  ExpectNumber(EvalSourceIn("=SUM(Sheet2!A1:OtherNm)", wb, s1), 1098.0, "SUM(Sheet2!A1:OtherNm)");
  ExpectNumber(EvalSourceIn("=SUM(A1:Sheet1!CellNm)", wb, s1), 198.0, "SUM(A1:Sheet1!CellNm)");
  ExpectNumber(EvalSourceIn("=SUM(A1:[0]!CellNm)", wb, s1), 198.0, "SUM(A1:[0]!CellNm)");
}

TEST(ReferenceCall, NonReferenceNameAsRangeEndpoint) {
  Workbook wb = NamedEndpointBook();
  const Sheet& s1 = wb.sheet(0);
  ExpectError(EvalSourceIn("=SUM(A1:Konst)", wb, s1), ErrorCode::Value, "SUM(A1:Konst)");
  ExpectError(EvalSourceIn("=SUM(Konst:A1)", wb, s1), ErrorCode::Value, "SUM(Konst:A1)");
  ExpectError(EvalSourceIn("=ROWS(A1:Konst)", wb, s1), ErrorCode::Value, "ROWS(A1:Konst)");
  ExpectError(EvalSourceIn("=SUM(A1:ExprNm)", wb, s1), ErrorCode::Value, "SUM(A1:ExprNm)");
  ExpectError(EvalSourceIn("=SUM(A1:TextNm)", wb, s1), ErrorCode::Value, "SUM(A1:TextNm)");
  ExpectError(EvalSourceIn("=SUM(A1:Nope)", wb, s1), ErrorCode::Name, "SUM(A1:Nope)");
  ExpectError(EvalSourceIn("=SUM(A1:ErrNm)", wb, s1), ErrorCode::NA, "SUM(A1:ErrNm)");
  ExpectError(EvalSourceIn("=SUM(A1:NaNm)", wb, s1), ErrorCode::NA, "SUM(A1:NaNm)");
  ExpectError(EvalSourceIn("=LET(r,5,SUM(A1:r))", wb, s1), ErrorCode::Value, "LET(r,5,SUM(A1:r))");
}

TEST(ReferenceCall, RangeEndpointsOnDifferentSheets) {
  Workbook wb = NamedEndpointBook();
  const Sheet& s1 = wb.sheet(0);
  const Sheet& s2 = wb.sheet(1);
  ExpectError(EvalSourceIn("=SUM(A1:OtherNm)", wb, s1), ErrorCode::Value, "Sheet1: SUM(A1:OtherNm)");
  ExpectError(EvalSourceIn("=SUM(Sheet1!A1:OtherNm)", wb, s1), ErrorCode::Value, "SUM(Sheet1!A1:OtherNm)");
  ExpectError(EvalSourceIn("=SUM(A1:INDEX(Sheet2!A1:D4,3,3))", wb, s1), ErrorCode::Value,
              "Sheet1: SUM(A1:INDEX(Sheet2!A1:D4,3,3))");
  ExpectError(EvalSourceIn("=SUM(Sheet1!A1:INDEX(Sheet2!A1:D4,3,3))", wb, s1), ErrorCode::Value,
              "SUM(Sheet1!A1:INDEX(Sheet2!A1:D4,3,3))");
  ExpectNumber(EvalSourceIn("=SUM(A1:OtherNm)", wb, s2), 1098.0, "Sheet2: SUM(A1:OtherNm)");
  ExpectNumber(EvalSourceIn("=SUM(A1:INDEX(Sheet2!A1:D4,3,3))", wb, s2), 1098.0,
               "Sheet2: SUM(A1:INDEX(Sheet2!A1:D4,3,3))");
  ExpectError(EvalSourceIn("=SUM(A1:CellNm)", wb, s2), ErrorCode::Value, "Sheet2: SUM(A1:CellNm)");
  ExpectError(EvalSourceIn("=SUM(A1:INDEX(Sheet1!A1:D4,3,3))", wb, s2), ErrorCode::Value,
              "Sheet2: SUM(A1:INDEX(Sheet1!A1:D4,3,3))");
}

TEST(RangeEndpoint, RefToOffsetSingleCol) {
  // SUM(A1:OFFSET(A1,4,0)) — lhs literal Ref, rhs OFFSET single-cell.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUM(A1:OFFSET(A1,4,0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(RangeEndpoint, OffsetToOffsetSingleCol) {
  // SUM(OFFSET(A1,0,0):OFFSET(A1,4,0)) — both endpoints are OFFSET calls.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUM(OFFSET(A1,0,0):OFFSET(A1,4,0))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(RangeEndpoint, IndirectToIndirect) {
  // SUM(INDIRECT("B2"):INDIRECT("B6")) — both endpoints INDIRECT.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(4.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(6.0));
  wb.sheet(0).set_cell_value(4, 1, Value::number(8.0));
  wb.sheet(0).set_cell_value(5, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=SUM(INDIRECT(\"B2\"):INDIRECT(\"B6\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(RangeEndpoint, OffsetAboveRefNormalises) {
  // SUM(OFFSET(A5,-4,0):A5) — lhs OFFSET resolves to A1, rhs A5;
  // endpoint ordering is normalised so the union is A1:A5.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  const Value v = EvalSourceIn("=SUM(OFFSET(A5,-4,0):A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(RangeEndpoint, RefToMultiCellOffset) {
  // SUM(A1:OFFSET(A1,2,1,2,1)) — rhs OFFSET resolves to a 2x1
  // rectangle B3:B4. The rectangle's bottom-right corner participates
  // in the union so the unioned range is A1:B4 =
  //   1 + 2 + 3 + 4 + 10 + 20 + 30 + 40 = 110.
  // Verified against Mac Excel 365 in
  // tests/oracle/targets/mac-365-ja_JP/golden/range_endpoint_calls.golden.json.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>((r + 1) * 10)));
  }
  const Value v = EvalSourceIn("=SUM(A1:OFFSET(A1,2,1,2,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 110.0);
}

TEST(RangeEndpoint, RefToOffsetTwoCol) {
  // SUM(A1:OFFSET(A1,4,1)) — rhs OFFSET steps to column B; spans A1:B5.
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
    wb.sheet(0).set_cell_value(r, 1, Value::number(static_cast<double>((r + 1) * 10)));
  }
  const Value v = EvalSourceIn("=SUM(A1:OFFSET(A1,4,1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 165.0);
}

TEST(RangeEndpoint, RefToIndirect) {
  // SUM(A1:INDIRECT("A3")) — lhs Ref, rhs INDIRECT.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(300.0));
  const Value v = EvalSourceIn("=SUM(A1:INDIRECT(\"A3\"))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 600.0);
}

TEST(ReferenceCall, IndexAsRangeEndpoint) {
  Workbook wb = ColumnOfTen();
  ExpectNumber(EvalSourceIn("=SUM(A1:INDEX(A1:A10,3))", wb, wb.sheet(0)), 6.0, "SUM(A1:INDEX(A1:A10,3))");
  ExpectNumber(EvalSourceIn("=SUM(A1:INDEX(A:A,COUNTA(A:A)))", wb, wb.sheet(0)), 55.0, "dynamic A1:INDEX(A:A,n)");
  ExpectNumber(EvalSourceIn("=ROWS(INDEX(A:A,2):INDEX(A:A,5))", wb, wb.sheet(0)), 4.0, "ROWS(INDEX:INDEX)");
  ExpectNumber(EvalSourceIn("=SUM(INDEX(A1:A10,4):INDEX(A1:A10,6))", wb, wb.sheet(0)), 15.0, "SUM(INDEX:INDEX)");
  const Value out_of_range = EvalSourceIn("=SUM(A1:INDEX(A1:A10,11))", wb, wb.sheet(0));
  ASSERT_TRUE(out_of_range.is_error());
  EXPECT_EQ(out_of_range.as_error(), ErrorCode::Ref);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
