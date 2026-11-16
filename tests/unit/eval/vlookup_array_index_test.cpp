//
// VLOOKUP / HLOOKUP with an array col_index_num / row_index_num: the lookup
// runs once and the column (row) pick lifts over the index array's shape,
// each element failing as the scalar call with that index would.

#include <cstdint>

#include "gtest/gtest.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSourceIn;

// A1 "k", B1 3, C1 7, E1 5.
Workbook MakeRowTable() {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(3.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(7.0));
  wb.sheet(0).set_cell_value(0, 4, Value::number(5.0));
  return wb;
}

// A1 "k", A2 3, A3 7.
Workbook MakeColumnTable() {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("k"));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(7.0));
  return wb;
}

void ExpectArrayShape(const Value& value, std::uint32_t rows, std::uint32_t cols) {
  ASSERT_TRUE(value.is_array()) << value.debug_to_string();
  EXPECT_EQ(value.as_array_rows(), rows);
  EXPECT_EQ(value.as_array_cols(), cols);
}

void ExpectNumber(const Value& value, double expected) {
  ASSERT_TRUE(value.is_number()) << value.debug_to_string();
  EXPECT_EQ(value.as_number(), expected);
}

void ExpectError(const Value& value, ErrorCode expected) {
  ASSERT_TRUE(value.is_error()) << value.debug_to_string();
  EXPECT_EQ(value.as_error(), expected);
}

TEST(VLookupArrayIndex, MedianReadsTheLiftedColumns) {
  Workbook wb = MakeRowTable();
  ExpectNumber(EvalSourceIn("=MEDIAN(E1,VLOOKUP(\"k\",A1:C1,{2,3},0))", wb, wb.sheet(0)), 5.0);
}

TEST(VLookupArrayIndex, ReferenceTableReturnsOneCellPerIndex) {
  Workbook wb = MakeRowTable();
  const Value v = EvalSourceIn("=VLOOKUP(\"k\",A1:C1,{2,3},0)", wb, wb.sheet(0));
  ExpectArrayShape(v, 1U, 2U);
  ExpectNumber(v.as_array_cells()[0], 3.0);
  ExpectNumber(v.as_array_cells()[1], 7.0);
}

TEST(VLookupArrayIndex, ArrayTableKeepsTheIndexShape) {
  Workbook wb = MakeRowTable();
  const Value v = EvalSourceIn("=VLOOKUP(\"k\",{\"k\",3,7},{3;2},0)", wb, wb.sheet(0));
  ExpectArrayShape(v, 2U, 1U);
  ExpectNumber(v.as_array_cells()[0], 7.0);
  ExpectNumber(v.as_array_cells()[1], 3.0);
}

TEST(VLookupArrayIndex, EachIndexFailsAsTheScalarCallWould) {
  Workbook wb = MakeRowTable();
  const Value v = EvalSourceIn("=VLOOKUP(\"k\",A1:C1,{2,4,0,#N/A},0)", wb, wb.sheet(0));
  ExpectArrayShape(v, 1U, 4U);
  const Value* cells = v.as_array_cells();
  ExpectNumber(cells[0], 3.0);
  ExpectError(cells[1], ErrorCode::Ref);
  ExpectError(cells[2], ErrorCode::Value);
  ExpectError(cells[3], ErrorCode::NA);
}

TEST(VLookupArrayIndex, ScalarIndexStaysScalar) {
  Workbook wb = MakeRowTable();
  ExpectNumber(EvalSourceIn("=VLOOKUP(\"k\",A1:C1,3,0)", wb, wb.sheet(0)), 7.0);
  ExpectError(EvalSourceIn("=VLOOKUP(\"k\",A1:C1,4,0)", wb, wb.sheet(0)), ErrorCode::Ref);
}

TEST(HLookupArrayIndex, ReferenceTableReturnsOneCellPerIndex) {
  Workbook wb = MakeColumnTable();
  const Value v = EvalSourceIn("=HLOOKUP(\"k\",A1:A3,{2,3},0)", wb, wb.sheet(0));
  ExpectArrayShape(v, 1U, 2U);
  ExpectNumber(v.as_array_cells()[0], 3.0);
  ExpectNumber(v.as_array_cells()[1], 7.0);
}

TEST(HLookupArrayIndex, EachIndexFailsAsTheScalarCallWould) {
  Workbook wb = MakeColumnTable();
  const Value v = EvalSourceIn("=HLOOKUP(\"k\",A1:A3,{3,4,0.5},0)", wb, wb.sheet(0));
  ExpectArrayShape(v, 1U, 3U);
  ExpectNumber(v.as_array_cells()[0], 7.0);
  ExpectError(v.as_array_cells()[1], ErrorCode::Ref);
  ExpectError(v.as_array_cells()[2], ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
