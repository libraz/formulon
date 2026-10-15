// Mac Excel 365 non-Sum pivot aggregations: per-step snap in a leaf, one accumulator over the covered records
// for a one-axis total, a raw combination of the row partials for a total across both axes. Expected values
// were read from Excel after refreshing a pivot over the same records, with items named in display order.

#include <cstddef>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_evaluator_fixtures.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_sum_amount_table;
using test::owned_text;

struct Cell {
  enum class Kind { Blank, Number, Text, Boolean, Error };
  Kind kind;
  double value;
  const char* str;

  static Cell blank() { return {Kind::Blank, 0.0, ""}; }
  static Cell number(double v) { return {Kind::Number, v, ""}; }
  static Cell text(const char* v) { return {Kind::Text, 0.0, v}; }
  static Cell boolean(bool v) { return {Kind::Boolean, v ? 1.0 : 0.0, ""}; }
  static Cell na() { return {Kind::Error, 0.0, ""}; }
};

struct Record {
  const char* region;
  const char* product;
  Cell value;
};

struct Exp {
  enum class Kind { Blank, Number, Error };
  Kind kind;
  double value;
  ErrorCode code;

  static Exp blank() { return {Kind::Blank, 0.0, ErrorCode::NA}; }
  static Exp number(double v) { return {Kind::Number, v, ErrorCode::NA}; }
  static Exp error(ErrorCode e) { return {Kind::Error, 0.0, e}; }
};

/// One rendered cell by row / column label; "GT" is the grand total on that axis, an empty column the only one.
struct Expected {
  const char* row;
  const char* col;
  Exp exp;
};

PivotCache build_cache(const std::vector<Record>& records) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (const Record& r : records) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, r.region));
    rec.cells.push_back(owned_text(cache, r.product));
    switch (r.value.kind) {
      case Cell::Kind::Blank:
        rec.cells.push_back(Value::blank());
        break;
      case Cell::Kind::Number:
        rec.cells.push_back(Value::number(r.value.value));
        break;
      case Cell::Kind::Text:
        rec.cells.push_back(owned_text(cache, r.value.str));
        break;
      case Cell::Kind::Boolean:
        rec.cells.push_back(Value::boolean(r.value.value != 0.0));
        break;
      case Cell::Kind::Error:
        rec.cells.push_back(Value::error(ErrorCode::NA));
        break;
    }
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

std::size_t label_index(const std::vector<AxisHierarchyNode>& nodes, const std::string& label) {
  for (std::size_t i = 0; i < nodes.size(); ++i) {
    if (nodes[i].label == label) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

const Value* find_cell(const PivotResult& r, bool has_cols, const Expected& e) {
  const std::string row = e.row;
  const std::string col = e.col;
  if (row == "GT") {
    if (!has_cols || col == "GT") {
      return &r.grand_total;
    }
    const std::size_t c = label_index(r.cols, col);
    return c < r.col_leaf_totals.size() ? &r.col_leaf_totals[c][0] : nullptr;
  }
  const std::size_t ri = label_index(r.rows, row);
  if (ri >= r.values.size()) {
    return nullptr;
  }
  if (!has_cols) {
    return &r.values[ri][0][0];
  }
  if (col == "GT") {
    return &r.row_leaf_totals[ri][0];
  }
  const std::size_t c = label_index(r.cols, col);
  return c < r.values[ri].size() ? &r.values[ri][c][0] : nullptr;
}

void check_case(const std::vector<Record>& records, Aggregation aggregation, bool has_cols,
                const std::vector<Expected>& expected) {
  const PivotCache cache = build_cache(records);
  PivotTable table =
      build_sum_amount_table({0}, has_cols ? std::vector<std::uint32_t>{1} : std::vector<std::uint32_t>{});
  table.mutable_data_fields()[0].aggregation = aggregation;
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or));
  const PivotResult& result = result_or.value();
  for (const Expected& e : expected) {
    SCOPED_TRACE(std::string(e.row) + " x " + e.col);
    const Value* cell = find_cell(result, has_cols, e);
    ASSERT_NE(cell, nullptr);
    switch (e.exp.kind) {
      case Exp::Kind::Blank:
        EXPECT_TRUE(cell->is_blank());
        break;
      case Exp::Kind::Number:
        ASSERT_TRUE(cell->is_number());
        EXPECT_EQ(cell->as_number(), e.exp.value);
        break;
      case Exp::Kind::Error:
        ASSERT_TRUE(cell->is_error());
        EXPECT_EQ(cell->as_error(), e.exp.code);
        break;
    }
  }
}

// text and booleans are non-numeric; all-blank group is blank; an error poisons its group and the grand total.
TEST(PivotAggregationTotals, KindsAverage) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(3.0)},
      {"R2", "", Exp::error(ErrorCode::Div0)},
      {"R3", "", Exp::blank()},
      {"R4", "", Exp::error(ErrorCode::Div0)},
      {"R5", "", Exp::error(ErrorCode::NA)},
      {"R6", "", Exp::number(3.5)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::Average, false, expected);
}

// Count counts text, booleans and errors; an all-blank group is blank.
TEST(PivotAggregationTotals, KindsCount) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(3.0)},  {"R2", "", Exp::number(2.0)}, {"R3", "", Exp::blank()},
      {"R4", "", Exp::number(2.0)},  {"R5", "", Exp::number(2.0)}, {"R6", "", Exp::number(2.0)},
      {"GT", "", Exp::number(11.0)},
  };
  check_case(records, Aggregation::Count, false, expected);
}

// CountNums counts numbers only; booleans and errors do not count.
TEST(PivotAggregationTotals, KindsCountNums) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(1.0)}, {"R2", "", Exp::number(0.0)}, {"R3", "", Exp::blank()},
      {"R4", "", Exp::number(0.0)}, {"R5", "", Exp::number(1.0)}, {"R6", "", Exp::number(2.0)},
      {"GT", "", Exp::number(4.0)},
  };
  check_case(records, Aggregation::CountNumbers, false, expected);
}

// Max over a non-numeric non-blank group is 0; an error propagates.
TEST(PivotAggregationTotals, KindsMax) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(3.0)},          {"R2", "", Exp::number(0.0)},          {"R3", "", Exp::blank()},
      {"R4", "", Exp::number(0.0)},          {"R5", "", Exp::error(ErrorCode::NA)}, {"R6", "", Exp::number(5.0)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::Max, false, expected);
}

// Min over a non-numeric non-blank group is 0; an error propagates.
TEST(PivotAggregationTotals, KindsMin) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(3.0)},          {"R2", "", Exp::number(0.0)},          {"R3", "", Exp::blank()},
      {"R4", "", Exp::number(0.0)},          {"R5", "", Exp::error(ErrorCode::NA)}, {"R6", "", Exp::number(2.0)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::Min, false, expected);
}

// Product over a non-numeric non-blank group is 0.
TEST(PivotAggregationTotals, KindsProduct) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(3.0)},          {"R2", "", Exp::number(0.0)},          {"R3", "", Exp::blank()},
      {"R4", "", Exp::number(0.0)},          {"R5", "", Exp::error(ErrorCode::NA)}, {"R6", "", Exp::number(10.0)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::Product, false, expected);
}

// StDev needs two numbers; one number is #DIV/0!.
TEST(PivotAggregationTotals, KindsStdev) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::error(ErrorCode::Div0)},
      {"R2", "", Exp::error(ErrorCode::Div0)},
      {"R3", "", Exp::blank()},
      {"R4", "", Exp::error(ErrorCode::Div0)},
      {"R5", "", Exp::error(ErrorCode::NA)},
      {"R6", "", Exp::number(2.1213203435596424)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::StdDev, false, expected);
}

// StDevP of one number is 0, of none #DIV/0!.
TEST(PivotAggregationTotals, KindsStdevp) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(0.0)},
      {"R2", "", Exp::error(ErrorCode::Div0)},
      {"R3", "", Exp::blank()},
      {"R4", "", Exp::error(ErrorCode::Div0)},
      {"R5", "", Exp::error(ErrorCode::NA)},
      {"R6", "", Exp::number(1.5)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::StdDevP, false, expected);
}

// Var needs two numbers; one number is #DIV/0!.
TEST(PivotAggregationTotals, KindsVar) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::error(ErrorCode::Div0)}, {"R2", "", Exp::error(ErrorCode::Div0)}, {"R3", "", Exp::blank()},
      {"R4", "", Exp::error(ErrorCode::Div0)}, {"R5", "", Exp::error(ErrorCode::NA)},   {"R6", "", Exp::number(4.5)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::Var, false, expected);
}

// VarP of one number is 0, of none #DIV/0!.
TEST(PivotAggregationTotals, KindsVarp) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(3.0)},   {"R1", "", Cell::text("x")},     {"R1", "", Cell::blank()},
      {"R1", "", Cell::boolean(true)}, {"R2", "", Cell::text("x")},     {"R2", "", Cell::text("y")},
      {"R3", "", Cell::blank()},       {"R4", "", Cell::boolean(true)}, {"R4", "", Cell::boolean(false)},
      {"R5", "", Cell::na()},          {"R5", "", Cell::number(7.0)},   {"R6", "", Cell::number(2.0)},
      {"R6", "", Cell::number(5.0)},   {"R6", "", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(0.0)},
      {"R2", "", Exp::error(ErrorCode::Div0)},
      {"R3", "", Exp::blank()},
      {"R4", "", Exp::error(ErrorCode::Div0)},
      {"R5", "", Exp::error(ErrorCode::NA)},
      {"R6", "", Exp::number(2.25)},
      {"GT", "", Exp::error(ErrorCode::NA)},
  };
  check_case(records, Aggregation::VarP, false, expected);
}

// Average numerator snaps per step, not once at the end.
TEST(PivotAggregationTotals, AverageStepSnap) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(0.5)},    {"R1", "", Cell::number(-0.4)},
      {"R1", "", Cell::number(-0.1)},   {"R2", "", Cell::number(9007199254740992.0)},
      {"R2", "", Cell::number(12.0)},   {"R2", "", Cell::number(-9007199254740992.0)},
      {"R2", "", Cell::number(100.0)},  {"R3", "", Cell::number(1e+16)},
      {"R3", "", Cell::number(2.0)},    {"R3", "", Cell::number(-1e+16)},
      {"R4", "", Cell::number(1e+16)},  {"R4", "", Cell::number(2.0)},
      {"R4", "", Cell::number(-1e+16)}, {"R4", "", Cell::number(0.5)},
      {"R5", "", Cell::number(1.0)},    {"R5", "", Cell::number(2.0)},
      {"R5", "", Cell::number(4.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(0.0)},
      {"R2", "", Exp::number(25.0)},
      {"R3", "", Exp::number(0.0)},
      {"R4", "", Exp::number(0.125)},
      {"R5", "", Exp::number(2.3333333333333335)},
      {"GT", "", Exp::number(6.5588235294117645)},
  };
  check_case(records, Aggregation::Average, false, expected);
}

// Var is one-pass sum of squares: near-equal values give 0.
TEST(PivotAggregationTotals, VarOnePass) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(100000000.1)}, {"R1", "", Cell::number(100000000.2)},
      {"R1", "", Cell::number(100000000.3)}, {"R1", "", Cell::number(100000000.4)},
      {"R2", "", Cell::number(1.0)},         {"R2", "", Cell::number(2.0)},
      {"R2", "", Cell::number(4.0)},         {"R3", "", Cell::number(0.1)},
      {"R3", "", Cell::number(0.2)},         {"R3", "", Cell::number(0.3)},
      {"R4", "", Cell::number(1e-160)},      {"R4", "", Cell::number(1e-160)},
      {"R4", "", Cell::number(1e+200)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(0.0)},
      {"R2", "", Exp::number(2.333333333333334)},
      {"R3", "", Exp::number(0.009999999999999988)},
      {"R4", "", Exp::error(ErrorCode::Num)},
      {"GT", "", Exp::error(ErrorCode::Num)},
  };
  check_case(records, Aggregation::Var, false, expected);
}

// Min of a group whose sum of squares overflows is still the minimum.
TEST(PivotAggregationTotals, MinOverflowGroup) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(100000000.1)}, {"R1", "", Cell::number(100000000.2)},
      {"R1", "", Cell::number(100000000.3)}, {"R1", "", Cell::number(100000000.4)},
      {"R2", "", Cell::number(1.0)},         {"R2", "", Cell::number(2.0)},
      {"R2", "", Cell::number(4.0)},         {"R3", "", Cell::number(0.1)},
      {"R3", "", Cell::number(0.2)},         {"R3", "", Cell::number(0.3)},
      {"R4", "", Cell::number(1e-160)},      {"R4", "", Cell::number(1e-160)},
      {"R4", "", Cell::number(1e+200)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(100000000.1)}, {"R2", "", Exp::number(1.0)}, {"R3", "", Exp::number(0.1)},
      {"R4", "", Exp::number(0.0)},         {"GT", "", Exp::number(0.0)},
  };
  check_case(records, Aggregation::Min, false, expected);
}

// Product total is a display-order chain.
TEST(PivotAggregationTotals, ProductNoSubnormalStep) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(100000000.1)}, {"R1", "", Cell::number(100000000.2)},
      {"R1", "", Cell::number(100000000.3)}, {"R1", "", Cell::number(100000000.4)},
      {"R2", "", Cell::number(1.0)},         {"R2", "", Cell::number(2.0)},
      {"R2", "", Cell::number(4.0)},         {"R3", "", Cell::number(0.1)},
      {"R3", "", Cell::number(0.2)},         {"R3", "", Cell::number(0.3)},
      {"R4", "", Cell::number(1e-160)},      {"R4", "", Cell::number(1e-160)},
      {"R4", "", Cell::number(1e+200)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(1.0000000100000002e+32)},
      {"R2", "", Exp::number(8.0)},
      {"R3", "", Exp::number(0.006000000000000001)},
      {"R4", "", Exp::number(0.0)},
      {"GT", "", Exp::number(0.0)},
  };
  check_case(records, Aggregation::Product, false, expected);
}

// Average 1-D totals chain records; the corner combines raw row numerators.
TEST(PivotAggregationTotals, AverageTotalsA1) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(1e-10)}, {"R1", "P1", Cell::number(1e+16)},  {"R2", "P2", Cell::number(1e-17)},
      {"R2", "P3", Cell::number(12.0)},  {"R1", "P2", Cell::number(1e-10)},  {"R1", "P2", Cell::number(0.2)},
      {"R2", "P3", Cell::number(0.5)},   {"R2", "P1", Cell::number(-1e+16)}, {"R1", "P1", Cell::number(2.0)},
      {"R1", "P1", Cell::number(1e-10)}, {"R2", "P1", Cell::number(1.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(2500000000000000.5)},
      {"R1", "P2", Exp::number(0.10000000005000001)},
      {"R1", "P3", Exp::blank()},
      {"R1", "GT", Exp::number(1666666666666667.0)},
      {"R2", "P1", Exp::number(-5000000000000000.0)},
      {"R2", "P2", Exp::number(1e-17)},
      {"R2", "P3", Exp::number(6.25)},
      {"R2", "GT", Exp::number(-1999999999999997.5)},
      {"GT", "P1", Exp::number(0.16666666666666666)},
      {"GT", "P2", Exp::number(0.0666666667)},
      {"GT", "P3", Exp::number(6.25)},
      {"GT", "GT", Exp::number(1.2727272727272727)},
  };
  check_case(records, Aggregation::Average, true, expected);
}

// Average corner adds raw row numerators that the chain would cancel.
TEST(PivotAggregationTotals, AverageTotalsA2) {
  const std::vector<Record> records = {
      {"R1", "P2", Cell::number(2.0)},
      {"R2", "P1", Cell::number(-0.4)},
      {"R1", "P3", Cell::number(0.2)},
      {"R2", "P1", Cell::number(1e-10)},
      {"R2", "P3", Cell::number(1.0)},
      {"R2", "P3", Cell::number(0.7)},
      {"R2", "P3", Cell::number(0.2)},
      {"R2", "P2", Cell::number(9007199254740992.0)},
      {"R1", "P1", Cell::number(-9007199254740992.0)},
      {"R2", "P1", Cell::number(0.2)},
      {"R1", "P2", Cell::number(1e-10)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(-9007199254740992.0)},
      {"R1", "P2", Exp::number(1.00000000005)},
      {"R1", "P3", Exp::number(0.2)},
      {"R1", "GT", Exp::number(-2251799813685247.5)},
      {"R2", "P1", Exp::number(-0.06666666663333333)},
      {"R2", "P2", Exp::number(9007199254740992.0)},
      {"R2", "P3", Exp::number(0.6333333333333333)},
      {"R2", "GT", Exp::number(1286742750677284.5)},
      {"GT", "P1", Exp::number(-2251799813685248.0)},
      {"GT", "P2", Exp::number(3002399751580331.5)},
      {"GT", "P3", Exp::number(0.525)},
      {"GT", "GT", Exp::number(0.18181818181818182)},
  };
  check_case(records, Aggregation::Average, true, expected);
}

// Var 1-D totals chain S1 and S2 over records; the corner sums row partials.
TEST(PivotAggregationTotals, VarTotalsV1) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(0.7)}, {"R2", "P1", Cell::number(1.1)}, {"R2", "P1", Cell::number(1.7)},
      {"R2", "P2", Cell::number(0.9)}, {"R3", "P1", Cell::number(1.1)}, {"R2", "P2", Cell::number(0.3)},
      {"R1", "P1", Cell::number(0.3)}, {"R1", "P2", Cell::number(2.5)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.07999999999999996)}, {"R1", "P2", Exp::error(ErrorCode::Div0)},
      {"R1", "GT", Exp::number(1.3733333333333335)},  {"R2", "P1", Exp::number(0.18000000000000016)},
      {"R2", "P2", Exp::number(0.18000000000000005)}, {"R2", "GT", Exp::number(0.33333333333333365)},
      {"R3", "P1", Exp::error(ErrorCode::Div0)},      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::error(ErrorCode::Div0)},      {"GT", "P1", Exp::number(0.2719999999999996)},
      {"GT", "P2", Exp::number(1.293333333333334)},   {"GT", "GT", Exp::number(0.5421428571428574)},
  };
  check_case(records, Aggregation::Var, true, expected);
}

// StDev is the square root of the Var total.
TEST(PivotAggregationTotals, StdevTotalsV1) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(0.7)}, {"R2", "P1", Cell::number(1.1)}, {"R2", "P1", Cell::number(1.7)},
      {"R2", "P2", Cell::number(0.9)}, {"R3", "P1", Cell::number(1.1)}, {"R2", "P2", Cell::number(0.3)},
      {"R1", "P1", Cell::number(0.3)}, {"R1", "P2", Cell::number(2.5)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.28284271247461895)}, {"R1", "P2", Exp::error(ErrorCode::Div0)},
      {"R1", "GT", Exp::number(1.171893055416463)},   {"R2", "P1", Exp::number(0.4242640687119287)},
      {"R2", "P2", Exp::number(0.42426406871192857)}, {"R2", "GT", Exp::number(0.5773502691896261)},
      {"R3", "P1", Exp::error(ErrorCode::Div0)},      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::error(ErrorCode::Div0)},      {"GT", "P1", Exp::number(0.5215361924162115)},
      {"GT", "P2", Exp::number(1.1372481406154655)},  {"GT", "GT", Exp::number(0.7363035088486659)},
  };
  check_case(records, Aggregation::StdDev, true, expected);
}

// VarP uses (n*S2 - S1*S1) / (n*n).
TEST(PivotAggregationTotals, VarpTotalsV1) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(0.7)}, {"R2", "P1", Cell::number(1.1)}, {"R2", "P1", Cell::number(1.7)},
      {"R2", "P2", Cell::number(0.9)}, {"R3", "P1", Cell::number(1.1)}, {"R2", "P2", Cell::number(0.3)},
      {"R1", "P1", Cell::number(0.3)}, {"R1", "P2", Cell::number(2.5)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.03999999999999998)},
      {"R1", "P2", Exp::number(0.0)},
      {"R1", "GT", Exp::number(0.9155555555555558)},
      {"R2", "P1", Exp::number(0.09000000000000008)},
      {"R2", "P2", Exp::number(0.09000000000000002)},
      {"R2", "GT", Exp::number(0.2500000000000002)},
      {"R3", "P1", Exp::number(0.0)},
      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::number(0.0)},
      {"GT", "P1", Exp::number(0.21759999999999977)},
      {"GT", "P2", Exp::number(0.8622222222222228)},
      {"GT", "GT", Exp::number(0.4743750000000002)},
  };
  check_case(records, Aggregation::VarP, true, expected);
}

// Var totals with large magnitudes.
TEST(PivotAggregationTotals, VarTotalsV2) {
  const std::vector<Record> records = {
      {"R1", "P2", Cell::number(100000000.2)}, {"R2", "P1", Cell::number(1.1)},         {"R1", "P2", Cell::number(1.1)},
      {"R1", "P2", Cell::number(7.7)},         {"R2", "P2", Cell::number(1.3)},         {"R1", "P1", Cell::number(0.5)},
      {"R2", "P1", Cell::number(1.3)},         {"R1", "P1", Cell::number(100000000.2)}, {"R2", "P2", Cell::number(0.1)},
      {"R1", "P1", Cell::number(0.4)},         {"R1", "P1", Cell::number(0.25)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(2499999990833333.0)},
      {"R1", "P2", Exp::number(3333333053333351.0)},
      {"R1", "GT", Exp::number(2380952295714293.5)},
      {"R2", "P1", Exp::number(0.019999999999999574)},
      {"R2", "P2", Exp::number(0.72)},
      {"R2", "GT", Exp::number(0.33000000000000007)},
      {"GT", "P1", Exp::number(1666666649666667.0)},
      {"GT", "P2", Exp::number(1999999906000010.8)},
      {"GT", "GT", Exp::number(1636363592909095.5)},
  };
  check_case(records, Aggregation::Var, true, expected);
}

// VarP leaf with a large S2 residue.
TEST(PivotAggregationTotals, VarpTotalsV2) {
  const std::vector<Record> records = {
      {"R1", "P2", Cell::number(100000000.2)}, {"R2", "P1", Cell::number(1.1)},         {"R1", "P2", Cell::number(1.1)},
      {"R1", "P2", Cell::number(7.7)},         {"R2", "P2", Cell::number(1.3)},         {"R1", "P1", Cell::number(0.5)},
      {"R2", "P1", Cell::number(1.3)},         {"R1", "P1", Cell::number(100000000.2)}, {"R2", "P2", Cell::number(0.1)},
      {"R1", "P1", Cell::number(0.4)},         {"R1", "P1", Cell::number(0.25)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(1874999993124999.8)},
      {"R1", "P2", Exp::number(2222222035555567.0)},
      {"R1", "GT", Exp::number(2040816253469394.2)},
      {"R2", "P1", Exp::number(0.009999999999999787)},
      {"R2", "P2", Exp::number(0.36)},
      {"R2", "GT", Exp::number(0.24750000000000005)},
      {"GT", "P1", Exp::number(1388888874722222.5)},
      {"GT", "P2", Exp::number(1599999924800008.8)},
      {"GT", "GT", Exp::number(1487603266280996.0)},
  };
  check_case(records, Aggregation::VarP, true, expected);
}

// Product 1-D totals chain records; the corner multiplies row totals.
TEST(PivotAggregationTotals, ProductTotalsQ1) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(1.3)}, {"R2", "P2", Cell::number(1.7)}, {"R2", "P1", Cell::number(1.9)},
      {"R2", "P1", Cell::number(2.1)}, {"R2", "P3", Cell::number(3.3)}, {"R1", "P3", Cell::number(1.1)},
      {"R2", "P3", Cell::number(5.5)}, {"R1", "P1", Cell::number(0.7)}, {"R2", "P3", Cell::number(3.3)},
      {"R1", "P2", Cell::number(0.6)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.9099999999999999)},
      {"R1", "P2", Exp::number(0.6)},
      {"R1", "P3", Exp::number(1.1)},
      {"R1", "GT", Exp::number(0.6006)},
      {"R2", "P1", Exp::number(3.9899999999999998)},
      {"R2", "P2", Exp::number(1.7)},
      {"R2", "P3", Exp::number(59.89499999999999)},
      {"R2", "GT", Exp::number(406.26778499999995)},
      {"GT", "P1", Exp::number(3.6309)},
      {"GT", "P2", Exp::number(1.02)},
      {"GT", "P3", Exp::number(65.8845)},
      {"GT", "GT", Exp::number(244.00443167099996)},
  };
  check_case(records, Aggregation::Product, true, expected);
}

// Product corner is the raw product of row totals in display order.
TEST(PivotAggregationTotals, ProductTotalsQ2) {
  const std::vector<Record> records = {
      {"R1", "P2", Cell::number(0.7)}, {"R2", "P2", Cell::number(2.1)}, {"R2", "P2", Cell::number(1.1)},
      {"R2", "P2", Cell::number(5.5)}, {"R1", "P2", Cell::number(0.6)}, {"R1", "P1", Cell::number(1.1)},
      {"R2", "P1", Cell::number(1.9)}, {"R2", "P1", Cell::number(3.3)}, {"R1", "P1", Cell::number(1.7)},
      {"R2", "P1", Cell::number(1.9)}, {"R1", "P3", Cell::number(5.5)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(1.87)},
      {"R1", "P2", Exp::number(0.42)},
      {"R1", "P3", Exp::number(5.5)},
      {"R1", "GT", Exp::number(4.3197)},
      {"R2", "P1", Exp::number(11.912999999999998)},
      {"R2", "P2", Exp::number(12.705000000000002)},
      {"R2", "P3", Exp::blank()},
      {"R2", "GT", Exp::number(151.354665)},
      {"GT", "P1", Exp::number(22.27731)},
      {"GT", "P2", Exp::number(5.3361)},
      {"GT", "P3", Exp::number(5.5)},
      {"GT", "GT", Exp::number(653.8067464005001)},
  };
  check_case(records, Aggregation::Product, true, expected);
}

// Max row total takes the records (-5), not the max of leaf cells (0).
TEST(PivotAggregationTotals, MaxTextRecords) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::text("x")}, {"R1", "P2", Cell::number(-5.0)},  {"R2", "P1", Cell::number(3.0)},
      {"R2", "P2", Cell::text("y")}, {"R3", "P1", Cell::boolean(true)}, {"R3", "P2", Cell::number(-2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)}, {"R1", "P2", Exp::number(-5.0)}, {"R1", "GT", Exp::number(-5.0)},
      {"R2", "P1", Exp::number(3.0)}, {"R2", "P2", Exp::number(0.0)},  {"R2", "GT", Exp::number(3.0)},
      {"R3", "P1", Exp::number(0.0)}, {"R3", "P2", Exp::number(-2.0)}, {"R3", "GT", Exp::number(-2.0)},
      {"GT", "P1", Exp::number(3.0)}, {"GT", "P2", Exp::number(-2.0)}, {"GT", "GT", Exp::number(3.0)},
  };
  check_case(records, Aggregation::Max, true, expected);
}

// Min over text and boolean records.
TEST(PivotAggregationTotals, MinTextRecords) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::text("x")}, {"R1", "P2", Cell::number(-5.0)},  {"R2", "P1", Cell::number(3.0)},
      {"R2", "P2", Cell::text("y")}, {"R3", "P1", Cell::boolean(true)}, {"R3", "P2", Cell::number(-2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)}, {"R1", "P2", Exp::number(-5.0)}, {"R1", "GT", Exp::number(-5.0)},
      {"R2", "P1", Exp::number(3.0)}, {"R2", "P2", Exp::number(0.0)},  {"R2", "GT", Exp::number(3.0)},
      {"R3", "P1", Exp::number(0.0)}, {"R3", "P2", Exp::number(-2.0)}, {"R3", "GT", Exp::number(-2.0)},
      {"GT", "P1", Exp::number(3.0)}, {"GT", "P2", Exp::number(-5.0)}, {"GT", "GT", Exp::number(-5.0)},
  };
  check_case(records, Aggregation::Min, true, expected);
}

// Product over text and boolean records.
TEST(PivotAggregationTotals, ProductTextRecords) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::text("x")}, {"R1", "P2", Cell::number(-5.0)},  {"R2", "P1", Cell::number(3.0)},
      {"R2", "P2", Cell::text("y")}, {"R3", "P1", Cell::boolean(true)}, {"R3", "P2", Cell::number(-2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)}, {"R1", "P2", Exp::number(-5.0)}, {"R1", "GT", Exp::number(-5.0)},
      {"R2", "P1", Exp::number(3.0)}, {"R2", "P2", Exp::number(0.0)},  {"R2", "GT", Exp::number(3.0)},
      {"R3", "P1", Exp::number(0.0)}, {"R3", "P2", Exp::number(-2.0)}, {"R3", "GT", Exp::number(-2.0)},
      {"GT", "P1", Exp::number(3.0)}, {"GT", "P2", Exp::number(10.0)}, {"GT", "GT", Exp::number(30.0)},
  };
  check_case(records, Aggregation::Product, true, expected);
}

// a row total chains records and never sees the leaf cell that flushed to 0.
TEST(PivotAggregationTotals, ProductNoSubnormalChain) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(1e-160)}, {"R1", "P2", Cell::number(1e-160)}, {"R2", "P1", Cell::number(1e+200)},
      {"R2", "P2", Cell::number(1.0)},    {"R3", "P1", Cell::number(1e+200)}, {"R3", "P2", Cell::number(1e-160)},
      {"R3", "P2", Cell::number(1e-160)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(1e-160)}, {"R1", "P2", Exp::number(1e-160)}, {"R1", "GT", Exp::number(0.0)},
      {"R2", "P1", Exp::number(1e+200)}, {"R2", "P2", Exp::number(1.0)},    {"R2", "GT", Exp::number(1e+200)},
      {"R3", "P1", Exp::number(1e+200)}, {"R3", "P2", Exp::number(0.0)},    {"R3", "GT", Exp::number(1e-120)},
      {"GT", "P1", Exp::number(0.0)},    {"GT", "P2", Exp::number(0.0)},    {"GT", "GT", Exp::number(0.0)},
  };
  check_case(records, Aggregation::Product, true, expected);
}

// a negative one-pass variance is 0, not #NUM!.
TEST(PivotAggregationTotals, VarNegativeResidue) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(100000000.1)}, {"R1", "P1", Cell::number(100000000.2)},
      {"R2", "P1", Cell::number(1.7)},         {"R2", "P1", Cell::number(1.7)},
      {"R2", "P1", Cell::number(1.7)},         {"R3", "P1", Cell::number(0.1)},
      {"R3", "P1", Cell::number(0.1)},         {"R3", "P1", Cell::number(0.1)},
      {"R4", "P1", Cell::number(100000000.1)}, {"R4", "P1", Cell::number(100000000.1)},
      {"R4", "P1", Cell::number(100000000.3)}, {"R1", "P2", Cell::number(1.0)},
      {"R2", "P2", Cell::number(2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)},
      {"R1", "P2", Exp::error(ErrorCode::Div0)},
      {"R1", "GT", Exp::number(3333333276666665.0)},
      {"R2", "P1", Exp::number(0.0)},
      {"R2", "P2", Exp::error(ErrorCode::Div0)},
      {"R2", "GT", Exp::number(0.02249999999999967)},
      {"R3", "P1", Exp::number(0.0)},
      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::number(0.0)},
      {"R4", "P1", Exp::number(0.0)},
      {"R4", "P2", Exp::blank()},
      {"R4", "GT", Exp::number(0.0)},
      {"GT", "P1", Exp::number(2727272686909091.5)},
      {"GT", "P2", Exp::number(0.5)},
      {"GT", "GT", Exp::number(2564102518461538.5)},
  };
  check_case(records, Aggregation::Var, true, expected);
}

// a negative one-pass population variance is 0.
TEST(PivotAggregationTotals, VarpNegativeResidue) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(100000000.1)}, {"R1", "P1", Cell::number(100000000.2)},
      {"R2", "P1", Cell::number(1.7)},         {"R2", "P1", Cell::number(1.7)},
      {"R2", "P1", Cell::number(1.7)},         {"R3", "P1", Cell::number(0.1)},
      {"R3", "P1", Cell::number(0.1)},         {"R3", "P1", Cell::number(0.1)},
      {"R4", "P1", Cell::number(100000000.1)}, {"R4", "P1", Cell::number(100000000.1)},
      {"R4", "P1", Cell::number(100000000.3)}, {"R1", "P2", Cell::number(1.0)},
      {"R2", "P2", Cell::number(2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)},
      {"R1", "P2", Exp::number(0.0)},
      {"R1", "GT", Exp::number(2222222184444443.5)},
      {"R2", "P1", Exp::number(0.0)},
      {"R2", "P2", Exp::number(0.0)},
      {"R2", "GT", Exp::number(0.01687499999999975)},
      {"R3", "P1", Exp::number(0.0)},
      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::number(0.0)},
      {"R4", "P1", Exp::number(0.0)},
      {"R4", "P2", Exp::blank()},
      {"R4", "GT", Exp::number(0.0)},
      {"GT", "P1", Exp::number(2479338806280991.5)},
      {"GT", "P2", Exp::number(0.25)},
      {"GT", "GT", Exp::number(2366863863195266.5)},
  };
  check_case(records, Aggregation::VarP, true, expected);
}

// StDev of a negative residue is 0.
TEST(PivotAggregationTotals, StdevNegativeResidue) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(100000000.1)}, {"R1", "P1", Cell::number(100000000.2)},
      {"R2", "P1", Cell::number(1.7)},         {"R2", "P1", Cell::number(1.7)},
      {"R2", "P1", Cell::number(1.7)},         {"R3", "P1", Cell::number(0.1)},
      {"R3", "P1", Cell::number(0.1)},         {"R3", "P1", Cell::number(0.1)},
      {"R4", "P1", Cell::number(100000000.1)}, {"R4", "P1", Cell::number(100000000.1)},
      {"R4", "P1", Cell::number(100000000.3)}, {"R1", "P2", Cell::number(1.0)},
      {"R2", "P2", Cell::number(2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)},
      {"R1", "P2", Exp::error(ErrorCode::Div0)},
      {"R1", "GT", Exp::number(57735026.42821483)},
      {"R2", "P1", Exp::number(0.0)},
      {"R2", "P2", Exp::error(ErrorCode::Div0)},
      {"R2", "GT", Exp::number(0.1499999999999989)},
      {"R3", "P1", Exp::number(0.0)},
      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::number(0.0)},
      {"R4", "P1", Exp::number(0.0)},
      {"R4", "P2", Exp::blank()},
      {"R4", "GT", Exp::number(0.0)},
      {"GT", "P1", Exp::number(52223296.40025696)},
      {"GT", "P2", Exp::number(0.7071067811865476)},
      {"GT", "GT", Exp::number(50636967.90351431)},
  };
  check_case(records, Aggregation::StdDev, true, expected);
}

// StDevP of a negative residue is 0.
TEST(PivotAggregationTotals, StdevpNegativeResidue) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(100000000.1)}, {"R1", "P1", Cell::number(100000000.2)},
      {"R2", "P1", Cell::number(1.7)},         {"R2", "P1", Cell::number(1.7)},
      {"R2", "P1", Cell::number(1.7)},         {"R3", "P1", Cell::number(0.1)},
      {"R3", "P1", Cell::number(0.1)},         {"R3", "P1", Cell::number(0.1)},
      {"R4", "P1", Cell::number(100000000.1)}, {"R4", "P1", Cell::number(100000000.1)},
      {"R4", "P1", Cell::number(100000000.3)}, {"R1", "P2", Cell::number(1.0)},
      {"R2", "P2", Cell::number(2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.0)},
      {"R1", "P2", Exp::number(0.0)},
      {"R1", "GT", Exp::number(47140451.678409316)},
      {"R2", "P1", Exp::number(0.0)},
      {"R2", "P2", Exp::number(0.0)},
      {"R2", "GT", Exp::number(0.12990381056766484)},
      {"R3", "P1", Exp::number(0.0)},
      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::number(0.0)},
      {"R4", "P1", Exp::number(0.0)},
      {"R4", "P2", Exp::blank()},
      {"R4", "GT", Exp::number(0.0)},
      {"GT", "P1", Exp::number(49792959.404729016)},
      {"GT", "P2", Exp::number(0.5)},
      {"GT", "GT", Exp::number(48650425.108063206)},
  };
  check_case(records, Aggregation::StdDevP, true, expected);
}

// a row whose records are all blank is blank for Count in every column.
TEST(PivotAggregationTotals, CountBlankRow) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::blank()}, {"R1", "P2", Cell::number(1.0)}, {"R2", "P1", Cell::number(2.0)},
      {"R2", "P2", Cell::blank()}, {"R2", "P2", Cell::text("t")},   {"R3", "P1", Cell::blank()},
      {"R3", "P2", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},     {"R1", "P2", Exp::number(1.0)}, {"R1", "GT", Exp::number(1.0)},
      {"R2", "P1", Exp::number(1.0)}, {"R2", "P2", Exp::number(1.0)}, {"R2", "GT", Exp::number(2.0)},
      {"R3", "P1", Exp::blank()},     {"R3", "P2", Exp::blank()},     {"R3", "GT", Exp::blank()},
      {"GT", "P1", Exp::number(1.0)}, {"GT", "P2", Exp::number(2.0)}, {"GT", "GT", Exp::number(3.0)},
  };
  check_case(records, Aggregation::Count, true, expected);
}

// a row whose records are all blank is blank for CountNums.
TEST(PivotAggregationTotals, CountNumsBlankRow) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::blank()}, {"R1", "P2", Cell::number(1.0)}, {"R2", "P1", Cell::number(2.0)},
      {"R2", "P2", Cell::blank()}, {"R2", "P2", Cell::text("t")},   {"R3", "P1", Cell::blank()},
      {"R3", "P2", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},     {"R1", "P2", Exp::number(1.0)}, {"R1", "GT", Exp::number(1.0)},
      {"R2", "P1", Exp::number(1.0)}, {"R2", "P2", Exp::number(0.0)}, {"R2", "GT", Exp::number(1.0)},
      {"R3", "P1", Exp::blank()},     {"R3", "P2", Exp::blank()},     {"R3", "GT", Exp::blank()},
      {"GT", "P1", Exp::number(1.0)}, {"GT", "P2", Exp::number(1.0)}, {"GT", "GT", Exp::number(2.0)},
  };
  check_case(records, Aggregation::CountNumbers, true, expected);
}

// a blank row stays blank; a blank-and-text cell is #DIV/0!.
TEST(PivotAggregationTotals, AverageBlankRow) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::blank()}, {"R1", "P2", Cell::number(1.0)}, {"R2", "P1", Cell::number(2.0)},
      {"R2", "P2", Cell::blank()}, {"R2", "P2", Cell::text("t")},   {"R3", "P1", Cell::blank()},
      {"R3", "P2", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},
      {"R1", "P2", Exp::number(1.0)},
      {"R1", "GT", Exp::number(1.0)},
      {"R2", "P1", Exp::number(2.0)},
      {"R2", "P2", Exp::error(ErrorCode::Div0)},
      {"R2", "GT", Exp::number(2.0)},
      {"R3", "P1", Exp::blank()},
      {"R3", "P2", Exp::blank()},
      {"R3", "GT", Exp::blank()},
      {"GT", "P1", Exp::number(2.0)},
      {"GT", "P2", Exp::number(1.0)},
      {"GT", "GT", Exp::number(1.5)},
  };
  check_case(records, Aggregation::Average, true, expected);
}

// a blank-and-text cell has Max 0.
TEST(PivotAggregationTotals, MaxBlankRow) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::blank()}, {"R1", "P2", Cell::number(1.0)}, {"R2", "P1", Cell::number(2.0)},
      {"R2", "P2", Cell::blank()}, {"R2", "P2", Cell::text("t")},   {"R3", "P1", Cell::blank()},
      {"R3", "P2", Cell::blank()},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},     {"R1", "P2", Exp::number(1.0)}, {"R1", "GT", Exp::number(1.0)},
      {"R2", "P1", Exp::number(2.0)}, {"R2", "P2", Exp::number(0.0)}, {"R2", "GT", Exp::number(2.0)},
      {"R3", "P1", Exp::blank()},     {"R3", "P2", Exp::blank()},     {"R3", "GT", Exp::blank()},
      {"GT", "P1", Exp::number(2.0)}, {"GT", "P2", Exp::number(1.0)}, {"GT", "GT", Exp::number(2.0)},
  };
  check_case(records, Aggregation::Max, true, expected);
}

// Min snaps against the record just added.
TEST(PivotAggregationTotals, MinTinyMagnitudes) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1e-160)}, {"R02", "", Cell::number(1e-160)},   {"R02", "", Cell::number(1e+200)},
      {"R03", "", Cell::number(1e-20)},  {"R04", "", Cell::number(1e-10)},    {"R05", "", Cell::number(-1e-160)},
      {"R06", "", Cell::number(1e-300)}, {"R07", "", Cell::number(2.5e-308)}, {"R08", "", Cell::number(1e-160)},
      {"R08", "", Cell::number(1e-160)}, {"R09", "", Cell::number(1e+200)},   {"R10", "", Cell::number(5.0)},
      {"R10", "", Cell::number(1e-20)},  {"R11", "", Cell::number(1e-05)},    {"R12", "", Cell::number(1e-15)},
      {"R12", "", Cell::number(1.0)},    {"R13", "", Cell::number(1.0)},      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-100)}, {"R14", "", Cell::number(1e+100)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(1e-160)},   {"R02", "", Exp::number(0.0)},     {"R03", "", Exp::number(1e-20)},
      {"R04", "", Exp::number(1e-10)},    {"R05", "", Exp::number(-1e-160)}, {"R06", "", Exp::number(1e-300)},
      {"R07", "", Exp::number(2.5e-308)}, {"R08", "", Exp::number(1e-160)},  {"R09", "", Exp::number(1e+200)},
      {"R10", "", Exp::number(1e-20)},    {"R11", "", Exp::number(1e-05)},   {"R12", "", Exp::number(0.0)},
      {"R13", "", Exp::number(1e-15)},    {"R14", "", Exp::number(0.0)},     {"GT", "", Exp::number(0.0)},
  };
  check_case(records, Aggregation::Min, false, expected);
}

// Product snaps against the record just added.
TEST(PivotAggregationTotals, ProductTinyMagnitudes) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1e-160)}, {"R02", "", Cell::number(1e-160)},   {"R02", "", Cell::number(1e+200)},
      {"R03", "", Cell::number(1e-20)},  {"R04", "", Cell::number(1e-10)},    {"R05", "", Cell::number(-1e-160)},
      {"R06", "", Cell::number(1e-300)}, {"R07", "", Cell::number(2.5e-308)}, {"R08", "", Cell::number(1e-160)},
      {"R08", "", Cell::number(1e-160)}, {"R09", "", Cell::number(1e+200)},   {"R10", "", Cell::number(5.0)},
      {"R10", "", Cell::number(1e-20)},  {"R11", "", Cell::number(1e-05)},    {"R12", "", Cell::number(1e-15)},
      {"R12", "", Cell::number(1.0)},    {"R13", "", Cell::number(1.0)},      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-100)}, {"R14", "", Cell::number(1e+100)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(1e-160)},   {"R02", "", Exp::number(0.0)},     {"R03", "", Exp::number(1e-20)},
      {"R04", "", Exp::number(1e-10)},    {"R05", "", Exp::number(-1e-160)}, {"R06", "", Exp::number(1e-300)},
      {"R07", "", Exp::number(2.5e-308)}, {"R08", "", Exp::number(0.0)},     {"R09", "", Exp::number(1e+200)},
      {"R10", "", Exp::number(5e-20)},    {"R11", "", Exp::number(1e-05)},   {"R12", "", Exp::number(0.0)},
      {"R13", "", Exp::number(1e-15)},    {"R14", "", Exp::number(0.0)},     {"GT", "", Exp::number(0.0)},
  };
  check_case(records, Aggregation::Product, false, expected);
}

// Product snap boundary at 8 ulps, post-operation.
TEST(PivotAggregationTotals, ProductSnapBoundary) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1.5543122344752192e-15)},
      {"R01", "", Cell::number(1.0)},
      {"R02", "", Cell::number(1.7763568394002505e-15)},
      {"R02", "", Cell::number(1.0)},
      {"R03", "", Cell::number(1.9984014443252818e-15)},
      {"R03", "", Cell::number(1.0)},
      {"R04", "", Cell::number(2.6645352591003757e-15)},
      {"R04", "", Cell::number(2.0)},
      {"R05", "", Cell::number(2.6645352591003757e-15)},
      {"R05", "", Cell::number(1.9)},
      {"R06", "", Cell::number(1e-15)},
      {"R06", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(-1.0)},
      {"R08", "", Cell::number(-1e-15)},
      {"R08", "", Cell::number(-1.0)},
      {"R09", "", Cell::number(1e-15)},
      {"R09", "", Cell::number(1.0)},
      {"R09", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1.0)},
      {"R11", "", Cell::number(1000000000000000.0)},
      {"R12", "", Cell::number(1e-20)},
      {"R12", "", Cell::number(3e-20)},
      {"R13", "", Cell::number(3.0)},
      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(3.0)},
      {"R15", "", Cell::number(1e-20)},
      {"R15", "", Cell::number(3e-20)},
      {"R15", "", Cell::number(5e-20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R17", "", Cell::number(0.5)},
      {"R17", "", Cell::number(-0.4)},
      {"R17", "", Cell::number(-0.1)},
      {"R17", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(0.0)},
      {"R02", "", Exp::number(1.7763568394002505e-15)},
      {"R03", "", Exp::number(1.9984014443252818e-15)},
      {"R04", "", Exp::number(5.329070518200751e-15)},
      {"R05", "", Exp::number(5.062616992290713e-15)},
      {"R06", "", Exp::number(0.0)},
      {"R07", "", Exp::number(0.0)},
      {"R08", "", Exp::number(0.0)},
      {"R09", "", Exp::number(0.0)},
      {"R10", "", Exp::number(0.0)},
      {"R11", "", Exp::number(0.0)},
      {"R12", "", Exp::number(0.0)},
      {"R13", "", Exp::number(3.0000000000000002e-15)},
      {"R14", "", Exp::number(0.0)},
      {"R15", "", Exp::number(0.0)},
      {"R16", "", Exp::number(1e+60)},
      {"R17", "", Exp::number(2.0000000000000005e-17)},
      {"R18", "", Exp::number(0.0)},
      {"GT", "", Exp::number(0.0)},
  };
  check_case(records, Aggregation::Product, false, expected);
}

// Min snap boundary at 8 ulps.
TEST(PivotAggregationTotals, MinSnapBoundary) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1.5543122344752192e-15)},
      {"R01", "", Cell::number(1.0)},
      {"R02", "", Cell::number(1.7763568394002505e-15)},
      {"R02", "", Cell::number(1.0)},
      {"R03", "", Cell::number(1.9984014443252818e-15)},
      {"R03", "", Cell::number(1.0)},
      {"R04", "", Cell::number(2.6645352591003757e-15)},
      {"R04", "", Cell::number(2.0)},
      {"R05", "", Cell::number(2.6645352591003757e-15)},
      {"R05", "", Cell::number(1.9)},
      {"R06", "", Cell::number(1e-15)},
      {"R06", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(-1.0)},
      {"R08", "", Cell::number(-1e-15)},
      {"R08", "", Cell::number(-1.0)},
      {"R09", "", Cell::number(1e-15)},
      {"R09", "", Cell::number(1.0)},
      {"R09", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1.0)},
      {"R11", "", Cell::number(1000000000000000.0)},
      {"R12", "", Cell::number(1e-20)},
      {"R12", "", Cell::number(3e-20)},
      {"R13", "", Cell::number(3.0)},
      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(3.0)},
      {"R15", "", Cell::number(1e-20)},
      {"R15", "", Cell::number(3e-20)},
      {"R15", "", Cell::number(5e-20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R17", "", Cell::number(0.5)},
      {"R17", "", Cell::number(-0.4)},
      {"R17", "", Cell::number(-0.1)},
      {"R17", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(0.0)},
      {"R02", "", Exp::number(1.7763568394002505e-15)},
      {"R03", "", Exp::number(1.9984014443252818e-15)},
      {"R04", "", Exp::number(0.0)},
      {"R05", "", Exp::number(2.6645352591003757e-15)},
      {"R06", "", Exp::number(1e-15)},
      {"R07", "", Exp::number(-1.0)},
      {"R08", "", Exp::number(-1.0)},
      {"R09", "", Exp::number(0.0)},
      {"R10", "", Exp::number(1e-16)},
      {"R11", "", Exp::number(0.0)},
      {"R12", "", Exp::number(1e-20)},
      {"R13", "", Exp::number(1e-15)},
      {"R14", "", Exp::number(0.0)},
      {"R15", "", Exp::number(1e-20)},
      {"R16", "", Exp::number(1e+20)},
      {"R17", "", Exp::number(-0.4)},
      {"R18", "", Exp::number(1e-15)},
      {"GT", "", Exp::number(-0.4)},
  };
  check_case(records, Aggregation::Min, false, expected);
}

// Max snap against a negative or tiny record.
TEST(PivotAggregationTotals, MaxSnapBoundary) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1.5543122344752192e-15)},
      {"R01", "", Cell::number(1.0)},
      {"R02", "", Cell::number(1.7763568394002505e-15)},
      {"R02", "", Cell::number(1.0)},
      {"R03", "", Cell::number(1.9984014443252818e-15)},
      {"R03", "", Cell::number(1.0)},
      {"R04", "", Cell::number(2.6645352591003757e-15)},
      {"R04", "", Cell::number(2.0)},
      {"R05", "", Cell::number(2.6645352591003757e-15)},
      {"R05", "", Cell::number(1.9)},
      {"R06", "", Cell::number(1e-15)},
      {"R06", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(-1.0)},
      {"R08", "", Cell::number(-1e-15)},
      {"R08", "", Cell::number(-1.0)},
      {"R09", "", Cell::number(1e-15)},
      {"R09", "", Cell::number(1.0)},
      {"R09", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1.0)},
      {"R11", "", Cell::number(1000000000000000.0)},
      {"R12", "", Cell::number(1e-20)},
      {"R12", "", Cell::number(3e-20)},
      {"R13", "", Cell::number(3.0)},
      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(3.0)},
      {"R15", "", Cell::number(1e-20)},
      {"R15", "", Cell::number(3e-20)},
      {"R15", "", Cell::number(5e-20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R17", "", Cell::number(0.5)},
      {"R17", "", Cell::number(-0.4)},
      {"R17", "", Cell::number(-0.1)},
      {"R17", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(1.0)},
      {"R02", "", Exp::number(1.0)},
      {"R03", "", Exp::number(1.0)},
      {"R04", "", Exp::number(2.0)},
      {"R05", "", Exp::number(1.9)},
      {"R06", "", Exp::number(1e-15)},
      {"R07", "", Exp::number(0.0)},
      {"R08", "", Exp::number(0.0)},
      {"R09", "", Exp::number(1.0)},
      {"R10", "", Exp::number(1e-15)},
      {"R11", "", Exp::number(1000000000000000.0)},
      {"R12", "", Exp::number(3e-20)},
      {"R13", "", Exp::number(3.0)},
      {"R14", "", Exp::number(3.0)},
      {"R15", "", Exp::number(5e-20)},
      {"R16", "", Exp::number(1e+20)},
      {"R17", "", Exp::number(0.5)},
      {"R18", "", Exp::number(1e-15)},
      {"GT", "", Exp::number(1e+20)},
  };
  check_case(records, Aggregation::Max, false, expected);
}

// Var sum of squares is not snapped.
TEST(PivotAggregationTotals, VarSnapBoundary) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1.5543122344752192e-15)},
      {"R01", "", Cell::number(1.0)},
      {"R02", "", Cell::number(1.7763568394002505e-15)},
      {"R02", "", Cell::number(1.0)},
      {"R03", "", Cell::number(1.9984014443252818e-15)},
      {"R03", "", Cell::number(1.0)},
      {"R04", "", Cell::number(2.6645352591003757e-15)},
      {"R04", "", Cell::number(2.0)},
      {"R05", "", Cell::number(2.6645352591003757e-15)},
      {"R05", "", Cell::number(1.9)},
      {"R06", "", Cell::number(1e-15)},
      {"R06", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(-1.0)},
      {"R08", "", Cell::number(-1e-15)},
      {"R08", "", Cell::number(-1.0)},
      {"R09", "", Cell::number(1e-15)},
      {"R09", "", Cell::number(1.0)},
      {"R09", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1.0)},
      {"R11", "", Cell::number(1000000000000000.0)},
      {"R12", "", Cell::number(1e-20)},
      {"R12", "", Cell::number(3e-20)},
      {"R13", "", Cell::number(3.0)},
      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(3.0)},
      {"R15", "", Cell::number(1e-20)},
      {"R15", "", Cell::number(3e-20)},
      {"R15", "", Cell::number(5e-20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R17", "", Cell::number(0.5)},
      {"R17", "", Cell::number(-0.4)},
      {"R17", "", Cell::number(-0.1)},
      {"R17", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(0.49999999999999845)},
      {"R02", "", Exp::number(0.4999999999999982)},
      {"R03", "", Exp::number(0.499999999999998)},
      {"R04", "", Exp::number(1.9999999999999947)},
      {"R05", "", Exp::number(1.804999999999995)},
      {"R06", "", Exp::number(0.0)},
      {"R07", "", Exp::number(0.500000000000001)},
      {"R08", "", Exp::number(0.4999999999999989)},
      {"R09", "", Exp::number(0.3333333333333329)},
      {"R10", "", Exp::number(4.049999999999999e-31)},
      {"R11", "", Exp::number(3.3333333333333296e+29)},
      {"R12", "", Exp::number(2e-40)},
      {"R13", "", Exp::number(4.499999999999997)},
      {"R14", "", Exp::number(4.499999999999997)},
      {"R15", "", Exp::number(4e-40)},
      {"R16", "", Exp::number(0.0)},
      {"R17", "", Exp::number(0.14)},
      {"R18", "", Exp::number(0.0)},
      {"GT", "", Exp::number(6.644514950398672e+38)},
  };
  check_case(records, Aggregation::Var, false, expected);
}

// VarP sum of squares is not snapped.
TEST(PivotAggregationTotals, VarpSnapBoundary) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1.5543122344752192e-15)},
      {"R01", "", Cell::number(1.0)},
      {"R02", "", Cell::number(1.7763568394002505e-15)},
      {"R02", "", Cell::number(1.0)},
      {"R03", "", Cell::number(1.9984014443252818e-15)},
      {"R03", "", Cell::number(1.0)},
      {"R04", "", Cell::number(2.6645352591003757e-15)},
      {"R04", "", Cell::number(2.0)},
      {"R05", "", Cell::number(2.6645352591003757e-15)},
      {"R05", "", Cell::number(1.9)},
      {"R06", "", Cell::number(1e-15)},
      {"R06", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(-1.0)},
      {"R08", "", Cell::number(-1e-15)},
      {"R08", "", Cell::number(-1.0)},
      {"R09", "", Cell::number(1e-15)},
      {"R09", "", Cell::number(1.0)},
      {"R09", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1.0)},
      {"R11", "", Cell::number(1000000000000000.0)},
      {"R12", "", Cell::number(1e-20)},
      {"R12", "", Cell::number(3e-20)},
      {"R13", "", Cell::number(3.0)},
      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(3.0)},
      {"R15", "", Cell::number(1e-20)},
      {"R15", "", Cell::number(3e-20)},
      {"R15", "", Cell::number(5e-20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R17", "", Cell::number(0.5)},
      {"R17", "", Cell::number(-0.4)},
      {"R17", "", Cell::number(-0.1)},
      {"R17", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(0.24999999999999922)},    {"R02", "", Exp::number(0.2499999999999991)},
      {"R03", "", Exp::number(0.249999999999999)},      {"R04", "", Exp::number(0.9999999999999973)},
      {"R05", "", Exp::number(0.9024999999999975)},     {"R06", "", Exp::number(0.0)},
      {"R07", "", Exp::number(0.2500000000000005)},     {"R08", "", Exp::number(0.24999999999999944)},
      {"R09", "", Exp::number(0.222222222222222)},      {"R10", "", Exp::number(2.0249999999999996e-31)},
      {"R11", "", Exp::number(2.2222222222222204e+29)}, {"R12", "", Exp::number(1e-40)},
      {"R13", "", Exp::number(2.2499999999999987)},     {"R14", "", Exp::number(2.2499999999999987)},
      {"R15", "", Exp::number(2.6666666666666657e-40)}, {"R16", "", Exp::number(0.0)},
      {"R17", "", Exp::number(0.10500000000000001)},    {"R18", "", Exp::number(0.0)},
      {"GT", "", Exp::number(6.489991346901028e+38)},
  };
  check_case(records, Aggregation::VarP, false, expected);
}

// Count is never snapped.
TEST(PivotAggregationTotals, CountNeverSnapped) {
  const std::vector<Record> records = {
      {"R01", "", Cell::number(1.5543122344752192e-15)},
      {"R01", "", Cell::number(1.0)},
      {"R02", "", Cell::number(1.7763568394002505e-15)},
      {"R02", "", Cell::number(1.0)},
      {"R03", "", Cell::number(1.9984014443252818e-15)},
      {"R03", "", Cell::number(1.0)},
      {"R04", "", Cell::number(2.6645352591003757e-15)},
      {"R04", "", Cell::number(2.0)},
      {"R05", "", Cell::number(2.6645352591003757e-15)},
      {"R05", "", Cell::number(1.9)},
      {"R06", "", Cell::number(1e-15)},
      {"R06", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(1e-15)},
      {"R07", "", Cell::number(-1.0)},
      {"R08", "", Cell::number(-1e-15)},
      {"R08", "", Cell::number(-1.0)},
      {"R09", "", Cell::number(1e-15)},
      {"R09", "", Cell::number(1.0)},
      {"R09", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-16)},
      {"R10", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1e-15)},
      {"R11", "", Cell::number(1.0)},
      {"R11", "", Cell::number(1000000000000000.0)},
      {"R12", "", Cell::number(1e-20)},
      {"R12", "", Cell::number(3e-20)},
      {"R13", "", Cell::number(3.0)},
      {"R13", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(1e-15)},
      {"R14", "", Cell::number(3.0)},
      {"R15", "", Cell::number(1e-20)},
      {"R15", "", Cell::number(3e-20)},
      {"R15", "", Cell::number(5e-20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R16", "", Cell::number(1e+20)},
      {"R17", "", Cell::number(0.5)},
      {"R17", "", Cell::number(-0.4)},
      {"R17", "", Cell::number(-0.1)},
      {"R17", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
      {"R18", "", Cell::number(1e-15)},
  };
  const std::vector<Expected> expected = {
      {"R01", "", Exp::number(2.0)}, {"R02", "", Exp::number(2.0)}, {"R03", "", Exp::number(2.0)},
      {"R04", "", Exp::number(2.0)}, {"R05", "", Exp::number(2.0)}, {"R06", "", Exp::number(2.0)},
      {"R07", "", Exp::number(2.0)}, {"R08", "", Exp::number(2.0)}, {"R09", "", Exp::number(3.0)},
      {"R10", "", Exp::number(2.0)}, {"R11", "", Exp::number(3.0)}, {"R12", "", Exp::number(2.0)},
      {"R13", "", Exp::number(2.0)}, {"R14", "", Exp::number(2.0)}, {"R15", "", Exp::number(3.0)},
      {"R16", "", Exp::number(3.0)}, {"R17", "", Exp::number(4.0)}, {"R18", "", Exp::number(3.0)},
      {"GT", "", Exp::number(43.0)},
  };
  check_case(records, Aggregation::Count, false, expected);
}

// VarP association is (n*S2 - S1*S1) / (n*n).
TEST(PivotAggregationTotals, VarpAssociation) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(0.4)},   {"R1", "", Cell::number(12.5)}, {"R1", "", Cell::number(0.4)},
      {"R1", "", Cell::number(12.5)},  {"R1", "", Cell::number(3.3)},  {"R2", "", Cell::number(0.9)},
      {"R2", "", Cell::number(1.3)},   {"R2", "", Cell::number(0.2)},  {"R3", "", Cell::number(0.6)},
      {"R3", "", Cell::number(0.4)},   {"R3", "", Cell::number(0.9)},  {"R4", "", Cell::number(1.3)},
      {"R4", "", Cell::number(0.3)},   {"R4", "", Cell::number(0.4)},  {"R5", "", Cell::number(1e-20)},
      {"R5", "", Cell::number(3e-20)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(30.869599999999995)},
      {"R2", "", Exp::number(0.2066666666666665)},
      {"R3", "", Exp::number(0.04222222222222226)},
      {"R4", "", Exp::number(0.20222222222222236)},
      {"R5", "", Exp::number(1e-40)},
      {"GT", "", Exp::number(15.699843750000005)},
  };
  check_case(records, Aggregation::VarP, false, expected);
}

// Var association is (S2 - S1*S1/n) / (n-1).
TEST(PivotAggregationTotals, VarAssociation) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(0.4)},   {"R1", "", Cell::number(12.5)}, {"R1", "", Cell::number(0.4)},
      {"R1", "", Cell::number(12.5)},  {"R1", "", Cell::number(3.3)},  {"R2", "", Cell::number(0.9)},
      {"R2", "", Cell::number(1.3)},   {"R2", "", Cell::number(0.2)},  {"R3", "", Cell::number(0.6)},
      {"R3", "", Cell::number(0.4)},   {"R3", "", Cell::number(0.9)},  {"R4", "", Cell::number(1.3)},
      {"R4", "", Cell::number(0.3)},   {"R4", "", Cell::number(0.4)},  {"R5", "", Cell::number(1e-20)},
      {"R5", "", Cell::number(3e-20)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(38.58699999999999)},
      {"R2", "", Exp::number(0.3099999999999997)},
      {"R3", "", Exp::number(0.06333333333333335)},
      {"R4", "", Exp::number(0.30333333333333357)},
      {"R5", "", Exp::number(2e-40)},
      {"GT", "", Exp::number(16.746500000000005)},
  };
  check_case(records, Aggregation::Var, false, expected);
}

// StDevP is sqrt of VarP bit for bit.
TEST(PivotAggregationTotals, StdevpAssociation) {
  const std::vector<Record> records = {
      {"R1", "", Cell::number(0.4)},   {"R1", "", Cell::number(12.5)}, {"R1", "", Cell::number(0.4)},
      {"R1", "", Cell::number(12.5)},  {"R1", "", Cell::number(3.3)},  {"R2", "", Cell::number(0.9)},
      {"R2", "", Cell::number(1.3)},   {"R2", "", Cell::number(0.2)},  {"R3", "", Cell::number(0.6)},
      {"R3", "", Cell::number(0.4)},   {"R3", "", Cell::number(0.9)},  {"R4", "", Cell::number(1.3)},
      {"R4", "", Cell::number(0.3)},   {"R4", "", Cell::number(0.4)},  {"R5", "", Cell::number(1e-20)},
      {"R5", "", Cell::number(3e-20)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::number(5.556041756502555)},
      {"R2", "", Exp::number(0.454606056566195)},
      {"R3", "", Exp::number(0.20548046676563264)},
      {"R4", "", Exp::number(0.44969125210773486)},
      {"R5", "", Exp::number(1e-20)},
      {"GT", "", Exp::number(3.9623028342114393)},
  };
  check_case(records, Aggregation::StdDevP, false, expected);
}

// Var corner sums S1 and S2 over the row partials.
TEST(PivotAggregationTotals, VarCornerRowOrder) {
  const std::vector<Record> records = {
      {"R1", "P3", Cell::number(2.1)}, {"R1", "P2", Cell::number(3.3)}, {"R1", "P3", Cell::number(0.1)},
      {"R1", "P2", Cell::number(1.1)}, {"R1", "P2", Cell::number(2.1)}, {"R2", "P2", Cell::number(0.6)},
      {"R3", "P3", Cell::number(1.7)}, {"R1", "P2", Cell::number(0.3)}, {"R2", "P1", Cell::number(0.7)},
      {"R3", "P2", Cell::number(0.6)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},
      {"R1", "P2", Exp::number(1.6799999999999997)},
      {"R1", "P3", Exp::number(1.9999999999999996)},
      {"R1", "GT", Exp::number(1.504)},
      {"R2", "P1", Exp::error(ErrorCode::Div0)},
      {"R2", "P2", Exp::error(ErrorCode::Div0)},
      {"R2", "P3", Exp::blank()},
      {"R2", "GT", Exp::number(0.0050000000000001155)},
      {"R3", "P1", Exp::blank()},
      {"R3", "P2", Exp::error(ErrorCode::Div0)},
      {"R3", "P3", Exp::error(ErrorCode::Div0)},
      {"R3", "GT", Exp::number(0.605)},
      {"GT", "P1", Exp::error(ErrorCode::Div0)},
      {"GT", "P2", Exp::number(1.3306666666666664)},
      {"GT", "P3", Exp::number(1.1199999999999992)},
      {"GT", "GT", Exp::number(1.0271111111111106)},
  };
  check_case(records, Aggregation::Var, true, expected);
}

// VarP corner sums S1 and S2 over the row partials.
TEST(PivotAggregationTotals, VarpCornerRowOrder) {
  const std::vector<Record> records = {
      {"R1", "P3", Cell::number(2.1)}, {"R1", "P2", Cell::number(3.3)}, {"R1", "P3", Cell::number(0.1)},
      {"R1", "P2", Cell::number(1.1)}, {"R1", "P2", Cell::number(2.1)}, {"R2", "P2", Cell::number(0.6)},
      {"R3", "P3", Cell::number(1.7)}, {"R1", "P2", Cell::number(0.3)}, {"R2", "P1", Cell::number(0.7)},
      {"R3", "P2", Cell::number(0.6)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},
      {"R1", "P2", Exp::number(1.2599999999999998)},
      {"R1", "P3", Exp::number(0.9999999999999998)},
      {"R1", "GT", Exp::number(1.2533333333333334)},
      {"R2", "P1", Exp::number(0.0)},
      {"R2", "P2", Exp::number(0.0)},
      {"R2", "P3", Exp::blank()},
      {"R2", "GT", Exp::number(0.0025000000000000577)},
      {"R3", "P1", Exp::blank()},
      {"R3", "P2", Exp::number(0.0)},
      {"R3", "P3", Exp::number(0.0)},
      {"R3", "GT", Exp::number(0.3025)},
      {"GT", "P1", Exp::number(0.0)},
      {"GT", "P2", Exp::number(1.108888888888889)},
      {"GT", "P3", Exp::number(0.7466666666666664)},
      {"GT", "GT", Exp::number(0.9243999999999997)},
  };
  check_case(records, Aggregation::VarP, true, expected);
}

// Product corner multiplies row totals in display order.
TEST(PivotAggregationTotals, ProductCornerRowOrder) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(5.5)}, {"R2", "P1", Cell::number(1.7)}, {"R1", "P2", Cell::number(1.1)},
      {"R2", "P1", Cell::number(0.7)}, {"R3", "P2", Cell::number(3.3)}, {"R2", "P2", Cell::number(0.3)},
      {"R2", "P1", Cell::number(0.3)}, {"R3", "P1", Cell::number(0.9)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(5.5)},
      {"R1", "P2", Exp::number(1.1)},
      {"R1", "GT", Exp::number(6.050000000000001)},
      {"R2", "P1", Exp::number(0.357)},
      {"R2", "P2", Exp::number(0.3)},
      {"R2", "GT", Exp::number(0.10709999999999999)},
      {"R3", "P1", Exp::number(0.9)},
      {"R3", "P2", Exp::number(3.3)},
      {"R3", "GT", Exp::number(2.9699999999999998)},
      {"GT", "P1", Exp::number(1.7671499999999996)},
      {"GT", "P2", Exp::number(1.089)},
      {"GT", "GT", Exp::number(1.9244263499999996)},
  };
  check_case(records, Aggregation::Product, true, expected);
}

// Average corner adds row numerators in display order.
TEST(PivotAggregationTotals, AverageCornerRowOrder) {
  const std::vector<Record> records = {
      {"R1", "P2", Cell::number(1.0)}, {"R2", "P2", Cell::number(1e+16)},  {"R1", "P1", Cell::number(0.1)},
      {"R2", "P2", Cell::number(2.0)}, {"R1", "P2", Cell::number(-1e+16)}, {"R3", "P2", Cell::number(1.0)},
      {"R1", "P2", Cell::number(0.1)}, {"R1", "P1", Cell::number(1e-17)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(0.05000000000000001)},
      {"R1", "P2", Exp::number(-3333333333333333.5)},
      {"R1", "GT", Exp::number(-1999999999999999.5)},
      {"R2", "P1", Exp::blank()},
      {"R2", "P2", Exp::number(5000000000000001.0)},
      {"R2", "GT", Exp::number(5000000000000001.0)},
      {"R3", "P1", Exp::blank()},
      {"R3", "P2", Exp::number(1.0)},
      {"R3", "GT", Exp::number(1.0)},
      {"GT", "P1", Exp::number(0.05000000000000001)},
      {"GT", "P2", Exp::number(0.5)},
      {"GT", "GT", Exp::number(0.625)},
  };
  check_case(records, Aggregation::Average, true, expected);
}

// the Min corner is the raw min of row partials while column totals snap.
TEST(PivotAggregationTotals, MinCornerRaw) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(1e-15)},
      {"R1", "P2", Cell::number(1e-16)},
      {"R2", "P1", Cell::number(1.0)},
      {"R2", "P2", Cell::number(2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(1e-15)}, {"R1", "P2", Exp::number(1e-16)}, {"R1", "GT", Exp::number(1e-16)},
      {"R2", "P1", Exp::number(1.0)},   {"R2", "P2", Exp::number(2.0)},   {"R2", "GT", Exp::number(1.0)},
      {"GT", "P1", Exp::number(0.0)},   {"GT", "P2", Exp::number(0.0)},   {"GT", "GT", Exp::number(1e-16)},
  };
  check_case(records, Aggregation::Min, true, expected);
}

// the Max corner is the raw max of row partials.
TEST(PivotAggregationTotals, MaxCornerRaw) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::number(-1e-15)},
      {"R1", "P2", Cell::number(-1e-16)},
      {"R2", "P1", Cell::number(-1.0)},
      {"R2", "P2", Cell::number(-2.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::number(-1e-15)}, {"R1", "P2", Exp::number(-1e-16)}, {"R1", "GT", Exp::number(-1e-16)},
      {"R2", "P1", Exp::number(-1.0)},   {"R2", "P2", Exp::number(-2.0)},   {"R2", "GT", Exp::number(-1.0)},
      {"GT", "P1", Exp::number(0.0)},    {"GT", "P2", Exp::number(0.0)},    {"GT", "GT", Exp::number(-1e-16)},
  };
  check_case(records, Aggregation::Max, true, expected);
}
// A Sum leaf whose records are all blank is blank, and so is a total over blank cells only.
TEST(PivotAggregationTotals, SumBlankLeafAndTotals) {
  const std::vector<Record> records = {
      {"R1", "P1", Cell::blank()},
      {"R1", "P2", Cell::blank()},
      {"R2", "P1", Cell::number(2.0)},
      {"R2", "P2", Cell::number(5.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "P1", Exp::blank()},     {"R1", "P2", Exp::blank()},     {"R1", "GT", Exp::blank()},
      {"R2", "P1", Exp::number(2.0)}, {"R2", "GT", Exp::number(7.0)}, {"GT", "P1", Exp::number(2.0)},
      {"GT", "GT", Exp::number(7.0)},
  };
  check_case(records, Aggregation::Sum, true, expected);
}

// Rows only: a blank-only Sum row is blank and adds nothing to the grand total.
TEST(PivotAggregationTotals, SumBlankRowOnly) {
  const std::vector<Record> records = {
      {"R1", "", Cell::blank()},
      {"R2", "", Cell::number(3.0)},
  };
  const std::vector<Expected> expected = {
      {"R1", "", Exp::blank()},
      {"R2", "", Exp::number(3.0)},
      {"GT", "", Exp::number(3.0)},
  };
  check_case(records, Aggregation::Sum, false, expected);
}

}  // namespace
}  // namespace formulon::pivot
