// A scalar function that propagates errors reports its first failing
// argument in check order. Every registered one is called with 1 in each
// slot (every slot it has, two for a variadic one), then with one error and
// with two different errors in place of those 1s: the earlier error is the
// result, the later one for a function that checks its arguments last to
// first.

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "eval/tree_walker_lazy_table.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

void CollectName(std::string_view name, void* ctx) {
  static_cast<std::vector<std::string>*>(ctx)->emplace_back(name);
}

// The call with `first` in slot `i`, `second` in slot `j` (unless j == arity)
// and 1 everywhere else.
std::string CallText(std::string_view name, std::uint32_t arity, std::uint32_t i, std::uint32_t j) {
  std::string text = "=" + std::string(name) + "(";
  for (std::uint32_t k = 0; k < arity; ++k) {
    if (k != 0U) {
      text += ",";
    }
    text += k == i ? "H1" : k == j ? "H2" : "1";
  }
  return text + ")";
}

TEST(ErrorArgumentOrder, ScalarFunctionsReportTheFirstFailingArgument) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 7U, Value::error(ErrorCode::Div0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 7U, Value::error(ErrorCode::Num))));

  const FunctionRegistry& registry = default_registry();
  std::vector<std::string> names;
  registry.for_each_name(&CollectName, &names);

  std::string mismatches;
  std::size_t calls = 0;
  for (const std::string& name : names) {
    const FunctionDef* def = registry.lookup(name);
    if (def == nullptr || !def->propagate_errors || def->accepts_ranges || find_lazy_impl(name) != nullptr ||
        def->max_arity == 0U) {
      continue;
    }
    const std::uint32_t arity = def->max_arity != kVariadic ? def->max_arity
                                : def->min_arity > 2U       ? def->min_arity
                                                            : 2U;
    for (std::uint32_t i = 0; i < arity; ++i) {
      for (std::uint32_t j = i + 1U; j <= arity; ++j) {
        const std::string text = CallText(name, arity, i, j);
        const ErrorCode want = j < arity && def->checks_args_last_first ? ErrorCode::Num : ErrorCode::Div0;
        Arena arena;
        const Value got = evaluate_formula_text_array(wb, wb.sheet(0U), 4U, 0U, text, arena, registry);
        ++calls;
        if (!got.is_error() || got.as_error() != want) {
          mismatches +=
              text + " -> " + (got.is_error() ? std::string(display_name(got.as_error())) : "non-error") + "\n";
        }
      }
    }
  }
  EXPECT_GT(calls, 0U);
  EXPECT_TRUE(mismatches.empty()) << mismatches;
}

// The functions measured to check their arguments last to first, and no
// others, carry the flag the scalar test above reads.
TEST(ErrorArgumentOrder, ArgumentsCheckedLastToFirst) {
  const FunctionRegistry& registry = default_registry();
  std::vector<std::string> names;
  registry.for_each_name(&CollectName, &names);
  std::vector<std::string> flagged;
  for (const std::string& name : names) {
    const FunctionDef* def = registry.lookup(name);
    if (def != nullptr && def->checks_args_last_first) {
      flagged.push_back(name);
    }
  }
  std::sort(flagged.begin(), flagged.end());
  const std::vector<std::string> want = {"BIN2HEX", "BIN2OCT",     "DEC2BIN",    "DEC2HEX", "DEC2OCT",
                                         "DELTA",   "GESTEP",      "HEX2BIN",    "HEX2OCT", "OCT2BIN",
                                         "OCT2HEX", "RANDBETWEEN", "VALUETOTEXT"};
  EXPECT_EQ(flagged, want);
}

// A range-aware function with no range filter reads its range cells in
// argument order: a cell it cannot read ahead of an error cell is reported.
TEST(ErrorArgumentOrder, UnfilteredRangeFunctionsReadCellsInArgumentOrder) {
  const FunctionRegistry& registry = default_registry();
  std::vector<std::string> names;
  registry.for_each_name(&CollectName, &names);
  std::vector<std::string> in_order;
  for (const std::string& name : names) {
    if (reads_range_errors_in_order(*registry.lookup(name))) {
      in_order.push_back(name);
    }
  }
  std::sort(in_order.begin(), in_order.end());
  EXPECT_EQ(in_order, (std::vector<std::string>{"CONCAT", "GCD", "IMPRODUCT", "IMSUM", "LCM", "MULTINOMIAL"}));

  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::text("x"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 1U, Value::error(ErrorCode::Div0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 2U, Value::error(ErrorCode::Num))));
  const std::vector<std::pair<std::string, ErrorCode>> table = {
      {"=GCD(A1:B1)", ErrorCode::Value},
      {"=GCD(B1:C1)", ErrorCode::Div0},
      {"=GCD(C1,A1:B1)", ErrorCode::Num},
      {"=GCD(A1,B1:C1)", ErrorCode::Value},
      {"=LCM(A1:B1)", ErrorCode::Value},
      {"=LCM(A1,B1:C1)", ErrorCode::Value},
      {"=MULTINOMIAL(A1:B1)", ErrorCode::Value},
      {"=MULTINOMIAL(B1:C1)", ErrorCode::Div0},
      {"=IMSUM(A1:B1)", ErrorCode::Num},
      {"=IMSUM(B1:C1)", ErrorCode::Div0},
      {"=IMPRODUCT(A1:B1)", ErrorCode::Num},
      {"=IMPRODUCT(C1,B1:C1)", ErrorCode::Num},
      {"=CONCAT(B1:C1)", ErrorCode::Div0},
      {"=CONCAT(LAMBDA(x,x),B1:C1)", ErrorCode::Value},
      {"=GCD((B1,C1))", ErrorCode::Value},
      {"=CONCAT(C1,(A1:A1,A1:A1))", ErrorCode::Num},
      {"=CONCAT((A1:A1,A1:A1),C1)", ErrorCode::Value},
      {"=TEXTJOIN(\",\",TRUE,(A1,A1),C1)", ErrorCode::Value},
  };
  for (const auto& [formula, want] : table) {
    Arena arena;
    const Value got = evaluate_formula_text_array(wb, wb.sheet(0U), 4U, 0U, formula, arena, registry);
    ASSERT_TRUE(got.is_error()) << formula;
    EXPECT_EQ(got.as_error(), want) << formula;
  }
}

// A function taking an array beside single values reports an error held in a
// cell of the array only once those values have coerced; an error written in
// the array slot keeps its place in argument order.
TEST(ErrorArgumentOrder, ArrayCellErrorsRankAfterSingleValues) {
  const FunctionRegistry& registry = default_registry();
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::text("x"))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 7U, Value::error(ErrorCode::Div0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 7U, Value::error(ErrorCode::Num))));
  const std::vector<std::pair<std::string, ErrorCode>> table = {
      {"=LARGE(H1,H2)", ErrorCode::Num},          {"=SMALL(H1:H2,2)", ErrorCode::Div0},
      {"=PERCENTILE.INC(H1,H2)", ErrorCode::Num}, {"=QUARTILE.EXC(H1,H2)", ErrorCode::Num},
      {"=PERCENTRANK(H1,H2)", ErrorCode::Num},    {"=PERCENTRANK.EXC(H1,1,H2)", ErrorCode::Num},
      {"=TRIMMEAN(H1,H2)", ErrorCode::Num},       {"=TRIMMEAN(H1,A1)", ErrorCode::Value},
      {"=TRIMMEAN(A1,H1)", ErrorCode::Div0},      {"=TRIMMEAN(\"x\",1/0)", ErrorCode::Value},
      {"=TRIMMEAN(1/0,\"x\")", ErrorCode::Div0},  {"=IRR(H1,H2)", ErrorCode::Num},
      {"=MIRR(H1,H2,1)", ErrorCode::Num},         {"=MIRR(H1,1,H2)", ErrorCode::Num},
      {"=TEXTJOIN(H1,H2,1)", ErrorCode::Num},     {"=TEXTJOIN(1/0,H2,1)", ErrorCode::Div0},
  };
  for (const auto& [formula, want] : table) {
    Arena arena;
    const Value got = evaluate_formula_text_array(wb, wb.sheet(0U), 4U, 0U, formula, arena, registry);
    ASSERT_TRUE(got.is_error()) << formula;
    EXPECT_EQ(got.as_error(), want) << formula;
  }
}

// A one-cell data argument holding an error is reported as that error, ahead
// of the other arguments and without a spill.
TEST(ErrorArgumentOrder, LoneErrorCellArgumentIsTheResult) {
  const FunctionRegistry& registry = default_registry();
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 0U, Value::error(ErrorCode::Div0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 1U, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 2U, 1U, Value::number(3.0))));
  for (const char* formula : {"=EXPAND(A3,2,2)", "=EXPAND(A3,1,3)", "=WRAPCOLS(A3,2)", "=GROUPBY(B2,A3,SUM)",
                              "=PIVOTBY(B2,B3,A3,SUM)", "=T.TEST(A3,B2:B3,2,2)", "=SORT(A3,5)", "=UNIQUE(A3)"}) {
    Arena arena;
    const Value got = evaluate_formula_text_array(wb, wb.sheet(0U), 4U, 0U, formula, arena, registry);
    ASSERT_TRUE(got.is_error()) << formula;
    EXPECT_EQ(got.as_error(), ErrorCode::Div0) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
