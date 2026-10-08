// Conditional-format and data-validation formulas whose result is an array use its
// first element for every target cell, for dense and whole-axis results alike.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "cf/validation_eval.h"
#include "eval/adhoc_eval.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

constexpr std::uint32_t kColD = 3;
constexpr std::uint32_t kColE = 4;
constexpr std::uint32_t kRows = 6;

struct Grid {
  Workbook wb = Workbook::create_empty();
  std::size_t sheet_index = 0;

  Grid() { sheet_index = wb.add_sheet("S"); }

  void SetB(std::initializer_list<double> values) {
    std::uint32_t row = 0;
    for (double v : values) {
      wb.set_cell_value(sheet_index, row++, 1, Value::number(v));
    }
  }

  /// True when the CF expression anchored at D1 matches the target cell in `col`.
  bool CfMatches(const std::string& formula, std::uint32_t row, std::uint32_t col = kColD, Arena* arena_out = nullptr) {
    const Sheet& sheet = wb.sheet(sheet_index);
    Arena local;
    Arena& arena = arena_out != nullptr ? *arena_out : local;
    eval::EvalState state;
    eval::EvalContext eval_ctx(wb, sheet, state);
    cf::CFRule rule;
    rule.type = cf::RuleType::Expression;
    rule.formula1 = formula;
    cf::CFEvalContext ctx;
    ctx.anchor = CellAddress{0U, col};
    ctx.target = CellAddress{row, col};
    ctx.arena = &arena;
    ctx.registry = &eval::default_registry();
    ctx.eval_ctx = &eval_ctx;
    return cf::match_rule(rule, Value::blank(), ctx);
  }

  bool DvValid(const std::string& formula, std::uint32_t row) {
    DataValidation dv;
    dv.ranges.push_back(MergeRange{0, kColE, kRows - 1, kColE});
    dv.type = 7;  // custom
    dv.allow_blank = false;
    dv.formula1 = formula;
    wb.sheet(sheet_index).mutable_validations().clear();
    wb.sheet(sheet_index).mutable_validations().push_back(std::move(dv));
    auto r = validate_value(wb, wb.sheet(sheet_index), row, kColE, Value::number(1.0));
    EXPECT_TRUE(static_cast<bool>(r));
    return r.value().valid;
  }
};

TEST(CfWholeAxis, NoCellMatchesWhenFirstElementFails) {
  Grid g;
  g.SetB({10, 20, 30, 40, 50});
  for (std::uint32_t row = 0; row < kRows; ++row) {
    EXPECT_FALSE(g.CfMatches("$B:$B>25", row)) << row;
    EXPECT_FALSE(g.CfMatches("$B:$B+0>25", row, 6)) << row;
    EXPECT_FALSE(g.DvValid("$B:$B>25", row)) << row;
  }
}

TEST(CfWholeAxis, EveryCellMatchesWhenFirstElementPasses) {
  Grid g;
  g.SetB({100, 10, 10, 40, 10});
  for (std::uint32_t row = 0; row < kRows; ++row) {
    EXPECT_TRUE(g.CfMatches("$B:$B>25", row)) << row;
    EXPECT_TRUE(g.CfMatches("$B:$B+0>25", row, 6)) << row;
    EXPECT_TRUE(g.DvValid("$B:$B>25", row)) << row;
  }
}

TEST(CfWholeAxis, ReducedWholeAxisAggregateMatchesEveryCell) {
  Grid g;
  g.SetB({100, 10, 10, 40, 10});
  for (std::uint32_t row = 0; row < kRows; ++row) {
    EXPECT_TRUE(g.CfMatches("SUM($B:$B*1)>100", row)) << row;
  }
}

// Inferred from the whole-axis measurement, not measured: a bounded array result is also
// decided by its first element.
TEST(CfWholeAxis, BoundedArrayResultDecidedByFirstElement) {
  Grid g;
  g.SetB({10, 20, 30, 40, 50});
  for (std::uint32_t row = 0; row < kRows; ++row) {
    EXPECT_FALSE(g.CfMatches("$B$1:$B$5>25", row)) << row;
    EXPECT_FALSE(g.DvValid("$B$1:$B$5>25", row)) << row;
  }
  g.SetB({100, 10, 10, 40, 10});
  for (std::uint32_t row = 0; row < kRows; ++row) {
    EXPECT_TRUE(g.CfMatches("$B$1:$B$5>25", row)) << row;
    EXPECT_TRUE(g.DvValid("$B$1:$B$5>25", row)) << row;
  }
}

TEST(CfWholeAxis, ComparisonFormsStayUnderOneMiBOfArena) {
  Grid g;
  g.SetB({100, 10, 10, 40, 10});
  for (const char* formula : {"$B:$B>25", "$B:$B+0>25"}) {
    Arena arena;
    EXPECT_TRUE(g.CfMatches(formula, 2, kColD, &arena));
    EXPECT_LT(arena.bytes_allocated(), static_cast<std::size_t>(1) << 20) << formula << ": " << arena.bytes_allocated();
  }
}

// The ad-hoc rule probe takes the same first element: an array the target cell
// could not spill into still decides the rule, and a whole column is not expanded.
TEST(CfWholeAxis, AdhocProbeAgreesWithRuleEvaluation) {
  Grid g;
  g.SetB({100, 10, 10, 40, 10});
  g.wb.set_cell_value(g.sheet_index, 2, kColD, Value::number(1.0));
  const Sheet& sheet = g.wb.sheet(g.sheet_index);
  for (const char* formula : {"$B:$B>25", "$B$1:$B$5>25"}) {
    Arena arena;
    EXPECT_TRUE(eval::evaluate_cf_formula(g.wb, sheet, 0U, kColD, 0U, kColD, formula, arena, eval::default_registry()))
        << formula;
    EXPECT_TRUE(g.CfMatches(formula, 0U)) << formula;
    EXPECT_LT(arena.bytes_allocated(), static_cast<std::size_t>(1) << 20) << formula << ": " << arena.bytes_allocated();
  }
}

}  // namespace
}  // namespace formulon
