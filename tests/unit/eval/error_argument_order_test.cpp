// A scalar function that propagates errors reports its first failing
// argument in check order. Every registered one is called with 1 in each
// slot (every slot it has, two for a variadic one), then with one error and
// with two different errors in place of those 1s: the earlier error is the
// result, the later one for a function that checks its arguments last to
// first.

#include <cstdint>
#include <string>
#include <string_view>
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

}  // namespace
}  // namespace eval
}  // namespace formulon
