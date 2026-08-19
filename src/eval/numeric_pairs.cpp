//
// See `eval/numeric_pairs.h` for the contract.

#include "eval/numeric_pairs.h"

#include <cstddef>
#include <utility>

#include "eval/range_args.h"
#include "utils/error.h"

namespace formulon::eval {

std::variant<Value, NumericPairs> collect_numeric_pairs(const parser::AstNode& first_arg,
                                                        const parser::AstNode& second_arg, Arena& arena,
                                                        const FunctionRegistry& registry, const EvalContext& ctx,
                                                        bool allow_transpose) {
  auto first_resolved = resolve_array_arg_na(first_arg, arena, registry, ctx);
  if (!first_resolved) {
    return Value{Value::error(first_resolved.error())};
  }
  RangeResult first_arr = std::move(first_resolved.value());
  auto second_resolved = resolve_array_arg_na(second_arg, arena, registry, ctx);
  if (!second_resolved) {
    return Value{Value::error(second_resolved.error())};
  }
  RangeResult second_arr = std::move(second_resolved.value());

  // Shape mismatch is `#N/A` here, unlike SUMPRODUCT's `#VALUE!`.
  const bool shape_ok = allow_transpose ? (static_cast<std::size_t>(first_arr.rows) * first_arr.cols ==
                                           static_cast<std::size_t>(second_arr.rows) * second_arr.cols)
                                        : (first_arr.rows == second_arr.rows && first_arr.cols == second_arr.cols);
  if (!shape_ok) {
    return Value{Value::error(ErrorCode::NA)};
  }

  // Error propagation runs over every cell in both arrays, even cells that
  // the numeric-pair rule would drop. Scan the first argument first so the
  // leftmost-argument error wins, matching Excel's precedence.
  for (const Value& cell : first_arr.cells) {
    if (cell.is_error()) {
      return cell;
    }
  }
  for (const Value& cell : second_arr.cells) {
    if (cell.is_error()) {
      return cell;
    }
  }

  NumericPairs pairs;
  const std::size_t count = first_arr.cells.size();
  pairs.first.reserve(count);
  pairs.second.reserve(count);
  for (std::size_t i = 0; i < count; ++i) {
    const Value& lhs = first_arr.cells[i];
    const Value& rhs = second_arr.cells[i];
    if (!lhs.is_number() || !rhs.is_number()) {
      continue;
    }
    pairs.first.push_back(lhs.as_number());
    pairs.second.push_back(rhs.as_number());
  }
  return pairs;
}

}  // namespace formulon::eval
