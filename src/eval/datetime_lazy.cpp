//
// Tree-walker lazy impl for the date1904-sensitive calendar family. See
// `datetime_lazy.h` for the rationale (the workbook date epoch is only
// reachable via `EvalContext`, which the eager calling convention cannot
// carry). The actual calendar math lives in `builtins/datetime.cpp`; this
// file only evaluates the arguments and threads `ctx.date1904()` into the
// shared `DateEntry::impl`.

#include "eval/datetime_lazy.h"

#include <cstdint>
#include <optional>
#include <vector>

#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/scalar_args.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value invoke_date_entry(const DateEntry& entry, const Value* args, std::uint32_t arity, Arena& arena, bool date1904,
                        const date_time::CivilTime& now) {
  // The clock builtins are zero-arity, so none of the broadcasting below
  // applies to them; they only need the reading the caller resolved.
  if (entry.clock_impl != nullptr) {
    return entry.clock_impl(now, date1904);
  }
  // Threads `now` alongside the (possibly per-cell) argument slice for an
  // entry that needs both, e.g. DATEVALUE's year-less-date fallback year.
  const auto invoke_impl = [&](const Value* a, std::uint32_t n) -> Value {
    if (entry.clock_aware_impl != nullptr) {
      return entry.clock_aware_impl(a, n, arena, date1904, now);
    }
    return entry.impl(a, n, arena, date1904);
  };
  // An error argument the impl does not read still wins over its result.
  const auto invoke_reporting = [&](const Value* a, std::uint32_t n) -> Value {
    return report_error_arg(invoke_impl(a, n), a, n, false);
  };
  return lift_over_array_args(args, arity, arena, invoke_reporting);
}

Value eval_datetime_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const DateEntry* entry = find_date_entry(strip_future_prefix(call.as_call_name()));
  if (entry == nullptr) {
    // Registered names always resolve; a miss can only mean a table drift.
    return Value::error(ErrorCode::Name);
  }
  const std::uint32_t arity = call.as_call_arity();
  if (arity < entry->min_arity || arity > entry->max_arity) {
    return Value::error(ErrorCode::Value);
  }
  // Calendar functions report the first failing argument, as the eager
  // dispatcher does; `invoke_date_entry` lifts over arrays.
  std::vector<Value> args;
  std::optional<ErrorOriginSnapshot> error_arg_origin;
  eval_scalar_call_args(call, entry->min_arity, entry->max_arity, entry->analysis_toolpak, 0U, arena, registry, ctx,
                        args, error_arg_origin);
  const auto arity_n = static_cast<std::uint32_t>(args.size());
  const Value result =
      invoke_date_entry(*entry, args.empty() ? nullptr : args.data(), arity_n, arena, ctx.date1904(), ctx.wall_clock());
  restore_reported_origin(result, args.data(), arity_n, false, error_arg_origin);
  return result;
}

}  // namespace eval
}  // namespace formulon
