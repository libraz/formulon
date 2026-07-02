// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tree-walker lazy impl for the date1904-sensitive calendar family. See
// `datetime_lazy.h` for the rationale (the workbook date epoch is only
// reachable via `EvalContext`, which the eager calling convention cannot
// carry). The actual calendar math lives in `builtins/datetime.cpp`; this
// file only evaluates the arguments and threads `ctx.date1904()` into the
// shared `DateEntry::impl`.

#include "eval/datetime_lazy.h"

#include <cstdint>
#include <vector>

#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_datetime_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const DateEntry* entry = find_date_entry(call.as_call_name());
  if (entry == nullptr) {
    // Registered names always resolve; a miss can only mean a table drift.
    return Value::error(ErrorCode::Name);
  }
  const std::uint32_t arity = call.as_call_arity();
  if (arity < entry->min_arity || arity > entry->max_arity) {
    return Value::error(ErrorCode::Value);
  }
  // Calendar functions are scalar-only and propagate the left-most argument
  // error (they never opt out of `propagate_errors`), matching the eager
  // dispatcher's pre-evaluation contract.
  std::vector<Value> args;
  args.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    Value v = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    args.push_back(v);
  }
  return entry->impl(args.empty() ? nullptr : args.data(), arity, arena, ctx.date1904());
}

}  // namespace eval
}  // namespace formulon
