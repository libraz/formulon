// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the INDIRECT lazy impl. The textual A1 -> rectangle
// decoder lives in `reference/common.cpp`; this TU only owns the final
// dereference step.

#include "eval/reference/indirect.h"

#include "eval/eval_context.h"
#include "eval/reference/common.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_indirect_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  refs_internal::IndirectReference indirect{};
  ErrorCode err = ErrorCode::Value;
  if (!refs_internal::resolve_indirect_reference(call, arena, registry, ctx, &indirect, &err)) {
    return Value::error(err);
  }
  // Range INDIRECT is deferred until `Value::Array` lands. `=SUM(INDIRECT(
  // "A1:B2"))` therefore surfaces as `#REF!` in scalar context today.
  if (indirect.range_syntax) {
    return Value::error(ErrorCode::Ref);
  }

  parser::Reference ref{};
  ref.sheet = indirect.sheet;
  ref.row = indirect.top_row;
  ref.col = indirect.left_col;
  return ctx.resolve_ref(ref, arena, registry);
}

}  // namespace eval
}  // namespace formulon
