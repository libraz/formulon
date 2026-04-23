// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `resolve_range_arg`. See `range_args.h` for the
// public contract.

#include "eval/range_args.h"

#include <cstdint>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

bool resolve_range_arg(const parser::AstNode& arg_node, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                       std::uint32_t* out_rows, std::uint32_t* out_cols) {
  if (arg_node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg_node.as_range_lhs();
    const parser::AstNode& rhs_ast = arg_node.as_range_rhs();
    if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
      *out_err_code = ErrorCode::Ref;
      return false;
    }
    const parser::Reference& lhs_ref = lhs_ast.as_ref();
    const parser::Reference& rhs_ref = rhs_ast.as_ref();
    auto expanded = ctx.expand_range(lhs_ref, rhs_ref, arena, registry);
    if (!expanded) {
      *out_err_code = expanded.error();
      return false;
    }
    *out_cells = std::move(expanded.value());
    if (out_rows != nullptr || out_cols != nullptr) {
      // `expand_range` normalises endpoint ordering (A3:A1 == A1:A3) so
      // mirror that here. Full-col/full-row would have already failed the
      // expansion with #VALUE!, so we can safely take the absolute span.
      const std::uint32_t r_lo = lhs_ref.row < rhs_ref.row ? lhs_ref.row : rhs_ref.row;
      const std::uint32_t r_hi = lhs_ref.row < rhs_ref.row ? rhs_ref.row : lhs_ref.row;
      const std::uint32_t c_lo = lhs_ref.col < rhs_ref.col ? lhs_ref.col : rhs_ref.col;
      const std::uint32_t c_hi = lhs_ref.col < rhs_ref.col ? rhs_ref.col : lhs_ref.col;
      if (out_rows != nullptr) {
        *out_rows = r_hi - r_lo + 1U;
      }
      if (out_cols != nullptr) {
        *out_cols = c_hi - c_lo + 1U;
      }
    }
    return true;
  }
  if (arg_node.kind() == parser::NodeKind::Ref) {
    // Single-cell Ref: treat as a 1-element range so COUNTIF(A1, ">0") is
    // well-defined. Error / blank surface via `resolve_ref` as a Value and
    // are forwarded unchanged; the matcher handles them correctly.
    out_cells->clear();
    out_cells->push_back(ctx.resolve_ref(arg_node.as_ref(), arena, registry));
    if (out_rows != nullptr) {
      *out_rows = 1U;
    }
    if (out_cols != nullptr) {
      *out_cols = 1U;
    }
    return true;
  }
  // Any other shape — literal, call, arithmetic, array — is not a range
  // and Excel rejects it with #VALUE!.
  *out_err_code = ErrorCode::Value;
  return false;
}

}  // namespace eval
}  // namespace formulon
