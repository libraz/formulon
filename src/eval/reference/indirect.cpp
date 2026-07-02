// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the INDIRECT lazy impl. The textual A1 -> rectangle
// decoder lives in `reference/common.cpp`; this TU only owns the final
// dereference step.

#include "eval/reference/indirect.h"

#include <cstddef>
#include <cstdint>
#include <utility>
#include <vector>

#include "eval/dynamic_array/common.h"
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

  // Multi-cell INDIRECT (`=INDIRECT("A1:B2")`) resolves to the full
  // rectangle. We expand it row-major and return a `Value::Array` so the
  // spill committer places it on the sheet when the formula is used
  // directly, and so range-aware consumers (`SUM` / `VLOOKUP` / ...) that
  // route an INDIRECT argument through `resolve_range_arg` navigate the
  // rectangle. A 1x1 rectangle (range syntax that collapses to a single
  // cell, e.g. `INDIRECT("A1:A1")`) returns the scalar cell directly.
  if (indirect.is_range) {
    // Whole-column (`A:A` / `A:C`) and whole-row (`1:1` / `1:3`) text expands
    // against the target sheet's used range: `expand_range` clamps the
    // unbounded axis and reports the concrete shape, so `SUM(INDIRECT("A:A"))`
    // aggregates the populated cells. A sheet with no content in range yields
    // a blank scalar (an empty spill has no cells to place).
    parser::Reference top_left{};
    parser::Reference bottom_right{};
    top_left.sheet = indirect.sheet;
    bottom_right.sheet = indirect.sheet;
    if (indirect.is_full_col) {
      top_left.is_full_col = true;
      top_left.col = indirect.left_col;
      bottom_right.is_full_col = true;
      bottom_right.col = indirect.right_col;
    } else if (indirect.is_full_row) {
      top_left.is_full_row = true;
      top_left.row = indirect.top_row;
      bottom_right.is_full_row = true;
      bottom_right.row = indirect.bottom_row;
    } else {
      top_left.row = indirect.top_row;
      top_left.col = indirect.left_col;
      bottom_right.row = indirect.bottom_row;
      bottom_right.col = indirect.right_col;
    }
    std::uint32_t rows = 0;
    std::uint32_t cols = 0;
    auto expanded = ctx.expand_range(top_left, bottom_right, arena, registry, &rows, &cols);
    if (!expanded) {
      return Value::error(expanded.error());
    }
    std::vector<Value> cells = std::move(expanded.value());
    if (!indirect.is_full_col && !indirect.is_full_row) {
      rows = indirect.bottom_row - indirect.top_row + 1U;
      cols = indirect.right_col - indirect.left_col + 1U;
    }
    if (rows == 0U || cols == 0U) {
      // Empty whole-reference expansion (no content in range): nothing to
      // spill, so collapse to a blank scalar. `SUM`/`COUNT` of this is 0.
      return Value::blank();
    }
    Value* buffer = nullptr;
    ArrayValue* out = dynamic_array::allocate_array_value(rows, cols, arena, buffer);
    if (out == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    const std::size_t total = static_cast<std::size_t>(rows) * cols;
    for (std::size_t i = 0; i < total && i < cells.size(); ++i) {
      buffer[i] = cells[i];
    }
    return Value::array(out);
  }

  parser::Reference ref{};
  ref.sheet = indirect.sheet;
  ref.row = indirect.top_row;
  ref.col = indirect.left_col;
  return ctx.resolve_ref(ref, arena, registry);
}

}  // namespace eval
}  // namespace formulon
