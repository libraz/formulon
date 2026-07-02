// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Top-level array-broadcasting helpers for the tree-walk evaluator's
// BinaryOp / UnaryOp cases. Split out of `tree_walker.cpp` to keep the
// recursive walker compile unit small; see `tree_walker/broadcast.h`
// for the public contract.
//
// `broadcast_binop` is the single implementation of Excel 365 array
// broadcasting for the tree-walk evaluator. It takes *already evaluated*
// `Value`s so both the top-level BinaryOp dispatch and `shape_ops_lazy.cpp`'s
// array-context path (`eval_binop_array_ctx`, which delegates here) share one
// rule set. Shape rules follow Excel 365 dynamic arrays:
//
//   * The result is `max(r1, r2) x max(c1, c2)`.
//   * A dimension of size 1 broadcasts to the other operand's size (this is
//     what makes the outer product `{1;2;3}*{10,20}` -> 3x2 and the mixed
//     forms `RxC op Rx1` / `RxC op 1xC` work).
//   * A dimension where both operands are > 1 but unequal does NOT error:
//     Excel extends to the larger size and fills the cells an operand cannot
//     supply with `#N/A` (`{1,2,3}+{1,2}` -> `{2,4,#N/A}`).
//
// Per-cell error short-circuit: if either contributing operand cell is an
// Error, that error is written verbatim into the result cell (left-most wins
// via the lhs-first check).

#include "eval/tree_walker/broadcast.h"

#include <cstddef>
#include <cstdint>

#include "eval/coerce.h"
#include "eval/scalar_ops.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

struct ArrayView {
  std::uint32_t rows;
  std::uint32_t cols;
  const Value* cells;
};

// Resolves `v` to an ArrayView. For an Array the view aliases the existing
// cells buffer (no copy). For a scalar the caller-supplied 1-element backing
// slot `scalar_slot` is populated and aliased. Lifetime: the view is valid
// as long as either the source Array or `scalar_slot` outlives it.
ArrayView as_array_view(const Value& v, Value* scalar_slot) {
  if (v.is_array()) {
    const ArrayValue* a = v.as_array();
    return {a->rows, a->cols, a->cells};
  }
  *scalar_slot = v;
  return {1U, 1U, scalar_slot};
}

// Fetches the operand cell contributing to output position `(r, c)` under
// Excel broadcasting, or `nullptr` when the operand cannot supply that
// position (the axis is larger than this operand's non-1 extent -> `#N/A`).
// A size-1 axis always reads index 0 (broadcast); a size-N axis reads its own
// index when in range.
const Value* broadcast_cell(const ArrayView& v, std::uint32_t r, std::uint32_t c) {
  const std::uint32_t ri = v.rows == 1U ? 0U : r;
  const std::uint32_t ci = v.cols == 1U ? 0U : c;
  if (ri >= v.rows || ci >= v.cols) {
    return nullptr;
  }
  return &v.cells[static_cast<std::size_t>(ri) * v.cols + ci];
}

}  // namespace

Value apply_binop_per_cell(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena) {
  if (lhs.is_error()) {
    return lhs;
  }
  if (rhs.is_error()) {
    return rhs;
  }
  switch (op) {
    case parser::BinOp::Add:
    case parser::BinOp::Sub:
    case parser::BinOp::Mul:
    case parser::BinOp::Div:
    case parser::BinOp::Pow: {
      auto ln = coerce_to_number(lhs);
      if (!ln) {
        return Value::error(ln.error());
      }
      auto rn = coerce_to_number(rhs);
      if (!rn) {
        return Value::error(rn.error());
      }
      return apply_arithmetic(op, ln.value(), rn.value());
    }
    case parser::BinOp::Concat:
      return apply_concat(lhs, rhs, arena);
    case parser::BinOp::Eq:
    case parser::BinOp::NotEq:
    case parser::BinOp::Lt:
    case parser::BinOp::LtEq:
    case parser::BinOp::Gt:
    case parser::BinOp::GtEq:
      return apply_comparison(op, lhs, rhs);
  }
  return Value::error(ErrorCode::Value);
}

Value broadcast_binop(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena) {
  Value l_slot = Value::blank();
  Value r_slot = Value::blank();
  const ArrayView la = as_array_view(lhs, &l_slot);
  const ArrayView ra = as_array_view(rhs, &r_slot);

  // Excel 365 broadcast shape: each axis extends to the larger operand's
  // size. A size-1 axis broadcasts; a size-mismatch on a non-1 axis pads the
  // shortfall with #N/A (handled per-cell via `broadcast_cell`).
  const std::uint32_t out_rows = la.rows > ra.rows ? la.rows : ra.rows;
  const std::uint32_t out_cols = la.cols > ra.cols ? la.cols : ra.cols;

  const std::size_t n = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buf = arena.create_array<Value>(n);
  if (buf == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::size_t i = 0;
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c, ++i) {
      const Value* lv = broadcast_cell(la, r, c);
      const Value* rv = broadcast_cell(ra, r, c);
      if (lv == nullptr || rv == nullptr) {
        // One operand cannot supply this position: Excel fills #N/A.
        buf[i] = Value::error(ErrorCode::NA);
        continue;
      }
      buf[i] = apply_binop_per_cell(op, *lv, *rv, arena);
    }
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buf;
  return Value::array(out);
}

Value broadcast_unary(parser::UnaryOp op, const Value& operand, Arena& arena) {
  if (!operand.is_array()) {
    return apply_unary(op, operand);
  }
  const ArrayValue* in = operand.as_array();
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  Value* buf = arena.create_array<Value>(n);
  if (buf == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = in->cells[i];
    buf[i] = cell.is_error() ? cell : apply_unary(op, cell);
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = in->rows;
  out->cols = in->cols;
  out->cells = buf;
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
