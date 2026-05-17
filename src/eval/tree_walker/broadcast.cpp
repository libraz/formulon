// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Top-level array-broadcasting helpers for the tree-walk evaluator's
// BinaryOp / UnaryOp cases. Split out of `tree_walker.cpp` to keep the
// recursive walker compile unit small; see `tree_walker/broadcast.h`
// for the public contract.
//
// `eval_binop_array_ctx` in `shape_ops_lazy.cpp` re-walks the AST under
// SUMPRODUCT and other array-context callers; the helpers below take
// *already evaluated* `Value`s so the top-level dispatch can broadcast
// over arrays produced by SpillRef / TRANSPOSE / SEQUENCE without a
// second AST walk. Shape rules mirror `eval_binop_array_ctx` exactly:
//
//   * Both operands 1x1                 -> 1x1 result
//   * Either operand 1x1, other R x C   -> R x C result, scalar broadcasts
//   * Otherwise dimensions must match   -> mismatch surfaces scalar #VALUE!
//     (matches Mac Excel's whole-expression short-circuit; it does NOT
//     spill a sea of #VALUE! cells)
//
// Per-cell error short-circuit: if either operand cell is an Error, that
// error is written verbatim into the result cell (left-most wins on ties
// via the lhs-first check in the loop).

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

  std::uint32_t out_rows = la.rows;
  std::uint32_t out_cols = la.cols;
  bool l_broadcast = false;
  bool r_broadcast = false;
  if (la.rows == 1U && la.cols == 1U) {
    out_rows = ra.rows;
    out_cols = ra.cols;
    l_broadcast = true;
  } else if (ra.rows == 1U && ra.cols == 1U) {
    r_broadcast = true;
  } else if (la.rows != ra.rows || la.cols != ra.cols) {
    return Value::error(ErrorCode::Value);
  }

  const std::size_t n = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buf = arena.create_array<Value>(n);
  if (buf == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    const Value& lv = l_broadcast ? la.cells[0] : la.cells[i];
    const Value& rv = r_broadcast ? ra.cells[0] : ra.cells[i];
    buf[i] = apply_binop_per_cell(op, lv, rv, arena);
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
