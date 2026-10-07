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

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/scalar_ops.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

ArrayView as_array_view(const Value& v, Value* scalar_slot) {
  if (v.is_array()) {
    const ArrayValue* a = v.as_array();
    return {a->rows, a->cols, a->cells};
  }
  *scalar_slot = v;
  return {1U, 1U, scalar_slot};
}

ArrayView as_array_view(const Shaped& s, Value* scalar_slot) {
  if (s.tail_array == nullptr) {
    return as_array_view(s.value, scalar_slot);
  }
  const TailArray& t = *s.tail_array;
  return {t.rows, t.cols, t.cells, t.head, t.tail, t.axis};
}

const Value* broadcast_cell(const ArrayView& v, std::uint32_t r, std::uint32_t c) {
  const std::uint32_t ri = v.rows == 1U ? 0U : r;
  const std::uint32_t ci = v.cols == 1U ? 0U : c;
  if (ri >= v.rows || ci >= v.cols) {
    return nullptr;
  }
  if (v.tail != nullptr) {
    if (v.axis == TailAxis::kRows) {
      return ri < v.head ? &v.cells[static_cast<std::size_t>(ri) * v.cols + ci] : &v.tail[ci];
    }
    return ci < v.head ? &v.cells[static_cast<std::size_t>(ri) * v.head + ci] : &v.tail[ri];
  }
  return &v.cells[static_cast<std::size_t>(ri) * v.cols + ci];
}

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

  Value* buf = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, buf, kMaxDerivedArrayCells);
  if (out == nullptr) {
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
  return Value::array(out);
}

Value broadcast_unary(parser::UnaryOp op, const Value& operand, Arena& arena) {
  if (!operand.is_array()) {
    return apply_unary(op, operand);
  }
  const ArrayValue* in = operand.as_array();
  Value* buf = nullptr;
  ArrayValue* out = allocate_array_value(in->rows, in->cols, arena, buf, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = in->cells[i];
    buf[i] = cell.is_error() ? cell : apply_unary(op, cell);
  }
  return Value::array(out);
}

namespace {

std::uint32_t axis_extent(const ArrayView& v, TailAxis axis) {
  return axis == TailAxis::kRows ? v.rows : v.cols;
}

// Builds a result of declared size `rows` x `cols` whose position (r, c) is
// `f(r, c)`, computing `head` dense positions along `axis` and the tail
// row / column once at index `head`. A head reaching the full extent yields a
// plain dense array. Positions at and past `head` must all evaluate alike.
template <class F>
Shaped build_tail_result(std::uint32_t rows, std::uint32_t cols, std::uint32_t head, TailAxis axis, Arena& arena,
                         F&& f) {
  const bool by_rows = axis == TailAxis::kRows;
  if (head >= (by_rows ? rows : cols)) {
    Value* buf = nullptr;
    ArrayValue* out = allocate_array_value(rows, cols, arena, buf, kMaxDerivedArrayCells);
    if (out == nullptr) {
      return Shaped{Value::error(ErrorCode::Num), nullptr};
    }
    std::size_t i = 0;
    for (std::uint32_t r = 0; r < rows; ++r) {
      for (std::uint32_t c = 0; c < cols; ++c, ++i) {
        buf[i] = f(r, c);
      }
    }
    return Shaped{Value::array(out), nullptr};
  }
  const std::uint64_t head_cells = static_cast<std::uint64_t>(head) * (by_rows ? cols : rows);
  const std::size_t tail_len = by_rows ? cols : rows;
  if (head_cells > kMaxDerivedArrayCells) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  Value* cells = head_cells == 0 ? nullptr : arena.create_array<Value>(static_cast<std::size_t>(head_cells));
  Value* tail = arena.create_array<Value>(tail_len);
  if ((head_cells != 0 && cells == nullptr) || tail == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  std::size_t i = 0;
  for (std::uint32_t r = 0; r < (by_rows ? head : rows); ++r) {
    for (std::uint32_t c = 0; c < (by_rows ? cols : head); ++c, ++i) {
      cells[i] = f(r, c);
    }
  }
  for (std::size_t k = 0; k < tail_len; ++k) {
    tail[k] = by_rows ? f(head, static_cast<std::uint32_t>(k)) : f(static_cast<std::uint32_t>(k), head);
  }
  const TailArray* ta = make_tail_array(arena, rows, cols, head, axis, cells, tail, false);
  if (ta == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  Shaped s;
  s.tail_array = ta;
  return s;
}

}  // namespace

Shaped broadcast_binop(parser::BinOp op, const Shaped& lhs, const Shaped& rhs, Arena& arena) {
  // A scalar error operand short-circuits (left-most first) before any
  // broadcasting, as the walker does for plain values.
  if (lhs.tail_array == nullptr && lhs.value.is_error()) {
    return lhs;
  }
  if (rhs.tail_array == nullptr && rhs.value.is_error()) {
    return rhs;
  }
  if (lhs.tail_array == nullptr && rhs.tail_array == nullptr) {
    if (lhs.value.is_array() || rhs.value.is_array()) {
      return Shaped{broadcast_binop(op, lhs.value, rhs.value, arena), nullptr};
    }
    return Shaped{apply_binop_per_cell(op, lhs.value, rhs.value, arena), nullptr};
  }
  if (lhs.tail_array != nullptr && rhs.tail_array != nullptr && lhs.tail_array->axis != rhs.tail_array->axis) {
    return Shaped{broadcast_binop(op, densify(lhs, arena), densify(rhs, arena), arena), nullptr};
  }
  const TailAxis axis = lhs.tail_array != nullptr ? lhs.tail_array->axis : rhs.tail_array->axis;
  Value l_slot = Value::blank();
  Value r_slot = Value::blank();
  const ArrayView la = as_array_view(lhs, &l_slot);
  const ArrayView ra = as_array_view(rhs, &r_slot);
  const std::uint32_t out_rows = la.rows > ra.rows ? la.rows : ra.rows;
  const std::uint32_t out_cols = la.cols > ra.cols ? la.cols : ra.cols;
  const std::uint32_t out_ext = axis == TailAxis::kRows ? out_rows : out_cols;

  // Every position past the result head must read a constant from each
  // operand; a TailArray shorter than the output (but not stretched) ends
  // mid-way, so that shape goes through the dense path.
  std::uint32_t head = 0;
  for (const ArrayView* v : {&la, &ra}) {
    const std::uint32_t ext = axis_extent(*v, axis);
    if (v->tail != nullptr) {
      if (ext != out_ext && ext != 1U) {
        return Shaped{broadcast_binop(op, densify(lhs, arena), densify(rhs, arena), arena), nullptr};
      }
      head = v->head > head ? v->head : head;
    } else if (ext > 1U && ext > head) {
      head = ext;
    }
  }
  return build_tail_result(out_rows, out_cols, head, axis, arena, [&](std::uint32_t r, std::uint32_t c) {
    const Value* lv = broadcast_cell(la, r, c);
    const Value* rv = broadcast_cell(ra, r, c);
    if (lv == nullptr || rv == nullptr) {
      return Value::error(ErrorCode::NA);
    }
    return apply_binop_per_cell(op, *lv, *rv, arena);
  });
}

Shaped broadcast_unary(parser::UnaryOp op, const Shaped& operand, Arena& arena) {
  if (operand.tail_array == nullptr) {
    if (operand.value.is_error()) {
      return operand;
    }
    return Shaped{broadcast_unary(op, operand.value, arena), nullptr};
  }
  const TailArray& t = *operand.tail_array;
  return build_tail_result(t.rows, t.cols, t.head, t.axis, arena, [&](std::uint32_t r, std::uint32_t c) {
    const Value& cell = tail_array_at(t, r, c);
    return cell.is_error() ? cell : apply_unary(op, cell);
  });
}

}  // namespace eval
}  // namespace formulon
