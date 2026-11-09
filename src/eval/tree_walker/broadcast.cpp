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
//
// It also hosts the element-wise broadcast of scalar (non-range-aware)
// functions over array arguments, called from `dispatch.cpp`.

#include "eval/tree_walker/broadcast.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/function_registry.h"
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

void widen_shape(const ArrayView& v, std::uint32_t* rows, std::uint32_t* cols) {
  *rows = *rows > v.rows ? *rows : v.rows;
  *cols = *cols > v.cols ? *cols : v.cols;
}

namespace {

// Excel 365 broadcast shape: each axis extends to the larger operand's size.
// A size-1 axis broadcasts; a size-mismatch on a non-1 axis pads the shortfall
// with #N/A (handled per-cell via `broadcast_cell`).
void broadcast_shape(const ArrayView& a, const ArrayView& b, std::uint32_t* rows, std::uint32_t* cols) {
  *rows = a.rows;
  *cols = a.cols;
  widen_shape(b, rows, cols);
}

// Result cell at (r, c): the operator applied to the cells each operand
// contributes, or #N/A where one operand cannot supply the position.
Value broadcast_binop_cell(parser::BinOp op, const ArrayView& la, const ArrayView& ra, std::uint32_t r, std::uint32_t c,
                           Arena& arena) {
  const Value* lv = broadcast_cell(la, r, c);
  const Value* rv = broadcast_cell(ra, r, c);
  if (lv == nullptr || rv == nullptr) {
    return Value::error(ErrorCode::NA);
  }
  return apply_binop_per_cell(op, *lv, *rv, arena);
}

// Builds a dense `rows` x `cols` array whose position (r, c) is `f(r, c)`;
// `#NUM!` when the arena cannot hold it.
template <class F>
Value build_dense_result(std::uint32_t rows, std::uint32_t cols, Arena& arena, F&& f) {
  Value* buf = nullptr;
  ArrayValue* out = allocate_derived_array_value(rows, cols, arena, buf);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::size_t i = 0;
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c, ++i) {
      buf[i] = f(r, c);
    }
  }
  return Value::array(out);
}

}  // namespace

Value broadcast_binop(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena) {
  Value l_slot = Value::blank();
  Value r_slot = Value::blank();
  const ArrayView la = as_array_view(lhs, &l_slot);
  const ArrayView ra = as_array_view(rhs, &r_slot);
  std::uint32_t out_rows = 0;
  std::uint32_t out_cols = 0;
  broadcast_shape(la, ra, &out_rows, &out_cols);
  return build_dense_result(out_rows, out_cols, arena, [&](std::uint32_t r, std::uint32_t c) {
    return broadcast_binop_cell(op, la, ra, r, c, arena);
  });
}

Value broadcast_unary(parser::UnaryOp op, const Value& operand, Arena& arena) {
  if (!operand.is_array()) {
    return apply_unary(op, operand);
  }
  const ArrayValue* in = operand.as_array();
  return build_dense_result(in->rows, in->cols, arena, [&](std::uint32_t r, std::uint32_t c) {
    const Value& cell = in->cells[static_cast<std::size_t>(r) * in->cols + c];
    return cell.is_error() ? cell : apply_unary(op, cell);
  });
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
    return Shaped{build_dense_result(rows, cols, arena, f), nullptr};
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
  // Between two scalars the left-most error wins; against an array operand
  // a scalar error is broadcast like any other scalar, error per position.
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
  std::uint32_t out_rows = 0;
  std::uint32_t out_cols = 0;
  broadcast_shape(la, ra, &out_rows, &out_cols);
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
    return broadcast_binop_cell(op, la, ra, r, c, arena);
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

namespace {

// One argument slot for element-wise broadcasting of a scalar function
// over array-shaped arguments: a scalar (a 1x1 view, `scalar` set), an
// ArrayValue, or a whole column / row (`view.tail` set).
struct BroadcastArg {
  ArrayView view;
  bool scalar;
};

BroadcastArg scalar_broadcast_arg(const Value& v) {
  return {ArrayView{1U, 1U, &v}, true};
}

BroadcastArg array_broadcast_arg(const ArrayValue& a) {
  return {ArrayView{a.rows, a.cols, a.cells}, false};
}

BroadcastArg tail_broadcast_arg(const TailArray& t) {
  return {ArrayView{t.rows, t.cols, t.cells, t.head, t.tail, t.axis}, false};
}

// Element of `arg` at output position (r, c) under Excel's 1xN / Nx1
// broadcast rules. A scalar supplies every position; an array whose
// non-1 extent is smaller than the output cannot supply the position and
// yields `#N/A`, matching Excel's ragged-broadcast fill.
Value broadcast_element(const BroadcastArg& arg, std::uint32_t r, std::uint32_t c) {
  const Value* cell = broadcast_cell(arg.view, r, c);
  return cell != nullptr ? *cell : Value::error(ErrorCode::NA);
}

// `def` applied to the arguments at output position (r, c); with `propagate_errors`
// the first error argument short-circuits the cell before the impl runs.
Value eval_broadcast_cell(const FunctionDef& def, const std::vector<BroadcastArg>& views, std::vector<Value>& cell_args,
                          std::uint32_t r, std::uint32_t c, Arena& arena) {
  for (std::size_t a = 0; a < views.size(); ++a) {
    cell_args[a] = broadcast_element(views[a], r, c);
    if (def.propagate_errors && cell_args[a].is_error()) {
      return cell_args[a];
    }
  }
  return def.impl(cell_args.data(), static_cast<std::uint32_t>(cell_args.size()), arena);
}

}  // namespace

// Evaluates a scalar (non-range-aware) function element-wise across the
// broadcast rectangle of its already-collected arguments and returns a
// spilled `Value::Array`. `out_rows` / `out_cols` are the max extents
// over the array arguments; scalars broadcast to every cell. A 1x1
// result unwraps to a plain scalar so a single-cell range argument does
// not spill. Per-cell error propagation mirrors the scalar dispatch
// path: with `propagate_errors`, the first error argument short-circuits
// that cell before the impl runs.
Value broadcast_scalar_call(const FunctionDef& def, const std::vector<Value>& args, std::uint32_t out_rows,
                            std::uint32_t out_cols, Arena& arena) {
  std::vector<BroadcastArg> views;
  views.reserve(args.size());
  for (const Value& v : args) {
    views.push_back(v.is_array() ? array_broadcast_arg(*v.as_array()) : scalar_broadcast_arg(v));
  }
  std::vector<Value> cell_args(views.size(), Value::blank());
  const Value result = build_dense_result(out_rows, out_cols, arena, [&](std::uint32_t r, std::uint32_t c) {
    return eval_broadcast_cell(def, views, cell_args, r, c, arena);
  });
  if (out_rows == 1U && out_cols == 1U && result.is_array()) {
    return result.as_array()->cells[0];
  }
  return result;
}

// Evaluates a `last_arg_scalar` function once per element of `last`, the
// array its last argument evaluated to, with `values` holding the other
// arguments. A 1x1 result unwraps to a scalar.
Value lift_last_arg_call(const FunctionDef& def, std::vector<Value> values, const ArrayValue& last, Arena& arena) {
  return map_array_with(
      last,
      [&](const Value& cell) {
        values.back() = cell;
        return def.propagate_errors && cell.is_error()
                   ? cell
                   : def.impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
      },
      arena);
}

// `broadcast_scalar_call` for arguments of which `tails[i]` (when non-null) is
// a whole column / row standing in for `args[i]`. Arguments on one axis give a
// `TailArray`: the function is applied to the head cells and once to the tail
// row / column. Whole columns mixed with whole rows are expanded and take the
// dense route.
Shaped broadcast_scalar_call_shaped(const FunctionDef& def, const std::vector<Value>& args,
                                    const std::vector<const TailArray*>& tails, Arena& arena) {
  bool by_rows = false;
  bool by_cols = false;
  for (const TailArray* ta : tails) {
    if (ta != nullptr) {
      (ta->axis == TailAxis::kRows ? by_rows : by_cols) = true;
    }
  }
  std::vector<BroadcastArg> views;
  views.reserve(args.size());
  std::uint32_t out_rows = 1;
  std::uint32_t out_cols = 1;
  std::vector<Value> dense_args;
  const bool mixed = by_rows && by_cols;
  if (mixed) {
    dense_args = args;
    for (std::size_t i = 0; i < args.size(); ++i) {
      if (tails[i] != nullptr) {
        Shaped s;
        s.tail_array = tails[i];
        dense_args[i] = densify(s, arena);
        if (dense_args[i].is_error()) {
          return Shaped{dense_args[i], nullptr};
        }
      }
    }
  }
  const std::vector<Value>& source = mixed ? dense_args : args;
  for (std::size_t i = 0; i < source.size(); ++i) {
    if (!mixed && tails[i] != nullptr) {
      views.push_back(tail_broadcast_arg(*tails[i]));
    } else if (source[i].is_array()) {
      views.push_back(array_broadcast_arg(*source[i].as_array()));
    } else {
      views.push_back(scalar_broadcast_arg(source[i]));
    }
    if (!views.back().scalar) {
      out_rows = std::max(out_rows, views.back().view.rows);
      out_cols = std::max(out_cols, views.back().view.cols);
    }
  }
  std::vector<Value> cell_args(views.size(), Value::blank());
  const auto eval_cell = [&](std::vector<Value>& args_buf, std::uint32_t r, std::uint32_t c) {
    return eval_broadcast_cell(def, views, args_buf, r, c, arena);
  };
  if (mixed) {
    return Shaped{build_dense_result(out_rows, out_cols, arena,
                                     [&](std::uint32_t r, std::uint32_t c) { return eval_cell(cell_args, r, c); }),
                  nullptr};
  }

  // Positions past `head` read a constant from every argument: a tail, a
  // stretched size-1 axis, or the `#N/A` beyond a shorter dense array.
  const TailAxis axis = by_rows ? TailAxis::kRows : TailAxis::kCols;
  const bool along_rows = axis == TailAxis::kRows;
  const std::uint32_t extent = along_rows ? out_rows : out_cols;
  std::uint32_t head = 0;
  for (const BroadcastArg& v : views) {
    if (v.view.tail != nullptr) {
      head = std::max(head, v.view.head);
    } else if (!v.scalar) {
      const std::uint32_t ext = axis_extent(v.view, axis);
      if (ext > 1U) {
        head = std::max(head, ext);
      }
    }
  }
  head = std::min(head, extent);
  const std::size_t line = along_rows ? out_cols : out_rows;
  const std::uint64_t head_cells = static_cast<std::uint64_t>(head) * line;
  if (head_cells > kMaxDerivedArrayCells) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  Value* cells = head_cells == 0U ? nullptr : arena.create_array<Value>(static_cast<std::size_t>(head_cells));
  Value* tail = arena.create_array<Value>(line);
  if ((head_cells != 0U && cells == nullptr) || tail == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  std::size_t idx = 0;
  for (std::uint32_t r = 0; r < (along_rows ? head : out_rows); ++r) {
    for (std::uint32_t c = 0; c < (along_rows ? out_cols : head); ++c, ++idx) {
      cells[idx] = eval_cell(cell_args, r, c);
    }
  }
  const std::uint32_t tail_at = std::min(head, extent - 1U);
  for (std::uint32_t k = 0; k < line; ++k) {
    tail[k] = along_rows ? eval_cell(cell_args, tail_at, k) : eval_cell(cell_args, k, tail_at);
  }
  Shaped out;
  out.tail_array = make_tail_array(arena, out_rows, out_cols, head, axis, cells, tail, /*from_reference=*/false);
  if (out.tail_array == nullptr) {
    return Shaped{Value::error(ErrorCode::Num), nullptr};
  }
  return out;
}

}  // namespace eval
}  // namespace formulon
