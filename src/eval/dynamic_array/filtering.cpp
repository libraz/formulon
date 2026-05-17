// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "eval/dynamic_array/filtering.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <numeric>
#include <vector>

#include "eval/coerce.h"
#include "eval/dynamic_array/common.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {

Value eval_filter_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve `array` and `include` as Value::Array via the array-context seam
  // so range-shaped args (Ref / RangeOp / OFFSET / CHOOSE / IF / SpillRef)
  // keep their 2D shape. Errors short-circuit before we touch shape logic.
  const ArrayValue* array = nullptr;
  const ArrayValue* include = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err) ||
      !resolve_array_value(call.as_call_arg(1), arena, registry, ctx, &include, &err)) {
    return err;
  }

  // Axis decision. `include` must be a 1D vector matching one of `array`'s
  // dimensions. Both shapes equal -> filter rows (the column-vector form is
  // the dominant Excel idiom). Otherwise either an exact column-vector
  // match (rows-axis) or row-vector match (cols-axis); anything else is
  // ambiguous / mismatched.
  enum class Axis { Rows, Cols };
  Axis axis = Axis::Rows;
  if (include->rows == array->rows && include->cols == 1U) {
    axis = Axis::Rows;
  } else if (include->cols == array->cols && include->rows == 1U) {
    axis = Axis::Cols;
  } else {
    return Value::error(ErrorCode::Value);
  }

  // Coerce each include cell to bool, recording the kept indices. Coercion
  // errors and explicit Error cells short-circuit the whole call -- Mac
  // Excel's conservative behaviour for mixed-shape spill candidates.
  const std::uint32_t mask_n = (axis == Axis::Rows) ? include->rows : include->cols;
  std::vector<std::uint32_t> kept;
  kept.reserve(mask_n);
  for (std::uint32_t i = 0; i < mask_n; ++i) {
    const Value& cell = include->cells[i];
    if (cell.is_error()) {
      return cell;
    }
    auto coerced = coerce_to_bool(cell);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    if (coerced.value()) {
      kept.push_back(i);
    }
  }

  // Empty-result handling: return `if_empty` if provided (scalar, NOT
  // spilled), otherwise #CALC! per Mac Excel's documented surface.
  if (kept.empty()) {
    if (arity == 3U) {
      return eval_node(call.as_call_arg(2), arena, registry, ctx);
    }
    return Value::error(ErrorCode::Calc);
  }

  ArrayValue* out = dynamic_array::materialise_selected_lanes(*array, kept, axis == Axis::Cols, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_unique_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve `array` via the array-context seam so range-shaped args
  // (Ref / RangeOp / OFFSET / CHOOSE / IF / SpillRef) keep their 2D shape.
  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }
  if (array->rows == 0U || array->cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  // Optional flags. Both default to FALSE; both go through `coerce_to_bool`
  // so that 0/1, "TRUE"/"FALSE", and Blank scalars behave as Excel does
  // (Blank coerces to FALSE per `coerce_to_bool`).
  bool by_col = false;
  if (arity >= 2U) {
    const Value flag = eval_node(call.as_call_arg(1), arena, registry, ctx);
    if (flag.is_error()) {
      return flag;
    }
    auto coerced = coerce_to_bool(flag);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    by_col = coerced.value();
  }
  bool exactly_once = false;
  if (arity >= 3U) {
    const Value flag = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (flag.is_error()) {
      return flag;
    }
    auto coerced = coerce_to_bool(flag);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    exactly_once = coerced.value();
  }

  // Walk lanes (rows or columns) once. `distinct` records first-occurrence
  // lane indices in input order; `count[k]` totals how many input lanes
  // matched `distinct[k]` (used by exactly_once filtering). O(n^2) in the
  // lane count is intentional: UNIQUE inputs are typically small, and
  // avoiding a value-keyed hash table sidesteps a cross-kind hashing
  // contract we do not currently maintain.
  const std::uint32_t lanes = by_col ? array->cols : array->rows;
  std::vector<std::uint32_t> distinct;
  std::vector<std::uint32_t> count;
  distinct.reserve(lanes);
  count.reserve(lanes);
  for (std::uint32_t i = 0; i < lanes; ++i) {
    bool matched = false;
    for (std::uint32_t k = 0; k < distinct.size(); ++k) {
      if (dynamic_array::unique_lane_equal(*array, i, distinct[k], by_col)) {
        ++count[k];
        matched = true;
        break;
      }
    }
    if (!matched) {
      distinct.push_back(i);
      count.push_back(1U);
    }
  }

  // Select output lanes. Default mode keeps all distinct lanes in
  // first-occurrence order; exactly-once mode keeps only those with
  // count == 1. Empty selection -> #CALC! (Excel's documented surface).
  std::vector<std::uint32_t> kept;
  kept.reserve(distinct.size());
  for (std::uint32_t k = 0; k < distinct.size(); ++k) {
    if (!exactly_once || count[k] == 1U) {
      kept.push_back(distinct[k]);
    }
  }
  if (kept.empty()) {
    return Value::error(ErrorCode::Calc);
  }

  ArrayValue* out = dynamic_array::materialise_selected_lanes(*array, kept, by_col, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_sort_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve `array` via the array-context seam so range-shaped args keep
  // their 2D shape (matches FILTER / UNIQUE).
  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }
  if (array->rows == 0U || array->cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  // Optional args. by_col is parsed first because it determines the valid
  // range of sort_index (column-bound for row-sort, row-bound for col-sort).
  bool by_col = false;
  if (arity >= 4U) {
    const Value flag = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (flag.is_error()) {
      return flag;
    }
    auto coerced = coerce_to_bool(flag);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    by_col = coerced.value();
  }

  // sort_index defaults to 1 (1-based). Out-of-range -> #VALUE!. The
  // upper bound depends on the axis: row-sort uses a column index, so
  // limit is `cols`; col-sort uses a row index, so limit is `rows`.
  std::uint32_t sort_index = 1U;
  if (arity >= 2U) {
    const Value v = eval_node(call.as_call_arg(1), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    auto coerced = coerce_to_number(v);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    const double n = coerced.value();
    if (!(n >= 1.0)) {
      return Value::error(ErrorCode::Value);
    }
    sort_index = static_cast<std::uint32_t>(n);
  }
  const std::uint32_t key_max = by_col ? array->rows : array->cols;
  if (sort_index < 1U || sort_index > key_max) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t key_idx = sort_index - 1U;  // convert to 0-based

  // sort_order: 1 ascending (default), -1 descending. Anything else is
  // #VALUE! (matches Excel's "the sort order must be 1 or -1" surface).
  bool descending = false;
  if (arity >= 3U) {
    Value order_err = Value::error(ErrorCode::Value);
    if (!dynamic_array::resolve_sort_order_arg(call.as_call_arg(2), arena, registry, ctx, descending, order_err)) {
      return order_err;
    }
  }

  // Build a permutation of lane indices, sort it stably by the chosen
  // key column / row, then materialise the output by gathering lanes.
  const std::uint32_t lanes = by_col ? array->cols : array->rows;
  std::vector<std::uint32_t> perm(lanes);
  std::iota(perm.begin(), perm.end(), 0U);

  const ArrayValue& arr_ref = *array;
  std::stable_sort(perm.begin(), perm.end(), [&](std::uint32_t a, std::uint32_t b) {
    const Value& ka = by_col ? arr_ref.cells[static_cast<std::size_t>(key_idx) * arr_ref.cols + a]
                             : arr_ref.cells[static_cast<std::size_t>(a) * arr_ref.cols + key_idx];
    const Value& kb = by_col ? arr_ref.cells[static_cast<std::size_t>(key_idx) * arr_ref.cols + b]
                             : arr_ref.cells[static_cast<std::size_t>(b) * arr_ref.cols + key_idx];
    return dynamic_array::sort_lane_less(ka, kb, descending);
  });

  ArrayValue* out = dynamic_array::materialise_selected_lanes(arr_ref, perm, by_col, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_sortby_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Need at least `(array, by_array1)`. Cap at 13 (six keys) -- a safety
  // ceiling that comfortably exceeds any realistic SORTBY usage and keeps
  // the small-vectors local.
  if (arity < 2U || arity > 13U) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = nullptr;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_array_value(call.as_call_arg(0), arena, registry, ctx, &array, &err)) {
    return err;
  }
  if (array->rows == 0U || array->cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  // Axis inference + per-key collection. The first by_array picks the axis;
  // subsequent ones must agree on shape. We hold borrowed pointers to the
  // ArrayValue cells -- safe because each `eval_node_as_array` allocates
  // into the caller arena that outlives this call.
  bool axis_decided = false;
  bool by_col = false;
  std::uint32_t lanes = 0;
  struct KeySpec {
    const ArrayValue* arr;
    bool descending;
  };
  std::vector<KeySpec> keys;
  keys.reserve((arity - 1U + 1U) / 2U);

  for (std::uint32_t i = 1; i < arity; i += 2U) {
    const Value by_v = eval_node_as_array(call.as_call_arg(i), arena, registry, ctx);
    if (by_v.is_error()) {
      return by_v;
    }
    if (!by_v.is_array()) {
      return Value::error(ErrorCode::Value);
    }
    const ArrayValue* by = by_v.as_array();

    // Decide axis from the first key, then enforce match on the rest.
    if (!axis_decided) {
      if (by->cols == 1U && by->rows == array->rows) {
        by_col = false;
        lanes = array->rows;
      } else if (by->rows == 1U && by->cols == array->cols) {
        by_col = true;
        lanes = array->cols;
      } else {
        return Value::error(ErrorCode::Value);
      }
      axis_decided = true;
    } else {
      const bool matches = by_col ? (by->rows == 1U && by->cols == lanes) : (by->cols == 1U && by->rows == lanes);
      if (!matches) {
        return Value::error(ErrorCode::Value);
      }
    }

    // Optional per-key order (1 or -1). Default ascending when absent.
    bool descending = false;
    if (i + 1U < arity) {
      Value order_err = Value::error(ErrorCode::Value);
      if (!dynamic_array::resolve_sort_order_arg(call.as_call_arg(i + 1U), arena, registry, ctx, descending,
                                                 order_err)) {
        return order_err;
      }
    }
    keys.push_back(KeySpec{by, descending});
  }

  // Lane-key accessor. Both axes flatten to a 1D index since each by_array
  // is 1D by construction (column-vector or row-vector matching `lanes`).
  auto key_at = [](const ArrayValue* by, std::uint32_t idx) -> const Value& { return by->cells[idx]; };

  // Build permutation, sort with multi-key compare. Stable so rows that
  // tie on every key keep their input order.
  std::vector<std::uint32_t> perm(lanes);
  std::iota(perm.begin(), perm.end(), 0U);
  std::stable_sort(perm.begin(), perm.end(), [&](std::uint32_t a, std::uint32_t b) {
    for (const KeySpec& k : keys) {
      const Value& ka = key_at(k.arr, a);
      const Value& kb = key_at(k.arr, b);
      if (dynamic_array::sort_lane_less(ka, kb, k.descending)) {
        return true;
      }
      if (dynamic_array::sort_lane_less(kb, ka, k.descending)) {
        return false;
      }
      // Equal under this key -> fall through to next.
    }
    return false;
  });

  // Materialise output, preserving input shape (axis-only reorder).
  const ArrayValue& arr_ref = *array;
  ArrayValue* out = dynamic_array::materialise_selected_lanes(arr_ref, perm, by_col, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
