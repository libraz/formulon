// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "eval/dynamic_array_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <numeric>
#include <vector>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/reference_lazy.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

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
  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const Value include_v = eval_node_as_array(call.as_call_arg(1), arena, registry, ctx);
  if (include_v.is_error()) {
    return include_v;
  }
  if (!include_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }

  const ArrayValue* array = array_v.as_array();
  const ArrayValue* include = include_v.as_array();

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

  // Build the output array. For axis=Rows the output is (kept.size(),
  // array->cols); for axis=Cols it is (array->rows, kept.size()). Cells are
  // copied row-major.
  std::uint32_t out_rows = 0;
  std::uint32_t out_cols = 0;
  if (axis == Axis::Rows) {
    out_rows = static_cast<std::uint32_t>(kept.size());
    out_cols = array->cols;
  } else {
    out_rows = array->rows;
    out_cols = static_cast<std::uint32_t>(kept.size());
  }
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  if (axis == Axis::Rows) {
    for (std::size_t i = 0; i < kept.size(); ++i) {
      const std::uint32_t src_row = kept[i];
      for (std::uint32_t c = 0; c < array->cols; ++c) {
        buffer[i * static_cast<std::size_t>(out_cols) + c] =
            array->cells[static_cast<std::size_t>(src_row) * array->cols + c];
      }
    }
  } else {
    for (std::uint32_t r = 0; r < array->rows; ++r) {
      for (std::size_t i = 0; i < kept.size(); ++i) {
        const std::uint32_t src_col = kept[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            array->cells[static_cast<std::size_t>(r) * array->cols + src_col];
      }
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

namespace {

// Excel-canonical cell equality for UNIQUE. Mirrors `Value::operator==` for
// Number / Bool / Error / Blank, but uses ASCII case-insensitive compare for
// Text (matching `=A1=B1`, COUNTIF, and the SWITCH precedent in
// `special_forms_lazy`). Cross-kind pairs are never equal.
bool unique_cell_equal(const Value& a, const Value& b) {
  if (a.kind() != b.kind()) {
    return false;
  }
  switch (a.kind()) {
    case ValueKind::Blank:
      return true;
    case ValueKind::Number:
      return a.as_number() == b.as_number();
    case ValueKind::Bool:
      return a.as_boolean() == b.as_boolean();
    case ValueKind::Error:
      return a.as_error() == b.as_error();
    case ValueKind::Text:
      return strings::case_insensitive_eq(a.as_text(), b.as_text());
    default:
      // Array / Ref / Lambda are not produced by cell reads; treat any
      // hypothetical match conservatively as "not equal" to avoid silent
      // dedup of structurally complex values we have no canonical form for.
      return false;
  }
}

// Returns true iff the i-th and j-th rows (axis=Rows) or columns
// (axis=Cols) of `arr` are cellwise equal under `unique_cell_equal`.
bool unique_lane_equal(const ArrayValue& arr, std::uint32_t i, std::uint32_t j, bool by_col) {
  if (by_col) {
    for (std::uint32_t r = 0; r < arr.rows; ++r) {
      const Value& av = arr.cells[static_cast<std::size_t>(r) * arr.cols + i];
      const Value& bv = arr.cells[static_cast<std::size_t>(r) * arr.cols + j];
      if (!unique_cell_equal(av, bv)) {
        return false;
      }
    }
    return true;
  }
  for (std::uint32_t c = 0; c < arr.cols; ++c) {
    const Value& av = arr.cells[static_cast<std::size_t>(i) * arr.cols + c];
    const Value& bv = arr.cells[static_cast<std::size_t>(j) * arr.cols + c];
    if (!unique_cell_equal(av, bv)) {
      return false;
    }
  }
  return true;
}

}  // namespace

Value eval_unique_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve `array` via the array-context seam so range-shaped args
  // (Ref / RangeOp / OFFSET / CHOOSE / IF / SpillRef) keep their 2D shape.
  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();
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
      if (unique_lane_equal(*array, i, distinct[k], by_col)) {
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

  // Build the output array. Layout mirrors FILTER: by_col=false produces
  // (kept.size(), array->cols); by_col=true produces (array->rows,
  // kept.size()).
  std::uint32_t out_rows = 0;
  std::uint32_t out_cols = 0;
  if (by_col) {
    out_rows = array->rows;
    out_cols = static_cast<std::uint32_t>(kept.size());
  } else {
    out_rows = static_cast<std::uint32_t>(kept.size());
    out_cols = array->cols;
  }
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  if (by_col) {
    for (std::uint32_t r = 0; r < array->rows; ++r) {
      for (std::size_t i = 0; i < kept.size(); ++i) {
        const std::uint32_t src_col = kept[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            array->cells[static_cast<std::size_t>(r) * array->cols + src_col];
      }
    }
  } else {
    for (std::size_t i = 0; i < kept.size(); ++i) {
      const std::uint32_t src_row = kept[i];
      for (std::uint32_t c = 0; c < array->cols; ++c) {
        buffer[i * static_cast<std::size_t>(out_cols) + c] =
            array->cells[static_cast<std::size_t>(src_row) * array->cols + c];
      }
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

namespace {

// Excel-canonical kind rank for SORT. Lower rank sorts earlier under
// ascending order. Blank deliberately sits at rank 5 -- the SORT impl
// special-cases blanks so they always sink to the end regardless of
// `sort_order`. Numbers / Text / Bool / Error participate normally.
int sort_kind_rank(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number:
      return 0;
    case ValueKind::Text:
      return 1;
    case ValueKind::Bool:
      return 2;
    case ValueKind::Error:
      return 3;
    case ValueKind::Blank:
      return 5;
    default:
      // Array / Ref / Lambda are not produced by cell reads. Slot them
      // between Error and Blank so they participate predictably without
      // colliding with real-world cells.
      return 4;
  }
}

// ASCII case-insensitive lex comparison. Returns true iff `a` < `b`.
// Walks both strings byte-by-byte with `tolower` applied, matching the
// Excel-canonical text ordering used by SWITCH / UNIQUE for equality.
// Avoids materialising a full lowercase copy.
bool ascii_ci_less(std::string_view a, std::string_view b) {
  const std::size_t n = std::min(a.size(), b.size());
  for (std::size_t i = 0; i < n; ++i) {
    const unsigned char ca = static_cast<unsigned char>(a[i]);
    const unsigned char cb = static_cast<unsigned char>(b[i]);
    const unsigned char la = (ca >= 'A' && ca <= 'Z') ? static_cast<unsigned char>(ca + 32) : ca;
    const unsigned char lb = (cb >= 'A' && cb <= 'Z') ? static_cast<unsigned char>(cb + 32) : cb;
    if (la != lb) {
      return la < lb;
    }
  }
  return a.size() < b.size();
}

// Strict weak ordering for SORT keys, ascending. Cross-kind ordering
// follows `sort_kind_rank`; within-kind ordering matches Excel:
//   Number  -> numeric compare via `<`.
//   Text    -> ASCII case-insensitive lex.
//   Bool    -> FALSE < TRUE.
//   Error   -> error code value compare (gives a stable, deterministic
//             order; matches "errors group together" surface).
//   Blank   -> equal (handled by the rank, but a same-rank pair is also
//             a no-op here).
bool sort_cell_less_asc(const Value& a, const Value& b) {
  const int ra = sort_kind_rank(a);
  const int rb = sort_kind_rank(b);
  if (ra != rb) {
    return ra < rb;
  }
  switch (a.kind()) {
    case ValueKind::Number:
      return a.as_number() < b.as_number();
    case ValueKind::Text:
      return ascii_ci_less(a.as_text(), b.as_text());
    case ValueKind::Bool:
      return !a.as_boolean() && b.as_boolean();
    case ValueKind::Error:
      return static_cast<int>(a.as_error()) < static_cast<int>(b.as_error());
    case ValueKind::Blank:
    default:
      return false;
  }
}

// Lane-level less for SORT. Always sinks blank keys to the end regardless
// of `sort_order`; otherwise applies `sort_cell_less_asc` (or its mirror
// for descending). The stable_sort caller relies on this being a strict
// weak ordering.
bool sort_lane_less(const Value& key_a, const Value& key_b, bool descending) {
  const bool a_blank = key_a.is_blank();
  const bool b_blank = key_b.is_blank();
  if (a_blank != b_blank) {
    // Non-blank wins (sorts earlier) regardless of direction.
    return !a_blank;
  }
  if (a_blank && b_blank) {
    return false;
  }
  return descending ? sort_cell_less_asc(key_b, key_a) : sort_cell_less_asc(key_a, key_b);
}

}  // namespace

Value eval_sort_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve `array` via the array-context seam so range-shaped args keep
  // their 2D shape (matches FILTER / UNIQUE).
  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();
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
    const Value v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (v.is_error()) {
      return v;
    }
    auto coerced = coerce_to_number(v);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    const double n = coerced.value();
    if (n == 1.0) {
      descending = false;
    } else if (n == -1.0) {
      descending = true;
    } else {
      return Value::error(ErrorCode::Value);
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
    return sort_lane_less(ka, kb, descending);
  });

  // SORT preserves the input shape; just reorder the lanes in place.
  const std::uint32_t out_rows = array->rows;
  const std::uint32_t out_cols = array->cols;
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  if (by_col) {
    for (std::uint32_t r = 0; r < out_rows; ++r) {
      for (std::uint32_t i = 0; i < lanes; ++i) {
        const std::uint32_t src_col = perm[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            arr_ref.cells[static_cast<std::size_t>(r) * arr_ref.cols + src_col];
      }
    }
  } else {
    for (std::uint32_t i = 0; i < lanes; ++i) {
      const std::uint32_t src_row = perm[i];
      for (std::uint32_t c = 0; c < out_cols; ++c) {
        buffer[static_cast<std::size_t>(i) * out_cols + c] =
            arr_ref.cells[static_cast<std::size_t>(src_row) * arr_ref.cols + c];
      }
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
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

  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();
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
      const Value v = eval_node(call.as_call_arg(i + 1U), arena, registry, ctx);
      if (v.is_error()) {
        return v;
      }
      auto coerced = coerce_to_number(v);
      if (!coerced) {
        return Value::error(coerced.error());
      }
      const double n = coerced.value();
      if (n == 1.0) {
        descending = false;
      } else if (n == -1.0) {
        descending = true;
      } else {
        return Value::error(ErrorCode::Value);
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
      if (sort_lane_less(ka, kb, k.descending)) {
        return true;
      }
      if (sort_lane_less(kb, ka, k.descending)) {
        return false;
      }
      // Equal under this key -> fall through to next.
    }
    return false;
  });

  // Materialise output, preserving input shape (axis-only reorder).
  const ArrayValue& arr_ref = *array;
  const std::uint32_t out_rows = array->rows;
  const std::uint32_t out_cols = array->cols;
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  if (by_col) {
    for (std::uint32_t r = 0; r < out_rows; ++r) {
      for (std::uint32_t i = 0; i < lanes; ++i) {
        const std::uint32_t src_col = perm[i];
        buffer[static_cast<std::size_t>(r) * out_cols + i] =
            arr_ref.cells[static_cast<std::size_t>(r) * arr_ref.cols + src_col];
      }
    }
  } else {
    for (std::uint32_t i = 0; i < lanes; ++i) {
      const std::uint32_t src_row = perm[i];
      for (std::uint32_t c = 0; c < out_cols; ++c) {
        buffer[static_cast<std::size_t>(i) * out_cols + c] =
            arr_ref.cells[static_cast<std::size_t>(src_row) * arr_ref.cols + c];
      }
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

namespace {

/// Resolve every call argument as a Value::Array, propagating the first
/// error verbatim. Returns true on success and populates `out`. The
/// returned ArrayValue pointers borrow from the caller arena, which must
/// outlive their use.
bool resolve_stack_args(const parser::AstNode& call, std::uint32_t arity, Arena& arena,
                        const FunctionRegistry& registry, const EvalContext& ctx, std::vector<const ArrayValue*>& out,
                        Value& error_out) {
  out.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    const Value v = eval_node_as_array(call.as_call_arg(i), arena, registry, ctx);
    if (v.is_error()) {
      error_out = v;
      return false;
    }
    if (!v.is_array()) {
      error_out = Value::error(ErrorCode::Value);
      return false;
    }
    out.push_back(v.as_array());
  }
  return true;
}

}  // namespace

Value eval_hstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Excel's documented HSTACK accepts up to 254 array references; reject
  // 0-arg explicitly (no array context to spill into) and use 254 as the
  // safety ceiling.
  if (arity < 1U || arity > 254U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<const ArrayValue*> arrays;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_stack_args(call, arity, arena, registry, ctx, arrays, err)) {
    return err;
  }

  // Output shape: row count is the max input row count; column count is
  // the sum of input column counts. Empty inputs are impossible because
  // `eval_node_as_array` always wraps to at least 1x1.
  std::uint32_t out_rows = 0;
  std::size_t out_cols_sz = 0;
  for (const ArrayValue* a : arrays) {
    if (a->rows > out_rows) {
      out_rows = a->rows;
    }
    out_cols_sz += static_cast<std::size_t>(a->cols);
  }
  if (out_cols_sz > static_cast<std::size_t>(0xFFFFFFFFU)) {
    return Value::error(ErrorCode::Num);
  }
  const auto out_cols = static_cast<std::uint32_t>(out_cols_sz);
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  // Fill the buffer column-band by column-band. Cells above each input's
  // row count get #N/A per Mac Excel's stack-family contract.
  std::uint32_t col_off = 0;
  for (const ArrayValue* a : arrays) {
    for (std::uint32_t r = 0; r < out_rows; ++r) {
      for (std::uint32_t c = 0; c < a->cols; ++c) {
        const std::size_t dst = static_cast<std::size_t>(r) * out_cols + col_off + c;
        if (r < a->rows) {
          buffer[dst] = a->cells[static_cast<std::size_t>(r) * a->cols + c];
        } else {
          buffer[dst] = Value::error(ErrorCode::NA);
        }
      }
    }
    col_off += a->cols;
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

Value eval_vstack_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 254U) {
    return Value::error(ErrorCode::Value);
  }

  std::vector<const ArrayValue*> arrays;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_stack_args(call, arity, arena, registry, ctx, arrays, err)) {
    return err;
  }

  // Output shape: row count is the sum of input row counts; column count
  // is the max input column count.
  std::size_t out_rows_sz = 0;
  std::uint32_t out_cols = 0;
  for (const ArrayValue* a : arrays) {
    out_rows_sz += static_cast<std::size_t>(a->rows);
    if (a->cols > out_cols) {
      out_cols = a->cols;
    }
  }
  if (out_rows_sz > static_cast<std::size_t>(0xFFFFFFFFU)) {
    return Value::error(ErrorCode::Num);
  }
  const auto out_rows = static_cast<std::uint32_t>(out_rows_sz);
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  // Fill the buffer row-band by row-band. Cells past each input's column
  // count get #N/A.
  std::uint32_t row_off = 0;
  for (const ArrayValue* a : arrays) {
    for (std::uint32_t r = 0; r < a->rows; ++r) {
      for (std::uint32_t c = 0; c < out_cols; ++c) {
        const std::size_t dst = static_cast<std::size_t>(row_off + r) * out_cols + c;
        if (c < a->cols) {
          buffer[dst] = a->cells[static_cast<std::size_t>(r) * a->cols + c];
        } else {
          buffer[dst] = Value::error(ErrorCode::NA);
        }
      }
    }
    row_off += a->rows;
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

namespace {

/// Resolve a 1-based / negative index argument to a 0-based axis index.
/// Returns `true` and writes `out` on success; `false` and writes the
/// caller-visible error value on coercion failure or out-of-range.
bool resolve_choose_index(const parser::AstNode& node, std::uint32_t axis_size, Arena& arena,
                          const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& out,
                          Value& error_out) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  // Truncate-toward-zero on the user-supplied index. `0` after truncation
  // is invalid; positives map to `[1, axis_size]`, negatives to
  // `[-axis_size, -1]`. Anything else surfaces #VALUE!.
  const double truncated = std::trunc(coerced.value());
  if (truncated == 0.0) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  if (truncated > 0.0) {
    if (truncated > static_cast<double>(axis_size)) {
      error_out = Value::error(ErrorCode::Value);
      return false;
    }
    out = static_cast<std::uint32_t>(truncated) - 1U;
    return true;
  }
  // Negative index path. `-1` maps to the last element (`axis_size - 1`).
  const double abs_idx = -truncated;
  if (abs_idx > static_cast<double>(axis_size)) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  out = axis_size - static_cast<std::uint32_t>(abs_idx);
  return true;
}

}  // namespace

Value eval_choosecols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Need `(array, col_num1, ...)`; cap at 254 indices (Excel ceiling) +
  // 1 array slot = 255 args.
  if (arity < 2U || arity > 255U) {
    return Value::error(ErrorCode::Value);
  }

  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  // Resolve each index in turn. `resolve_choose_index` handles coercion,
  // truncation, sign mapping, and bounds in one place.
  std::vector<std::uint32_t> picks;
  picks.reserve(arity - 1U);
  for (std::uint32_t i = 1; i < arity; ++i) {
    std::uint32_t idx = 0;
    Value err = Value::error(ErrorCode::Value);
    if (!resolve_choose_index(call.as_call_arg(i), array->cols, arena, registry, ctx, idx, err)) {
      return err;
    }
    picks.push_back(idx);
  }

  const std::uint32_t out_rows = array->rows;
  const auto out_cols = static_cast<std::uint32_t>(picks.size());
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t k = 0; k < out_cols; ++k) {
      const std::uint32_t src_col = picks[k];
      buffer[static_cast<std::size_t>(r) * out_cols + k] =
          array->cells[static_cast<std::size_t>(r) * array->cols + src_col];
    }
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

Value eval_chooserows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 255U) {
    return Value::error(ErrorCode::Value);
  }

  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  std::vector<std::uint32_t> picks;
  picks.reserve(arity - 1U);
  for (std::uint32_t i = 1; i < arity; ++i) {
    std::uint32_t idx = 0;
    Value err = Value::error(ErrorCode::Value);
    if (!resolve_choose_index(call.as_call_arg(i), array->rows, arena, registry, ctx, idx, err)) {
      return err;
    }
    picks.push_back(idx);
  }

  const auto out_rows = static_cast<std::uint32_t>(picks.size());
  const std::uint32_t out_cols = array->cols;
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t k = 0; k < out_rows; ++k) {
    const std::uint32_t src_row = picks[k];
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(k) * out_cols + c] =
          array->cells[static_cast<std::size_t>(src_row) * array->cols + c];
    }
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

namespace {

/// Decode a signed (rows or columns) argument for TAKE / DROP into a
/// half-open `[lo, hi)` slice of the source axis. `take == true` selects
/// the cells that will be retained; `take == false` selects the cells
/// that will be dropped (and so the retained range is the complement,
/// computed in-place by this helper). On success populates `lo` / `hi`
/// and returns `true`. On argument error / coercion failure / zero (TAKE
/// only) / over-removal (DROP only), returns `false` and writes the
/// caller-visible error.
///
/// Mac Excel semantics summary:
///   * Positive count -> operate on the leading edge.
///   * Negative count -> operate on the trailing edge.
///   * `|count|` clamps to `axis_size` (TAKE returns the whole axis;
///     DROP returns nothing -> #CALC!).
///   * Omitted argument: TAKE keeps all cells on that axis; DROP drops
///     none. Caller signals "omitted" by passing `nullptr` for `node`.
bool resolve_take_drop_range(const parser::AstNode* node, std::uint32_t axis_size, bool take, Arena& arena,
                             const FunctionRegistry& registry, const EvalContext& ctx, std::uint32_t& lo,
                             std::uint32_t& hi, Value& error_out) {
  if (node == nullptr) {
    lo = 0;
    hi = axis_size;
    return true;
  }
  const Value v = eval_node(*node, arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    error_out = Value::error(coerced.error());
    return false;
  }
  const double count = std::trunc(coerced.value());
  if (take) {
    if (count == 0.0) {
      error_out = Value::error(ErrorCode::Calc);
      return false;
    }
    if (count > 0.0) {
      const auto take_n = (count >= static_cast<double>(axis_size)) ? axis_size : static_cast<std::uint32_t>(count);
      lo = 0;
      hi = take_n;
      return true;
    }
    const double abs_count = -count;
    const auto take_n =
        (abs_count >= static_cast<double>(axis_size)) ? axis_size : static_cast<std::uint32_t>(abs_count);
    lo = axis_size - take_n;
    hi = axis_size;
    return true;
  }
  if (count >= 0.0) {
    if (count >= static_cast<double>(axis_size)) {
      error_out = Value::error(ErrorCode::Calc);
      return false;
    }
    lo = static_cast<std::uint32_t>(count);
    hi = axis_size;
    return true;
  }
  const double abs_count = -count;
  if (abs_count >= static_cast<double>(axis_size)) {
    error_out = Value::error(ErrorCode::Calc);
    return false;
  }
  lo = 0;
  hi = axis_size - static_cast<std::uint32_t>(abs_count);
  return true;
}

/// Materialise a 2D row-major slice `[row_lo, row_hi) x [col_lo, col_hi)`
/// of `src` into the caller arena. Returns nullptr on arena OOM. Both
/// half-open intervals must satisfy `lo <= hi <= src axis size`.
ArrayValue* materialise_slice(const ArrayValue& src, std::uint32_t row_lo, std::uint32_t row_hi, std::uint32_t col_lo,
                              std::uint32_t col_hi, Arena& arena) {
  const std::uint32_t out_rows = row_hi - row_lo;
  const std::uint32_t out_cols = col_hi - col_lo;
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return nullptr;
  }
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      buffer[static_cast<std::size_t>(r) * out_cols + c] =
          src.cells[static_cast<std::size_t>(row_lo + r) * src.cols + (col_lo + c)];
    }
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return nullptr;
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return out;
}

}  // namespace

Value eval_take_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  std::uint32_t row_lo = 0;
  std::uint32_t row_hi = 0;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_take_drop_range(&call.as_call_arg(1), array->rows, /*take=*/true, arena, registry, ctx, row_lo, row_hi,
                               err)) {
    return err;
  }
  std::uint32_t col_lo = 0;
  std::uint32_t col_hi = array->cols;
  if (arity == 3U) {
    if (!resolve_take_drop_range(&call.as_call_arg(2), array->cols, /*take=*/true, arena, registry, ctx, col_lo, col_hi,
                                 err)) {
      return err;
    }
  }

  ArrayValue* out = materialise_slice(*array, row_lo, row_hi, col_lo, col_hi, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_drop_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }

  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  std::uint32_t row_lo = 0;
  std::uint32_t row_hi = 0;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_take_drop_range(&call.as_call_arg(1), array->rows, /*take=*/false, arena, registry, ctx, row_lo, row_hi,
                               err)) {
    return err;
  }
  std::uint32_t col_lo = 0;
  std::uint32_t col_hi = array->cols;
  if (arity == 3U) {
    if (!resolve_take_drop_range(&call.as_call_arg(2), array->cols, /*take=*/false, arena, registry, ctx, col_lo,
                                 col_hi, err)) {
      return err;
    }
  }

  ArrayValue* out = materialise_slice(*array, row_lo, row_hi, col_lo, col_hi, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_expand_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2U || arity > 4U) {
    return Value::error(ErrorCode::Value);
  }

  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  // Resolve the new row count. `rows` is required.
  const Value rows_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (rows_v.is_error()) {
    return rows_v;
  }
  auto rows_c = coerce_to_number(rows_v);
  if (!rows_c) {
    return Value::error(rows_c.error());
  }
  const double rows_d = std::trunc(rows_c.value());
  if (rows_d < static_cast<double>(array->rows)) {
    return Value::error(ErrorCode::Value);
  }

  // Resolve the new column count. Optional; defaults to the existing
  // column count (no horizontal expansion).
  std::uint32_t out_cols = array->cols;
  if (arity >= 3U) {
    const Value cols_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (cols_v.is_error()) {
      return cols_v;
    }
    auto cols_c = coerce_to_number(cols_v);
    if (!cols_c) {
      return Value::error(cols_c.error());
    }
    const double cols_d = std::trunc(cols_c.value());
    if (cols_d < static_cast<double>(array->cols)) {
      return Value::error(ErrorCode::Value);
    }
    out_cols = static_cast<std::uint32_t>(cols_d);
  }
  const auto out_rows = static_cast<std::uint32_t>(rows_d);

  // Resolve the pad value. Default is #N/A. If `pad_with` is supplied as
  // an array (range / range literal), Excel takes the scalar at cell 0
  // and pads with that — matching the dispatcher's scalar coercion path
  // for non-array contexts. Errors at the argument level propagate.
  Value pad = Value::error(ErrorCode::NA);
  if (arity == 4U) {
    const Value pad_v = eval_node(call.as_call_arg(3), arena, registry, ctx);
    if (pad_v.is_error()) {
      return pad_v;
    }
    if (pad_v.is_array()) {
      const ArrayValue* pa = pad_v.as_array();
      if (pa->rows == 0U || pa->cols == 0U) {
        return Value::error(ErrorCode::Value);
      }
      pad = pa->cells[0];
    } else {
      pad = pad_v;
    }
  }

  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c) {
      const std::size_t dst = static_cast<std::size_t>(r) * out_cols + c;
      if (r < array->rows && c < array->cols) {
        buffer[dst] = array->cells[static_cast<std::size_t>(r) * array->cols + c];
      } else {
        buffer[dst] = pad;
      }
    }
  }

  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

namespace {

/// Walk every cell of `array` and append it to `out` (subject to the
/// `ignore` bitmask) using the iteration order set by `scan_by_column`.
/// Returns `false` on a coercion / argument-validation failure with the
/// caller-visible error written into `error_out`.
bool collect_tocol_torow_cells(const ArrayValue& array, std::int64_t ignore_mask, bool scan_by_column,
                               std::vector<Value>& out, Value& error_out) {
  // Mac Excel restricts the bitmask to 0..3 inclusive. Any other integer
  // surfaces #VALUE!.
  if (ignore_mask < 0 || ignore_mask > 3) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const bool skip_blanks = (ignore_mask & 1) != 0;
  const bool skip_errors = (ignore_mask & 2) != 0;

  auto consider = [&](const Value& v) {
    if (skip_blanks && v.is_blank()) {
      return;
    }
    if (skip_errors && v.is_error()) {
      return;
    }
    out.push_back(v);
  };

  if (scan_by_column) {
    for (std::uint32_t c = 0; c < array.cols; ++c) {
      for (std::uint32_t r = 0; r < array.rows; ++r) {
        consider(array.cells[static_cast<std::size_t>(r) * array.cols + c]);
      }
    }
  } else {
    for (std::uint32_t r = 0; r < array.rows; ++r) {
      for (std::uint32_t c = 0; c < array.cols; ++c) {
        consider(array.cells[static_cast<std::size_t>(r) * array.cols + c]);
      }
    }
  }
  return true;
}

/// Resolve the (`ignore`, `scan_by_column`) optional args shared by
/// TOCOL / TOROW. `arity` is the call arity (1..3); the call's first
/// argument is the source array, hence the optional args are at indices
/// 1 and 2 here.
bool resolve_tocol_torow_options(const parser::AstNode& call, std::uint32_t arity, Arena& arena,
                                 const FunctionRegistry& registry, const EvalContext& ctx, std::int64_t& ignore_mask,
                                 bool& scan_by_column, Value& error_out) {
  ignore_mask = 0;
  scan_by_column = false;
  if (arity >= 2U) {
    const Value ig_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
    if (ig_v.is_error()) {
      error_out = ig_v;
      return false;
    }
    auto coerced = coerce_to_number(ig_v);
    if (!coerced) {
      error_out = Value::error(coerced.error());
      return false;
    }
    ignore_mask = static_cast<std::int64_t>(std::trunc(coerced.value()));
  }
  if (arity >= 3U) {
    const Value sc_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (sc_v.is_error()) {
      error_out = sc_v;
      return false;
    }
    auto coerced_b = coerce_to_bool(sc_v);
    if (!coerced_b) {
      error_out = Value::error(coerced_b.error());
      return false;
    }
    scan_by_column = coerced_b.value();
  }
  return true;
}

/// Build a 1xN or Nx1 ArrayValue from the gathered cells. Returns nullptr
/// on arena OOM. Callers that need to surface a specific error should
/// pre-validate and not call this with `cells.empty()`.
ArrayValue* materialise_vector(std::vector<Value>&& cells, bool as_column, Arena& arena) {
  const auto n = static_cast<std::uint32_t>(cells.size());
  Value* buffer = arena.create_array<Value>(cells.size());
  if (buffer == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < cells.size(); ++i) {
    buffer[i] = cells[i];
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return nullptr;
  }
  if (as_column) {
    out->rows = n;
    out->cols = 1;
  } else {
    out->rows = 1;
    out->cols = n;
  }
  out->cells = buffer;
  return out;
}

}  // namespace

Value eval_tocol_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  std::int64_t ignore_mask = 0;
  bool scan_by_column = false;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_tocol_torow_options(call, arity, arena, registry, ctx, ignore_mask, scan_by_column, err)) {
    return err;
  }

  std::vector<Value> kept;
  kept.reserve(static_cast<std::size_t>(array->rows) * static_cast<std::size_t>(array->cols));
  if (!collect_tocol_torow_cells(*array, ignore_mask, scan_by_column, kept, err)) {
    return err;
  }
  if (kept.empty()) {
    return Value::error(ErrorCode::Calc);
  }
  ArrayValue* out = materialise_vector(std::move(kept), /*as_column=*/true, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

Value eval_torow_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 3U) {
    return Value::error(ErrorCode::Value);
  }
  const Value array_v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (array_v.is_error()) {
    return array_v;
  }
  if (!array_v.is_array()) {
    return Value::error(ErrorCode::Value);
  }
  const ArrayValue* array = array_v.as_array();

  std::int64_t ignore_mask = 0;
  bool scan_by_column = false;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_tocol_torow_options(call, arity, arena, registry, ctx, ignore_mask, scan_by_column, err)) {
    return err;
  }

  std::vector<Value> kept;
  kept.reserve(static_cast<std::size_t>(array->rows) * static_cast<std::size_t>(array->cols));
  if (!collect_tocol_torow_cells(*array, ignore_mask, scan_by_column, kept, err)) {
    return err;
  }
  if (kept.empty()) {
    return Value::error(ErrorCode::Calc);
  }
  ArrayValue* out = materialise_vector(std::move(kept), /*as_column=*/false, arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

namespace {

/// WRAPCOLS / WRAPROWS share an arg layout: a 1D `vector`, an integer
/// `wrap_count >= 1`, and an optional scalar `pad_with` (default #N/A).
/// Returns true and writes `flat`, `wrap_count`, `pad` on success;
/// `false` and the caller-visible error otherwise.
bool resolve_wrap_args(const parser::AstNode& call, std::uint32_t arity, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, const ArrayValue*& vector_arr, std::uint32_t& wrap_count, Value& pad,
                       Value& error_out) {
  if (arity < 2U || arity > 3U) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const Value v = eval_node_as_array(call.as_call_arg(0), arena, registry, ctx);
  if (v.is_error()) {
    error_out = v;
    return false;
  }
  if (!v.is_array()) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  const ArrayValue* arr = v.as_array();
  // `vector` must be 1D in either orientation; a 2D rectangle is rejected
  // here (the WRAP family is for vectors only).
  if (arr->rows != 1U && arr->cols != 1U) {
    error_out = Value::error(ErrorCode::Value);
    return false;
  }
  vector_arr = arr;

  const Value wc_v = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (wc_v.is_error()) {
    error_out = wc_v;
    return false;
  }
  auto wc_c = coerce_to_number(wc_v);
  if (!wc_c) {
    error_out = Value::error(wc_c.error());
    return false;
  }
  const double wc_d = std::trunc(wc_c.value());
  if (wc_d < 1.0) {
    error_out = Value::error(ErrorCode::Num);
    return false;
  }
  wrap_count = static_cast<std::uint32_t>(wc_d);

  pad = Value::error(ErrorCode::NA);
  if (arity == 3U) {
    const Value p_v = eval_node(call.as_call_arg(2), arena, registry, ctx);
    if (p_v.is_error()) {
      error_out = p_v;
      return false;
    }
    if (p_v.is_array()) {
      const ArrayValue* pa = p_v.as_array();
      if (pa->rows == 0U || pa->cols == 0U) {
        error_out = Value::error(ErrorCode::Value);
        return false;
      }
      pad = pa->cells[0];
    } else {
      pad = p_v;
    }
  }
  return true;
}

}  // namespace

Value eval_wraprows_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  const ArrayValue* vector_arr = nullptr;
  std::uint32_t wrap_count = 0;
  Value pad = Value::error(ErrorCode::NA);
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_wrap_args(call, arity, arena, registry, ctx, vector_arr, wrap_count, pad, err)) {
    return err;
  }
  // The vector's natural flat order is row-major, which is what we want
  // both for a row vector (1xN) and a column vector (Nx1) — both have
  // their cells already laid out left-to-right / top-to-bottom in
  // `cells` (since rows or cols == 1). Total cell count is rows * cols.
  const std::size_t n = static_cast<std::size_t>(vector_arr->rows) * static_cast<std::size_t>(vector_arr->cols);
  // Output rows is ceil(n / wrap_count); output cols is wrap_count.
  // The trailing row is padded with `pad` if the division leaves a remainder.
  const std::uint32_t out_cols = wrap_count;
  const std::uint32_t out_rows =
      static_cast<std::uint32_t>((n + static_cast<std::size_t>(wrap_count) - 1) / wrap_count);
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = (i < n) ? vector_arr->cells[i] : pad;
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

Value eval_wrapcols_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  const ArrayValue* vector_arr = nullptr;
  std::uint32_t wrap_count = 0;
  Value pad = Value::error(ErrorCode::NA);
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_wrap_args(call, arity, arena, registry, ctx, vector_arr, wrap_count, pad, err)) {
    return err;
  }
  const std::size_t n = static_cast<std::size_t>(vector_arr->rows) * static_cast<std::size_t>(vector_arr->cols);
  // Output rows is wrap_count; output cols is ceil(n / wrap_count). Cells
  // fill column-major: the i-th input cell goes to (i % wrap_count, i /
  // wrap_count). Trailing slots in the final column become `pad`.
  const std::uint32_t out_rows = wrap_count;
  const std::uint32_t out_cols =
      static_cast<std::uint32_t>((n + static_cast<std::size_t>(wrap_count) - 1) / wrap_count);
  const std::size_t total = static_cast<std::size_t>(out_rows) * static_cast<std::size_t>(out_cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < total; ++i) {
    const std::size_t r = i / out_cols;
    const std::size_t c = i % out_cols;
    const std::size_t flat = c * out_rows + r;
    buffer[i] = (flat < n) ? vector_arr->cells[flat] : pad;
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = out_rows;
  out->cols = out_cols;
  out->cells = buffer;
  return Value::array(out);
}

Value eval_anchorarray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);

  // Resolve the argument AST to (sheet_qualifier, top-left cell). The spill
  // anchor is always a single cell, so we ignore any range extent that
  // `resolve_reference_call` may report for OFFSET / INDIRECT — Excel treats
  // `_xlfn.ANCHORARRAY(B2:D5)` the same as `ANCHORARRAY(B2)`: the anchor
  // is whichever cell sits at the top-left.
  std::string_view sheet_name;
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  switch (arg.kind()) {
    case parser::NodeKind::Ref: {
      const parser::Reference& r = arg.as_ref();
      if (r.is_full_col || r.is_full_row) {
        return Value::error(ErrorCode::Ref);
      }
      sheet_name = r.sheet;
      anchor_row = r.row;
      anchor_col = r.col;
      break;
    }
    case parser::NodeKind::SpillRef: {
      const parser::Reference& r = arg.as_spill_ref();
      sheet_name = r.sheet;
      anchor_row = r.row;
      anchor_col = r.col;
      break;
    }
    case parser::NodeKind::Call: {
      std::uint32_t bottom_row = 0;
      std::uint32_t right_col = 0;
      bool is_range = false;
      ErrorCode err = ErrorCode::Ref;
      if (!resolve_reference_call(arg, arena, registry, ctx, &sheet_name, &anchor_row, &anchor_col, &bottom_row,
                                  &right_col, &is_range, &err)) {
        return Value::error(err);
      }
      break;
    }
    default:
      return Value::error(ErrorCode::Value);
  }

  const Sheet* current = ctx.current_sheet();
  if (current == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  const Sheet* target = current;
  if (!sheet_name.empty()) {
    const Workbook* wb = ctx.workbook();
    if (wb == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
    target = wb->sheet_by_name(sheet_name);
    if (target == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
  }
  if (anchor_row >= Sheet::kMaxRows || anchor_col >= Sheet::kMaxCols) {
    return Value::error(ErrorCode::Ref);
  }
  const SpillRegion* region = target->spill_region_at_anchor(anchor_row, anchor_col);
  if (region == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  const std::size_t n = static_cast<std::size_t>(region->rows) * static_cast<std::size_t>(region->cols);
  Value* buffer = arena.create_array<Value>(n);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = region->cells[i];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = region->rows;
  arr->cols = region->cols;
  arr->cells = buffer;
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
