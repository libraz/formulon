// Copyright 2026 libraz. Licensed under the MIT License.

#include "eval/dynamic_array_lazy.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <numeric>
#include <vector>

#include "eval/coerce.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
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

}  // namespace eval
}  // namespace formulon
