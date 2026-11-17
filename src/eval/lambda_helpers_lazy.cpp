
#include "eval/lambda_helpers_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/array_alloc.h"
#include "eval/array_context.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/omitted_arg.h"
#include "eval/reference/ref_value.h"
#include "eval/scalar_args.h"
#include "eval/structured_ref_project.h"
#include "eval/tree_walker/dispatch.h"
#include "parser/ast.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Allocates the `(rows, cols)` result array through the evaluator's shared
// seam, handing back the header to publish once the cells are written and,
// via `out_cells`, the uninitialised row-major buffer the caller fills.
// Returns `nullptr` when the shape is rejected or the arena is exhausted
// (caller surfaces `#NUM!`).
ArrayValue* alloc_cells(Arena& arena, std::uint32_t rows, std::uint32_t cols, Value*& out_cells) {
  return allocate_derived_array_value(rows, cols, arena, out_cells);
}

// If `v` is a 1x1 Array, returns its single cell unchanged. Otherwise
// returns `v` verbatim.
//
// BYROW / BYCOL / MAP / SCAN / MAKEARRAY require their lambda body to
// produce a scalar per output slot. Mac Excel's lambda dispatch
// implicitly anchor-unwraps a 1x1 Array result (e.g. `LAMBDA(row, row*10)`
// applied to a 1x1 row slice produces {10}, which Excel projects to 10).
// Multi-cell Arrays still surface #CALC! at the call site.
Value unwrap_1x1_array(const Value& v) {
  if (!v.is_array()) {
    return v;
  }
  const ArrayValue* a = v.as_array();
  if (a->rows == 1U && a->cols == 1U) {
    return a->cells[0];
  }
  return v;
}

// Stores a lambda result into one output slot: an error lands verbatim, a 1x1
// Array is anchor-unwrapped. False when the result is a multi-cell Array or a
// lambda value, which has no single-cell slot; the caller surfaces #CALC!.
bool store_scalar_result(const Value& res, Value* slot) {
  if (res.is_error()) {
    *slot = res;
    return true;
  }
  const Value scalar_res = unwrap_1x1_array(res);
  if (scalar_res.is_array() || scalar_res.is_lambda()) {
    return false;
  }
  *slot = scalar_res;
  return true;
}

// Where a helper's source argument lies on the grid when it is a reference.
// Excel binds each callback argument taken from such a source as the cell
// (MAP, REDUCE, SCAN) or the row / column (BYROW, BYCOL) it came from, so
// ISREF, ROW or OFFSET inside the body see a reference. `valid` is false for
// an array or computed source, whose elements bind by value. A grid of
// references (`grid`) binds each element as its own reference, a value
// element as its value, and a row / column as a grid slice.
struct SourceOrigin {
  bool valid = false;
  parser::Reference top_left{};
  Value grid = Value::blank();
};

SourceOrigin source_origin(const parser::AstNode& node, const ArrayValue* arr, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx) {
  SourceOrigin out;
  if (node.kind() == parser::NodeKind::Literal && node.as_literal().is_ref() &&
      node.as_literal().ref_form() == RefForm::kGrid) {
    out.grid = node.as_literal();
    return out;
  }
  const parser::AstNode* ref = resolve_binding_reference(node, arena, registry, ctx);
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (ref == nullptr || !reference_endpoint_pair(*ref, &lhs, &rhs, ctx)) {
    return out;
  }
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  // The evaluated array must lie inside the rectangle for every element to
  // have a cell of its own.
  if (!rect || arr->rows > rect.value().rows() || arr->cols > rect.value().cols()) {
    return out;
  }
  out.valid = true;
  out.top_left.sheet = lhs.sheet;
  out.top_left.sheet_quoted = lhs.sheet_quoted;
  out.top_left.row = rect.value().row_first;
  out.top_left.col = rect.value().col_first;
  return out;
}

// The `RangeOp` over rows `[r1, r2]` x columns `[c1, c2]` of `origin`'s source
// (a single cell stays a one-cell `RangeOp`, the shape a resolved reference
// binds as), or that block of a grid source. Returns nullptr when the source
// is not a reference or the arena is exhausted.
const parser::AstNode* origin_rect_ast(const SourceOrigin& origin, std::uint32_t r1, std::uint32_t c1, std::uint32_t r2,
                                       std::uint32_t c2, Arena& arena) {
  if (origin.grid.is_ref()) {
    return parser::make_literal(arena, grid_slice(origin.grid, r1, c1, r2 - r1 + 1U, c2 - c1 + 1U, arena));
  }
  if (!origin.valid) {
    return nullptr;
  }
  parser::Reference first = origin.top_left;
  parser::Reference last = origin.top_left;
  first.row += r1;
  first.col += c1;
  last.row += r2;
  last.col += c2;
  parser::AstNode* lhs = parser::make_ref(arena, first);
  parser::AstNode* rhs = parser::make_ref(arena, last);
  if (lhs == nullptr || rhs == nullptr) {
    return nullptr;
  }
  return parser::make_range_op(arena, lhs, rhs, /*spliced=*/false);
}

// Binds row-major element `i` of `arr`: `*ast` its one-cell reference (nullptr
// for a non-reference source), `*value` the value it binds otherwise. A grid
// element binds as its own reference, a value or error element as itself.
void bind_element(const SourceOrigin& origin, const ArrayValue* arr, std::size_t i, Arena& arena, Value* value,
                  const parser::AstNode** ast) {
  if (origin.grid.is_ref()) {
    const Value element = grid_element(origin.grid.as_ref(), i, arena);
    *value = element.is_ref() ? arr->cells[i] : element;
    *ast = element.is_ref() ? parser::make_literal(arena, element) : nullptr;
    return;
  }
  const auto r = static_cast<std::uint32_t>(i / arr->cols);
  const auto c = static_cast<std::uint32_t>(i % arr->cols);
  *value = arr->cells[i];
  *ast = origin_rect_ast(origin, r, c, r, c, arena);
}

// Evaluates an argument in array context, returning the `ArrayValue*` on
// success. On failure paths (argument error, non-array result) writes the
// appropriate scalar error to `*out_err` and returns nullptr.
const ArrayValue* eval_array_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                                 const EvalContext& ctx, Value* out_err) {
  const Value v = eval_node_as_array(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return nullptr;
  }
  if (!v.is_array()) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  return v.as_array();
}

// A count argument's value: an omitted count is 1.
Value count_arg_value(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return is_omitted_arg(node) ? Value::number(1.0) : eval_node(node, arena, registry, ctx);
}

// Coerces a scalar to a positive integer count for MAKEARRAY's `rows` and
// `cols`, truncated toward zero; a blank coerces to 0. Non-numeric or below-1
// values are #VALUE!, past `max_value` (the per-axis grid limit) #NUM!, and an
// error is itself.
bool count_from_value(const Value& v, std::uint32_t max_value, std::uint32_t* out, Value* out_err) {
  if (v.is_error()) {
    *out_err = v;
    return false;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  const double n = coerced.value();
  if (std::isnan(n) || std::isinf(n)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  const double truncated = std::trunc(n);
  if (truncated < 1.0) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  if (truncated > static_cast<double>(max_value)) {
    *out_err = Value::error(ErrorCode::Num);
    return false;
  }
  *out = static_cast<std::uint32_t>(truncated);
  return true;
}

// Implements BYROW (axis = rows) and BYCOL (axis = cols). The axis
// parameter selects which dimension we iterate along: BYROW emits one
// scalar per row, BYCOL emits one scalar per column. The output shape is
// `(rows, 1)` for BYROW and `(1, cols)` for BYCOL.
//
// The callable argument, like every lambda helper's, resolves through
// `resolve_callable`, so a bare built-in name (`BYROW(data, SUM)`) is the
// eta-reduced lambda Excel reads it as.
Value byrow_or_bycol(const parser::AstNode& call, bool by_row, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 2U) {
    return Value::error(ErrorCode::Value);
  }
  Value err = Value::blank();
  const ArrayValue* in = eval_array_arg(call.as_call_arg(0), arena, registry, ctx, &err);
  if (in == nullptr) {
    return err;
  }
  const LambdaValue* lv = resolve_callable(call.as_call_arg(1), /*call_arity=*/1U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }
  const SourceOrigin origin = source_origin(call.as_call_arg(0), in, arena, registry, ctx);
  if (in->rows == 0U || in->cols == 0U) {
    // Mac Excel: an empty input has no row / column to apply the lambda to.
    note_unknown_error_origin();
    return Value::error(ErrorCode::Calc);
  }

  const std::uint32_t rows_in = in->rows;
  const std::uint32_t cols_in = in->cols;
  const std::uint32_t out_rows = by_row ? rows_in : 1U;
  const std::uint32_t out_cols = by_row ? 1U : cols_in;
  const std::uint32_t iter_count = by_row ? rows_in : cols_in;
  const std::uint32_t slice_rows = by_row ? 1U : rows_in;
  const std::uint32_t slice_cols = by_row ? cols_in : 1U;
  const std::uint32_t slice_size = by_row ? cols_in : rows_in;

  Value* out_cells = nullptr;
  ArrayValue* out_arr = alloc_cells(arena, out_rows, out_cols, out_cells);
  if (out_arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  for (std::uint32_t i = 0; i < iter_count; ++i) {
    Value* slice_buf = nullptr;
    ArrayValue* slice_arr = alloc_cells(arena, slice_rows, slice_cols, slice_buf);
    if (slice_arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    if (by_row) {
      const std::size_t base = static_cast<std::size_t>(i) * static_cast<std::size_t>(cols_in);
      for (std::uint32_t c = 0; c < slice_size; ++c) {
        slice_buf[c] = in->cells[base + c];
      }
    } else {
      for (std::uint32_t r = 0; r < slice_size; ++r) {
        slice_buf[r] = in->cells[static_cast<std::size_t>(r) * static_cast<std::size_t>(cols_in) + i];
      }
    }
    const Value slice = Value::array(slice_arr);
    // A reference source binds the slice as its row / column reference;
    // otherwise a synthetic ArrayLiteral AST lets range-aware functions
    // inside the body (`SUM(r)`, ...) flatten the slice through the
    // dispatcher's ArrayLiteral branch.
    const parser::AstNode* slice_ast = origin.valid || origin.grid.is_ref()
                                           ? (by_row ? origin_rect_ast(origin, i, 0U, i, cols_in - 1U, arena)
                                                     : origin_rect_ast(origin, 0U, i, rows_in - 1U, i, arena))
                                           : array_literal_ast(slice_arr, arena);
    if (slice_ast == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    const parser::AstNode* ast_args[1] = {slice_ast};
    const Value res = invoke_lambda_values_with_ast(lv, 1U, &slice, ast_args, arena, registry, ctx);
    // Each row / column is reduced independently, so an error lands in that
    // slice's output cell only. Short-circuiting the whole call would hide
    // the other slices' results, which Excel still spills. A 1x1 Array is
    // anchor-unwrapped; a multi-cell Array or lambda has no slot (#CALC!).
    if (!store_scalar_result(res, &out_cells[i])) {
      note_calc_sub_type(2U);
      return Value::error(ErrorCode::Calc);
    }
  }

  return Value::array(out_arr);
}

/// Shared inputs of REDUCE / SCAN: seed, source array and the 2-arity callable.
struct FoldInputs {
  Value acc = Value::blank();
  const ArrayValue* in = nullptr;
  const LambdaValue* lv = nullptr;
  SourceOrigin origin;
};

/// Reads `(initial_value, array, lambda)`; on failure returns false with the error in `*out_err`.
bool read_fold_inputs(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, FoldInputs* fi, Value* out_err) {
  if (call.as_call_arity() != 3U) {
    *out_err = Value::error(ErrorCode::Value);
    return false;
  }
  // The seed binds like a LAMBDA argument, so a reference seed stays a reference for the first call.
  fi->acc = eval_binding_source(call.as_call_arg(0), arena, registry, ctx);
  if (fi->acc.is_error()) {
    *out_err = fi->acc;
    return false;
  }
  fi->in = eval_array_arg(call.as_call_arg(1), arena, registry, ctx, out_err);
  if (fi->in == nullptr) {
    return false;
  }
  fi->lv = resolve_callable(call.as_call_arg(2), /*call_arity=*/2U, arena, registry, ctx, out_err);
  if (fi->lv == nullptr) {
    return false;
  }
  fi->origin = source_origin(call.as_call_arg(1), fi->in, arena, registry, ctx);
  return true;
}

}  // namespace

Value eval_byrow_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return byrow_or_bycol(call, /*by_row=*/true, arena, registry, ctx);
}

Value eval_bycol_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  return byrow_or_bycol(call, /*by_row=*/false, arena, registry, ctx);
}

Value eval_map_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // MAP requires at least one array and one lambda (the trailing arg).
  if (arity < 2U) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t array_count = arity - 1U;

  // Resolve every array argument first; mismatched shapes surface #N/A as
  // documented for MAP. An argument-evaluation error short-circuits the
  // whole call and propagates verbatim.
  Value err = Value::blank();
  // Stack-arena-friendly: use a small arena-backed buffer rather than
  // std::vector to keep this in the hot path's memory profile.
  const ArrayValue** arrays = arena.create_array<const ArrayValue*>(array_count);
  if (arrays == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < array_count; ++i) {
    const ArrayValue* a = eval_array_arg(call.as_call_arg(i), arena, registry, ctx, &err);
    if (a == nullptr) {
      return err;
    }
    arrays[i] = a;
  }

  std::uint32_t rows = arrays[0]->rows;
  std::uint32_t cols = arrays[0]->cols;
  for (std::uint32_t i = 1; i < array_count; ++i) {
    rows = std::max(rows, arrays[i]->rows);
    cols = std::max(cols, arrays[i]->cols);
  }

  // The lambda is the last positional argument. It must be able to accept
  // one argument per array — required params no more than `array_count`,
  // declared params no fewer; anything else surfaces #VALUE!.
  const LambdaValue* lv =
      resolve_callable(call.as_call_arg(arity - 1U), /*call_arity=*/array_count, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }
  SourceOrigin* origins = arena.create_array<SourceOrigin>(array_count);
  const parser::AstNode** arg_asts = arena.create_array<const parser::AstNode*>(array_count);
  if (origins == nullptr || arg_asts == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < array_count; ++i) {
    origins[i] = source_origin(call.as_call_arg(i), arrays[i], arena, registry, ctx);
  }

  if (rows == 0U || cols == 0U) {
    note_unknown_error_origin();
    return Value::error(ErrorCode::Calc);
  }

  Value* out_cells = nullptr;
  ArrayValue* out_arr = alloc_cells(arena, rows, cols, out_cells);
  if (out_arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  // Per-cell argument buffer reused across the iteration. Each cell call
  // gets `array_count` arguments — one element from each input array at
  // the current `(r, c)` coordinate.
  Value* args = arena.create_array<Value>(array_count);
  if (args == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t idx = 0; idx < total; ++idx) {
    const std::uint32_t r = static_cast<std::uint32_t>(idx / cols);
    const std::uint32_t c = static_cast<std::uint32_t>(idx % cols);
    for (std::uint32_t k = 0; k < array_count; ++k) {
      if (r >= arrays[k]->rows || c >= arrays[k]->cols) {
        args[k] = Value::error(ErrorCode::NA);
        arg_asts[k] = nullptr;
      } else {
        bind_element(origins[k], arrays[k], static_cast<std::size_t>(r) * arrays[k]->cols + c, arena, &args[k],
                     &arg_asts[k]);
      }
    }
    const Value res = invoke_lambda_values_with_ast(lv, array_count, args, arg_asts, arena, registry, ctx);
    if (!store_scalar_result(res, &out_cells[idx])) {
      note_calc_sub_type(1U);
      return Value::error(ErrorCode::Calc);
    }
  }

  return Value::array(out_arr);
}

namespace {

// One REDUCE / SCAN step: the lambda applied to the accumulator and element `i`.
// An errored element reaches the lambda verbatim, so a body that guards with
// IFERROR keeps folding and one that does not leaves the error in the
// accumulator.
Value fold_step(const FoldInputs& fi, std::size_t i, const Value& acc, Arena& arena, const FunctionRegistry& registry,
                const EvalContext& ctx) {
  Value args[2] = {acc, Value::blank()};
  // Only the seed can be a reference (a `Ref` accumulator); every later accumulator is a result.
  const parser::AstNode* arg_asts[2] = {nullptr, nullptr};
  bind_element(fi.origin, fi.in, i, arena, &args[1], &arg_asts[1]);
  return invoke_lambda_values_with_ast(fi.lv, 2U, args, arg_asts, arena, registry, ctx);
}

}  // namespace

Value eval_reduce_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  FoldInputs fi;
  Value err = Value::blank();
  if (!read_fold_inputs(call, arena, registry, ctx, &fi, &err)) {
    return err;
  }
  Value acc = fi.acc;

  // An empty input is a no-op fold: REDUCE returns the seed unchanged
  // (the Mac Excel observed behaviour for `=REDUCE(0, FILTER(...empty...), ...)`).
  const std::size_t total = static_cast<std::size_t>(fi.in->rows) * static_cast<std::size_t>(fi.in->cols);
  if (total == 0U) {
    return acc;
  }

  for (std::size_t i = 0; i < total; ++i) {
    acc = fold_step(fi, i, acc, arena, registry, ctx);
  }
  return acc;
}

Value eval_scan_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  FoldInputs fi;
  Value err = Value::blank();
  if (!read_fold_inputs(call, arena, registry, ctx, &fi, &err)) {
    return err;
  }
  Value acc = fi.acc;

  const std::uint32_t rows = fi.in->rows;
  const std::uint32_t cols = fi.in->cols;
  if (rows == 0U || cols == 0U) {
    // SCAN must emit a shape-preserving result; an empty input has no
    // shape to spill into, so #CALC! mirrors BYROW / BYCOL.
    note_unknown_error_origin();
    return Value::error(ErrorCode::Calc);
  }

  Value* out_cells = nullptr;
  ArrayValue* out_arr = alloc_cells(arena, rows, cols, out_cells);
  if (out_arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < total; ++i) {
    const Value res = fold_step(fi, i, acc, arena, registry, ctx);
    // A single output slot per cell: a multi-cell lambda return would need nested spilling.
    if (!store_scalar_result(res, &out_cells[i])) {
      note_calc_sub_type(1U);
      return Value::error(ErrorCode::Calc);
    }
    acc = out_cells[i];
  }

  return Value::array(out_arr);
}

const LambdaValue* makearray_shape(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                   const EvalContext& ctx, std::uint32_t* out_rows, std::uint32_t* out_cols,
                                   Value* out_err) {
  if (call.as_call_arity() != 3U) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  const Value rows = count_arg_value(call.as_call_arg(0), arena, registry, ctx);
  const Value cols = count_arg_value(call.as_call_arg(1), arena, registry, ctx);
  if (rows.is_array() || cols.is_array()) {
    return nullptr;
  }
  if (!count_from_value(rows, a1::kMaxRows, out_rows, out_err) ||
      !count_from_value(cols, a1::kMaxCols, out_cols, out_err)) {
    return nullptr;
  }
  return resolve_callable(call.as_call_arg(2), /*call_arity=*/2U, arena, registry, ctx, out_err);
}

namespace {

// The `rows` x `cols` grid of `lv(r, c)`.
Value fill_makearray(const LambdaValue* lv, std::uint32_t rows, std::uint32_t cols, Arena& arena,
                     const FunctionRegistry& registry, const EvalContext& ctx) {
  if (abandons_formula(rows, cols) ||
      static_cast<std::uint64_t>(rows) * static_cast<std::uint64_t>(cols) > kMaxDynamicArrayCells) {
    return Value::error(ErrorCode::Num);
  }

  Value* out_cells = nullptr;
  ArrayValue* out_arr = alloc_cells(arena, rows, cols, out_cells);
  if (out_arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  Value args[2] = {Value::blank(), Value::blank()};
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      // Excel uses 1-based indices for the lambda parameters.
      args[0] = Value::number(static_cast<double>(r) + 1.0);
      args[1] = Value::number(static_cast<double>(c) + 1.0);
      // MAKEARRAY feeds scalar (row_index, col_index) numbers; no AST
      // hint is needed because no consumer would ever see them as a range.
      const Value res = invoke_lambda_values(lv, 2U, args, arena, registry, ctx);
      if (!store_scalar_result(res, &out_cells[static_cast<std::size_t>(r) * static_cast<std::size_t>(cols) + c])) {
        note_calc_sub_type(1U);
        return Value::error(ErrorCode::Calc);
      }
    }
  }

  return Value::array(out_arr);
}

}  // namespace

Value eval_makearray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  Value counts[2] = {count_arg_value(call.as_call_arg(0), arena, registry, ctx),
                     count_arg_value(call.as_call_arg(1), arena, registry, ctx)};
  for (const Value& v : counts) {
    if (v.is_error()) {
      return v;
    }
  }
  Value err = Value::blank();
  const LambdaValue* lv = resolve_callable(call.as_call_arg(2), /*call_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }
  // An array of counts lifts the call.
  return lift_over_array_args(counts, 2U, arena, [&](const Value* a, std::uint32_t) {
    std::uint32_t rows = 0;
    std::uint32_t cols = 0;
    Value count_err = Value::blank();
    if (!count_from_value(a[0], a1::kMaxRows, &rows, &count_err) ||
        !count_from_value(a[1], a1::kMaxCols, &cols, &count_err)) {
      return count_err;
    }
    return fill_makearray(lv, rows, cols, arena, registry, ctx);
  });
}

}  // namespace eval
}  // namespace formulon
