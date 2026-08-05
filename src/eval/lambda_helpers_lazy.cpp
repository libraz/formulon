
#include "eval/lambda_helpers_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/coerce.h"
#include "eval/dynamic_array_limits.h"
#include "eval/eval_context.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Excel worksheet grid limits. MAKEARRAY rejects shapes that exceed either
// dimension with `#NUM!` so the helper cannot be used to allocate an array
// the surrounding workbook could never contain.
constexpr std::uint32_t kExcelMaxRows = 1048576U;
constexpr std::uint32_t kExcelMaxCols = 16384U;

// Wraps a freshly populated `Value` buffer into an arena-allocated
// `ArrayValue`. The buffer must already live in `arena`. Returns a scalar
// `#NUM!` if either allocation fails.
Value wrap_array(Arena& arena, std::uint32_t rows, std::uint32_t cols, Value* buffer) {
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = buffer;
  return Value::array(arr);
}

// Allocates a fresh row-major `Value` buffer of size `rows * cols` in
// `arena`. Returns `nullptr` on allocation failure (caller surfaces
// `#NUM!`). The buffer is not initialised; the caller writes every cell.
Value* alloc_cells(Arena& arena, std::uint32_t rows, std::uint32_t cols) {
  const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  return arena.create_array<Value>(n);
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

// Builds a fresh 1-cell-wide / 1-cell-tall `ArrayValue` carrying `cells`.
// The buffer must already live in `arena`. Used by BYROW / BYCOL to hand a
// row or column slice of the input to the per-cell lambda invocation. The
// resulting `Value::Array` is a perfectly normal Array value: callers can
// `as_array()` / `as_array_cells()` it identically to any other array.
Value make_slice_array(Arena& arena, std::uint32_t rows, std::uint32_t cols, Value* cells) {
  return wrap_array(arena, rows, cols, cells);
}

// Builds a synthetic `ArrayLiteral` AST that mirrors the cells of `arr`.
// Used to give a per-row / per-column slice a range-shaped AST identity,
// so range-aware functions inside the lambda body (`SUM`, `AVERAGE`, ...)
// can flatten the slice through the dispatcher's existing ArrayLiteral
// branch instead of receiving an opaque `Value::Array` they cannot coerce.
//
// Each cell becomes a `Literal` AST node carrying the cell's `Value`. The
// resulting AST and all child nodes live in `arena` for the same lifetime
// as the lambda invocation. Returns `nullptr` on allocation failure.
const parser::AstNode* build_array_literal_for(const ArrayValue* arr, Arena& arena) {
  const std::uint32_t rows = arr->rows;
  const std::uint32_t cols = arr->cols;
  // `make_array_literal` precondition: rows / cols must be >= 1. The
  // empty-input guards in BYROW / BYCOL / SCAN ensure this never triggers
  // for a valid slice.
  if (rows == 0U || cols == 0U) {
    return nullptr;
  }
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  const parser::AstNode** children = arena.create_array<const parser::AstNode*>(total);
  if (children == nullptr) {
    return nullptr;
  }
  for (std::size_t i = 0; i < total; ++i) {
    parser::AstNode* lit = parser::make_literal(arena, arr->cells[i]);
    if (lit == nullptr) {
      return nullptr;
    }
    children[i] = lit;
  }
  return parser::make_array_literal(arena, rows, cols, children);
}

// Invokes `lv` with the pre-evaluated `args` values, binding each into a
// fresh frame rooted at the lambda's captured environment. Mirrors the
// AST-driven `invoke_lambda` in `tree_walker.cpp` but takes already-
// computed `Value`s rather than argument AST nodes — the per-cell helpers
// have no AST for their cell values, just the cell payload itself.
//
// `ast_args` is an optional parallel array of AST nodes (length `arity`)
// to record alongside each binding. When non-null, the AST node lets
// range-aware consumers inside the lambda body see the binding as a
// range-shaped expression (matching the LET-binding-passthrough seam used
// throughout the eager dispatcher and the lazy lookup family). Pass
// `nullptr` to fall back to scalar bindings.
//
// Strict arity match (Excel's `#VALUE!` policy). Argument errors are NOT
// short-circuited here — the caller is expected to filter them first if
// required (REDUCE / SCAN want to surface array-cell errors verbatim, but
// MAP / BYROW / BYCOL pre-check). The lambda body is evaluated in the
// caller's `EvalContext` with the caller-supplied `NameEnv` swapped for
// the freshly extended frame.
Value invoke_lambda_with_values(const LambdaValue* lv, const Value* args, const parser::AstNode* const* ast_args,
                                std::uint32_t arity, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  if (arity != lv->param_count) {
    return Value::error(ErrorCode::Value);
  }
  NameEnv env;
  if (lv->captured_env != nullptr) {
    env = *lv->captured_env;
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode* expr = (ast_args != nullptr) ? ast_args[i] : nullptr;
    env = env.extend(lv->params[i], args[i], expr, arena);
  }
  const EvalContext body_ctx = ctx.with_name_env(&env);
  return eval_node(*lv->body, arena, registry, body_ctx);
}

// Evaluates a single argument as a `LambdaValue*` closure. Returns nullptr
// and writes a scalar error to `*out_err` on failure paths:
//   * argument error -> propagate verbatim;
//   * argument is not a Lambda -> `#VALUE!`.
// `expected_arity` is the lambda parameter count required by the caller.
// `expected_arity == 0` means "any" (only used by MAP, which determines
// the required arity from the number of array arguments). Otherwise
// arity mismatch surfaces `#VALUE!`.
const LambdaValue* eval_lambda_arg(const parser::AstNode& node, std::uint32_t expected_arity, Arena& arena,
                                   const FunctionRegistry& registry, const EvalContext& ctx, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    *out_err = v;
    return nullptr;
  }
  if (!v.is_lambda()) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  const LambdaValue* lv = v.as_lambda();
  if (expected_arity != 0U && lv->param_count != expected_arity) {
    *out_err = Value::error(ErrorCode::Value);
    return nullptr;
  }
  return lv;
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

// Coerces a scalar argument to a positive integer count for MAKEARRAY's
// `rows` and `cols`. Negative / zero / non-numeric / out-of-grid values
// surface `#NUM!`. Argument errors propagate verbatim.
//
// `max_value` is the per-axis Excel grid limit. The numeric value is
// truncated toward zero (Excel rounds `2.9` down to 2 here, matching the
// behaviour observed for ROUND-free integer args across the dynamic-array
// helpers).
bool read_count_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    std::uint32_t max_value, std::uint32_t* out, Value* out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
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
  if (truncated < 1.0 || truncated > static_cast<double>(max_value)) {
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
  const LambdaValue* lv = eval_lambda_arg(call.as_call_arg(1), /*expected_arity=*/1U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }
  if (in->rows == 0U || in->cols == 0U) {
    // Mac Excel: an empty input has no row / column to apply the lambda to.
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

  Value* out_cells = alloc_cells(arena, out_rows, out_cols);
  if (out_cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  for (std::uint32_t i = 0; i < iter_count; ++i) {
    Value* slice_buf = alloc_cells(arena, slice_rows, slice_cols);
    if (slice_buf == nullptr) {
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
    const Value slice = make_slice_array(arena, slice_rows, slice_cols, slice_buf);
    Value arg = slice;
    // Bind the slice with both the Value and a synthetic ArrayLiteral AST
    // so range-aware functions inside the body (`SUM(r)`, `AVERAGE(r)`,
    // ...) flatten the slice through the dispatcher's ArrayLiteral branch
    // instead of receiving an opaque `Value::Array` they cannot coerce.
    const parser::AstNode* slice_ast = build_array_literal_for(slice.as_array(), arena);
    if (slice_ast == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    const parser::AstNode* ast_args[1] = {slice_ast};
    const Value res = invoke_lambda_with_values(lv, &arg, ast_args, 1U, arena, registry, ctx);
    if (res.is_error()) {
      return res;
    }
    // Mac Excel anchor-unwraps a 1x1 Array lambda result (e.g.
    // `LAMBDA(row, row*10)` applied to a 1x1 row slice produces {10},
    // which Excel projects to 10). Multi-cell Arrays and lambda values
    // still surface #CALC! because BYROW / BYCOL have no slot to spill
    // a sub-array or closure into.
    const Value scalar_res = unwrap_1x1_array(res);
    if (scalar_res.is_array() || scalar_res.is_lambda()) {
      return Value::error(ErrorCode::Calc);
    }
    out_cells[i] = scalar_res;
  }

  return wrap_array(arena, out_rows, out_cols, out_cells);
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

  // The lambda is the last positional argument. Its parameter count must
  // equal the number of array arguments; mismatch surfaces #VALUE!.
  const LambdaValue* lv =
      eval_lambda_arg(call.as_call_arg(arity - 1U), /*expected_arity=*/array_count, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }

  if (rows == 0U || cols == 0U) {
    return Value::error(ErrorCode::Calc);
  }

  Value* out_cells = alloc_cells(arena, rows, cols);
  if (out_cells == nullptr) {
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
      } else {
        args[k] = arrays[k]->cells[static_cast<std::size_t>(r) * arrays[k]->cols + c];
      }
    }
    // MAP per-cell args are scalars from the input arrays; no AST hint
    // needed because range-aware functions inside the lambda body would
    // see a single cell either way.
    const Value res = invoke_lambda_with_values(lv, args, /*ast_args=*/nullptr, array_count, arena, registry, ctx);
    if (res.is_error()) {
      out_cells[idx] = res;
      continue;
    }
    // Anchor-unwrap a 1x1 Array result; multi-cell Arrays / lambda values
    // still surface #CALC!.
    const Value scalar_res = unwrap_1x1_array(res);
    if (scalar_res.is_array() || scalar_res.is_lambda()) {
      return Value::error(ErrorCode::Calc);
    }
    out_cells[idx] = scalar_res;
  }

  return wrap_array(arena, rows, cols, out_cells);
}

Value eval_reduce_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  // initial_value is evaluated as a plain scalar Value; an error propagates.
  Value acc = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (acc.is_error()) {
    return acc;
  }
  Value err = Value::blank();
  const ArrayValue* in = eval_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (in == nullptr) {
    return err;
  }
  const LambdaValue* lv = eval_lambda_arg(call.as_call_arg(2), /*expected_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }

  // An empty input is a no-op fold: REDUCE returns the seed unchanged
  // (the Mac Excel observed behaviour for `=REDUCE(0, FILTER(...empty...), ...)`).
  const std::size_t total = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  if (total == 0U) {
    return acc;
  }

  Value args[2] = {Value::blank(), Value::blank()};
  for (std::size_t i = 0; i < total; ++i) {
    args[0] = acc;
    args[1] = in->cells[i];
    // An error in the current cell short-circuits the fold (matches Mac
    // Excel's spill semantics: a single bad cell taints the result).
    if (args[1].is_error()) {
      return args[1];
    }
    // REDUCE feeds scalar (accumulator, current) per call. The accumulator
    // can be any Value but is consumed inside the body via the parameter
    // name lookup, not as a range-aware seam, so no AST hint is required.
    const Value res = invoke_lambda_with_values(lv, args, /*ast_args=*/nullptr, 2U, arena, registry, ctx);
    if (res.is_error()) {
      return res;
    }
    acc = res;
  }
  return acc;
}

Value eval_scan_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  Value acc = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (acc.is_error()) {
    return acc;
  }
  Value err = Value::blank();
  const ArrayValue* in = eval_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (in == nullptr) {
    return err;
  }
  const LambdaValue* lv = eval_lambda_arg(call.as_call_arg(2), /*expected_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }

  const std::uint32_t rows = in->rows;
  const std::uint32_t cols = in->cols;
  if (rows == 0U || cols == 0U) {
    // SCAN must emit a shape-preserving result; an empty input has no
    // shape to spill into, so #CALC! mirrors BYROW / BYCOL.
    return Value::error(ErrorCode::Calc);
  }

  Value* out_cells = alloc_cells(arena, rows, cols);
  if (out_cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  Value args[2] = {Value::blank(), Value::blank()};
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < total; ++i) {
    args[0] = acc;
    args[1] = in->cells[i];
    if (args[1].is_error()) {
      return args[1];
    }
    // SCAN feeds scalar (accumulator, current) per call. No AST hint is
    // needed; the body sees both bindings as scalar Values.
    const Value res = invoke_lambda_with_values(lv, args, /*ast_args=*/nullptr, 2U, arena, registry, ctx);
    if (res.is_error()) {
      return res;
    }
    // SCAN, like BYROW / BYCOL / MAP, has a single output slot per cell;
    // a multi-cell lambda return would require nested spilling. A 1x1
    // Array result is anchor-unwrapped to its single cell.
    const Value scalar_res = unwrap_1x1_array(res);
    if (scalar_res.is_array() || scalar_res.is_lambda()) {
      return Value::error(ErrorCode::Calc);
    }
    acc = scalar_res;
    out_cells[i] = acc;
  }

  return wrap_array(arena, rows, cols, out_cells);
}

Value eval_makearray_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  Value err = Value::blank();
  std::uint32_t rows = 0;
  if (!read_count_arg(call.as_call_arg(0), arena, registry, ctx, kExcelMaxRows, &rows, &err)) {
    return err;
  }
  std::uint32_t cols = 0;
  if (!read_count_arg(call.as_call_arg(1), arena, registry, ctx, kExcelMaxCols, &cols, &err)) {
    return err;
  }
  if (static_cast<std::uint64_t>(rows) * static_cast<std::uint64_t>(cols) > kMaxSequenceCells) {
    return Value::error(ErrorCode::Num);
  }
  const LambdaValue* lv = eval_lambda_arg(call.as_call_arg(2), /*expected_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }

  Value* out_cells = alloc_cells(arena, rows, cols);
  if (out_cells == nullptr) {
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
      const Value res = invoke_lambda_with_values(lv, args, /*ast_args=*/nullptr, 2U, arena, registry, ctx);
      if (res.is_error()) {
        out_cells[static_cast<std::size_t>(r) * static_cast<std::size_t>(cols) + c] = res;
        continue;
      }
      // Anchor-unwrap a 1x1 Array result; multi-cell Arrays / lambda
      // values still surface #CALC!.
      const Value scalar_res = unwrap_1x1_array(res);
      if (scalar_res.is_array() || scalar_res.is_lambda()) {
        return Value::error(ErrorCode::Calc);
      }
      out_cells[static_cast<std::size_t>(r) * static_cast<std::size_t>(cols) + c] = scalar_res;
    }
  }

  return wrap_array(arena, rows, cols, out_cells);
}

}  // namespace eval
}  // namespace formulon
