
#include "eval/lambda_helpers_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/dynamic_array_limits.h"
#include "eval/eval_context.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/shape_ops_lazy.h"
#include "eval/tree_walker/dispatch.h"
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

// Allocates the `(rows, cols)` result array through the evaluator's shared
// seam, handing back the header to publish once the cells are written and,
// via `out_cells`, the uninitialised row-major buffer the caller fills.
// Returns `nullptr` when the shape is rejected or the arena is exhausted
// (caller surfaces `#NUM!`).
ArrayValue* alloc_cells(Arena& arena, std::uint32_t rows, std::uint32_t cols, Value*& out_cells) {
  return allocate_array_value(rows, cols, arena, out_cells, kMaxDerivedArrayCells);
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

// Where a helper's source argument lies on the grid when it is a reference.
// Excel binds each callback argument taken from such a source as the cell
// (MAP, REDUCE, SCAN) or the row / column (BYROW, BYCOL) it came from, so
// ISREF, ROW or OFFSET inside the body see a reference. `valid` is false for
// an array or computed source, whose elements bind by value.
struct SourceOrigin {
  bool valid = false;
  parser::Reference top_left{};
};

SourceOrigin source_origin(const parser::AstNode& node, const ArrayValue* arr, Arena& arena,
                           const FunctionRegistry& registry, const EvalContext& ctx) {
  SourceOrigin out;
  const parser::AstNode* ref = resolve_binding_reference(node, arena, registry, ctx);
  parser::Reference lhs{};
  parser::Reference rhs{};
  if (ref == nullptr || !declared_rect_endpoint_pair(*ref, &lhs, &rhs)) {
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
// binds as). Returns nullptr when the source is not a reference or the arena
// is exhausted.
const parser::AstNode* origin_rect_ast(const SourceOrigin& origin, std::uint32_t r1, std::uint32_t c1, std::uint32_t r2,
                                       std::uint32_t c2, Arena& arena) {
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
  return parser::make_range_op(arena, lhs, rhs);
}

// The one-cell reference of row-major element `i` of `arr`, or nullptr for a
// non-reference source.
const parser::AstNode* element_ast(const SourceOrigin& origin, const ArrayValue* arr, std::size_t i, Arena& arena) {
  const auto r = static_cast<std::uint32_t>(i / arr->cols);
  const auto c = static_cast<std::uint32_t>(i % arr->cols);
  return origin_rect_ast(origin, r, c, r, c, arena);
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
// `rows` and `cols`. Non-numeric or below-1 values surface `#VALUE!`
// (Microsoft's documented contract); only exceeding the per-axis grid
// limit surfaces `#NUM!`. Argument errors propagate verbatim.
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
    const parser::AstNode* slice_ast = origin.valid ? (by_row ? origin_rect_ast(origin, i, 0U, i, cols_in - 1U, arena)
                                                              : origin_rect_ast(origin, 0U, i, rows_in - 1U, i, arena))
                                                    : build_array_literal_for(slice_arr, arena);
    if (slice_ast == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    const parser::AstNode* ast_args[1] = {slice_ast};
    const Value res = invoke_lambda_values_with_ast(lv, 1U, &slice, ast_args, arena, registry, ctx);
    if (res.is_error()) {
      // Each row / column is reduced independently, so an error lands in
      // that slice's output cell only. Short-circuiting the whole call here
      // would hide the other slices' results, which Excel still spills.
      out_cells[i] = res;
      continue;
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

  return Value::array(out_arr);
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
        args[k] = arrays[k]->cells[static_cast<std::size_t>(r) * arrays[k]->cols + c];
        arg_asts[k] = origin_rect_ast(origins[k], r, c, r, c, arena);
      }
    }
    const Value res = invoke_lambda_values_with_ast(lv, array_count, args, arg_asts, arena, registry, ctx);
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

  return Value::array(out_arr);
}

Value eval_reduce_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  // initial_value binds like a LAMBDA argument, so a reference seed stays a
  // reference for the first call; an error propagates.
  const parser::AstNode* acc_ast = nullptr;
  Value acc = eval_binding_source(call.as_call_arg(0), arena, registry, ctx, &acc_ast);
  if (acc.is_error()) {
    return acc;
  }
  Value err = Value::blank();
  const ArrayValue* in = eval_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (in == nullptr) {
    return err;
  }
  const LambdaValue* lv = resolve_callable(call.as_call_arg(2), /*call_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }
  const SourceOrigin origin = source_origin(call.as_call_arg(1), in, arena, registry, ctx);

  // An empty input is a no-op fold: REDUCE returns the seed unchanged
  // (the Mac Excel observed behaviour for `=REDUCE(0, FILTER(...empty...), ...)`).
  const std::size_t total = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  if (total == 0U) {
    return acc;
  }

  Value args[2] = {Value::blank(), Value::blank()};
  const parser::AstNode* arg_asts[2] = {nullptr, nullptr};
  for (std::size_t i = 0; i < total; ++i) {
    args[0] = acc;
    args[1] = in->cells[i];
    // Only the seed can be a reference; every later accumulator is a result.
    arg_asts[0] = i == 0U ? acc_ast : nullptr;
    arg_asts[1] = element_ast(origin, in, i, arena);
    // An errored cell reaches the lambda verbatim, matching SCAN: a body
    // that guards with IFERROR keeps folding, and one that does not leaves
    // the error in the accumulator, which is what the fold returns.
    acc = invoke_lambda_values_with_ast(lv, 2U, args, arg_asts, arena, registry, ctx);
  }
  return acc;
}

Value eval_scan_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 3U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode* acc_ast = nullptr;
  Value acc = eval_binding_source(call.as_call_arg(0), arena, registry, ctx, &acc_ast);
  if (acc.is_error()) {
    return acc;
  }
  Value err = Value::blank();
  const ArrayValue* in = eval_array_arg(call.as_call_arg(1), arena, registry, ctx, &err);
  if (in == nullptr) {
    return err;
  }
  const LambdaValue* lv = resolve_callable(call.as_call_arg(2), /*call_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
  }
  const SourceOrigin origin = source_origin(call.as_call_arg(1), in, arena, registry, ctx);

  const std::uint32_t rows = in->rows;
  const std::uint32_t cols = in->cols;
  if (rows == 0U || cols == 0U) {
    // SCAN must emit a shape-preserving result; an empty input has no
    // shape to spill into, so #CALC! mirrors BYROW / BYCOL.
    return Value::error(ErrorCode::Calc);
  }

  Value* out_cells = nullptr;
  ArrayValue* out_arr = alloc_cells(arena, rows, cols, out_cells);
  if (out_arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }

  Value args[2] = {Value::blank(), Value::blank()};
  const parser::AstNode* arg_asts[2] = {nullptr, nullptr};
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < total; ++i) {
    args[0] = acc;
    args[1] = in->cells[i];
    arg_asts[0] = i == 0U ? acc_ast : nullptr;
    arg_asts[1] = element_ast(origin, in, i, arena);
    // An errored cell is handed to the lambda verbatim rather than
    // short-circuited: a body that guards with IFERROR recovers, and one
    // that does not returns the error, which then rides the accumulator
    // into every later cell. Both outcomes are per-cell, not whole-call.
    const Value res = invoke_lambda_values_with_ast(lv, 2U, args, arg_asts, arena, registry, ctx);
    if (res.is_error()) {
      acc = res;
      out_cells[i] = res;
      continue;
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

  return Value::array(out_arr);
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
  const LambdaValue* lv = resolve_callable(call.as_call_arg(2), /*call_arity=*/2U, arena, registry, ctx, &err);
  if (lv == nullptr) {
    return err;
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

  return Value::array(out_arr);
}

}  // namespace eval
}  // namespace formulon
