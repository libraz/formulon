//
// Implementation of the OFFSET lazy impl plus the related range-shape
// expanders (`expand_offset_call`, `expand_choose_call`, `expand_if_call`,
// `expand_row_call`, `expand_column_call`).
//
// The rectangle-construction core (`compute_offset_rect`, `OffsetBase`)
// lives in `reference/common.cpp` because the intersection resolver
// (`reference/intersection.cpp`) reaches it too. This TU only owns the
// OFFSET evaluator and the range-expander dispatch on top.

#include "eval/reference/offset.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/lookups/classic.h"
#include "eval/name_env_resolve.h"
#include "eval/range_args.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/reference/common.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// The rectangle an OFFSET call names: the base's sheet plus the 0-based
// top-left corner and the extent.
struct OffsetRect {
  refs_internal::OffsetBase base{};
  std::uint32_t top_row = 0;
  std::uint32_t left_col = 0;
  std::uint32_t height = 0;
  std::uint32_t width = 0;
};

bool resolve_offset_rect(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetRect* out, ErrorCode* out_err) {
  return refs_internal::compute_offset_rect(call, arena, registry, ctx, &out->base, &out->top_row, &out->left_col,
                                            &out->height, &out->width, out_err);
}

// Reads every cell of `rect` through `EvalContext::expand_range`, which
// already handles cross-sheet routing, cycle detection, and per-cell
// recursion.
Expected<std::vector<Value>, ErrorCode> expand_offset_rect(const OffsetRect& rect, Arena& arena,
                                                           const FunctionRegistry& registry, const EvalContext& ctx) {
  parser::Reference lhs{};
  parser::Reference rhs{};
  lhs.sheet = rect.base.sheet;
  lhs.row = rect.top_row;
  lhs.col = rect.left_col;
  rhs.sheet = rect.base.sheet;
  rhs.row = rect.top_row + rect.height - 1U;
  rhs.col = rect.left_col + rect.width - 1U;
  return ctx.expand_range(lhs, rhs, arena, registry);
}

// Resolves `node` through `resolve_range_arg` into the expander out-params.
bool expand_resolved_range(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                           std::uint32_t* out_rows, std::uint32_t* out_cols) {
  auto resolved = resolve_range_arg(node, arena, registry, ctx);
  if (!resolved) {
    *out_err_code = resolved.error();
    return false;
  }
  auto& rr = resolved.value();
  if (out_rows != nullptr) {
    *out_rows = rr.rows;
  }
  if (out_cols != nullptr) {
    *out_cols = rr.cols;
  }
  *out_cells = std::move(rr.cells);
  return true;
}

}  // namespace

Value eval_offset_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  OffsetRect rect{};
  ErrorCode err = ErrorCode::Value;
  if (!resolve_offset_rect(call, arena, registry, ctx, &rect, &err)) {
    return Value::error(err);
  }
  const std::uint32_t height = rect.height;
  const std::uint32_t width = rect.width;
  if (height == 1U && width == 1U) {
    parser::Reference target{};
    target.sheet = rect.base.sheet;
    target.row = rect.top_row;
    target.col = rect.left_col;
    return ctx.resolve_ref(target, arena, registry);
  }

  // A direct multi-cell OFFSET is a dynamic array. Materialise the whole
  // rectangle so it can spill at the formula cell; range-aware consumers
  // use the matching `expand_offset_call` path below.
  auto expanded = expand_offset_rect(rect, arena, registry, ctx);
  if (!expanded) {
    return Value::error(expanded.error());
  }
  const std::vector<Value>& cells = expanded.value();
  ArrayValue* out = array_from_values(height, width, cells.data(), cells.size(), arena);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  return Value::array(out);
}

bool expand_offset_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  OffsetRect rect{};
  ErrorCode err = ErrorCode::Value;
  if (!resolve_offset_rect(call, arena, registry, ctx, &rect, &err)) {
    *out_err_code = err;
    return false;
  }
  auto expanded = expand_offset_rect(rect, arena, registry, ctx);
  if (!expanded) {
    *out_err_code = expanded.error();
    return false;
  }
  *out_cells = std::move(expanded.value());
  if (out_rows != nullptr) {
    *out_rows = rect.height;
  }
  if (out_cols != nullptr) {
    *out_cols = rect.width;
  }
  return true;
}

bool expand_indirect_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                          std::uint32_t* out_rows, std::uint32_t* out_cols) {
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  bool is_range = false;
  ErrorCode err = ErrorCode::Ref;
  if (!resolve_reference_call(call, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, &is_range, &err)) {
    *out_err_code = err;
    return false;
  }
  parser::Reference lhs{};
  parser::Reference rhs{};
  lhs.sheet = sheet;
  lhs.row = top;
  lhs.col = left;
  rhs.sheet = sheet;
  rhs.row = bottom;
  rhs.col = right;
  auto expanded = ctx.expand_range(lhs, rhs, arena, registry);
  if (!expanded) {
    *out_err_code = expanded.error();
    return false;
  }
  *out_cells = std::move(expanded.value());
  if (out_rows != nullptr) {
    *out_rows = bottom - top + 1U;
  }
  if (out_cols != nullptr) {
    *out_cols = right - left + 1U;
  }
  return true;
}

bool expand_choose_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  const std::uint32_t arity = call.as_call_arity();
  // Need at least the index plus one value, matching `eval_choose_lazy`.
  if (arity < 2U) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  // Evaluate the index argument; CHOOSE expects a 1-based integer
  // selector. Errors propagate with their original code, mirroring
  // `eval_choose_lazy`.
  const Value idx_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (idx_val.is_error()) {
    *out_err_code = idx_val.as_error();
    return false;
  }
  if (idx_val.is_array()) {
    const Value result = eval_choose_array_index_lazy(call, idx_val, arena, registry, ctx);
    return expand_array_result(result, out_cells, out_err_code, out_rows, out_cols);
  }
  auto idx_num = coerce_to_number(idx_val);
  if (!idx_num) {
    *out_err_code = idx_num.error();
    return false;
  }
  // Excel truncates (toward zero) rather than rounds: CHOOSE(2.9, ...)
  // selects the 2nd value, not the 3rd. For valid `[1, arity-1]` indices
  // these are non-negative, so `std::floor` matches `eval_choose_lazy`.
  const double raw = std::floor(idx_num.value());
  if (!(raw >= 1.0 && raw <= static_cast<double>(arity - 1U))) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  const auto picked_slot = static_cast<std::uint32_t>(raw);
  // `resolve_range_arg` routes a nested OFFSET / CHOOSE / IF back through
  // these expanders, so chains flatten cleanly.
  return expand_resolved_range(call.as_call_arg(picked_slot), arena, registry, ctx, out_cells, out_err_code, out_rows,
                               out_cols);
}

bool expand_array_result(const Value& result, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                         std::uint32_t* out_rows, std::uint32_t* out_cols) {
  if (result.is_error()) {
    *out_err_code = result.as_error();
    return false;
  }
  if (!result.is_array()) {
    *out_err_code = ErrorCode::Value;
    return false;
  }
  const ArrayValue* array = result.as_array();
  const std::size_t count = static_cast<std::size_t>(array->rows) * array->cols;
  out_cells->assign(array->cells, array->cells + count);
  if (out_rows != nullptr) {
    *out_rows = array->rows;
  }
  if (out_cols != nullptr) {
    *out_cols = array->cols;
  }
  return true;
}

bool expand_if_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    std::vector<Value>* out_cells, ErrorCode* out_err_code, std::uint32_t* out_rows,
                    std::uint32_t* out_cols) {
  // `resolve_range_arg` owns the reference-shaped IF rule: it short-circuits
  // the condition, broadcasts an array condition, and recurses into the
  // picked branch.
  return expand_resolved_range(call, arena, registry, ctx, out_cells, out_err_code, out_rows, out_cols);
}

namespace {

// Shared body of `expand_row_call` / `expand_column_call`. `want_row`
// picks the axis: when true, fills `out_cells` with 1-based row indices
// drawn from `[top..bottom]` of the resolved rectangle and reports
// `(rows = bottom-top+1, cols = 1)`; when false, fills with column
// indices from `[left..right]` and reports `(rows = 1, cols = right-left+1)`.
//
// The shape inspection mirrors `eval_row_or_column` in `shape_ops_lazy.cpp`:
// LET-bound NameRefs are looked through, single-cell `Ref` becomes a 1x1,
// `RangeOp(Ref, Ref)` covers the full row / column span, and a nested
// reference-returning `Call` (INDIRECT / OFFSET / IF / CHOOSE) routes
// through `resolve_reference_call`. Anything else evaluates the subtree
// to surface errors and otherwise reports `#VALUE!`. The 0-arity branch
// (bare `=ROW()` / `=COLUMN()`) emits a single 1x1 indexed by the
// formula cell, or `#VALUE!` if no formula cell is bound.
bool expand_row_or_column_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, bool want_row, std::vector<Value>* out_cells,
                               ErrorCode* out_err_code, std::uint32_t* out_rows, std::uint32_t* out_cols) {
  const std::uint32_t arity = call.as_call_arity();
  out_cells->clear();
  if (arity == 0U) {
    if (!ctx.has_formula_cell()) {
      *out_err_code = ErrorCode::Value;
      return false;
    }
    const std::uint32_t idx = want_row ? ctx.formula_row() : ctx.formula_col();
    out_cells->push_back(Value::number(static_cast<double>(idx + 1U)));
    if (out_rows != nullptr) {
      *out_rows = 1U;
    }
    if (out_cols != nullptr) {
      *out_cols = 1U;
    }
    return true;
  }
  if (arity != 1U) {
    *out_err_code = ErrorCode::Value;
    return false;
  }

  // LET-binding passthrough: `=LET(r, A1:A3, SUM(ROW(r)))` parses `r` as
  // a NameRef. Mirror `eval_row_or_column`'s rule: accept the broader
  // "Ref OR range-shaped" set so a single-cell binding still yields a
  // meaningful row / column index.
  const parser::AstNode& raw_arg = call.as_call_arg(0);
  const parser::AstNode& arg = resolve_range_binding(raw_arg, ctx.name_env(), /*accept_ref=*/true);

  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  bool resolved_rect = false;

  const parser::NodeKind k = arg.kind();
  parser::Reference rect_lhs{};
  parser::Reference rect_rhs{};
  if (declared_rect_endpoint_pair(arg, &rect_lhs, &rect_rhs)) {
    // Same derivation the scalar seam in `shape_ops_lazy.cpp` uses: a
    // full-axis endpoint names a coordinate only on its bounded axis, so
    // the rectangle cannot come from the endpoints' raw row/col fields.
    const Expected<DeclaredRect, ErrorCode> rect = declared_rect(rect_lhs, rect_rhs);
    if (!rect) {
      *out_err_code = rect.error();
      return false;
    }
    top = rect.value().row_first;
    bottom = rect.value().row_last;
    left = rect.value().col_first;
    right = rect.value().col_last;
    resolved_rect = true;
  } else if (k == parser::NodeKind::RangeOp) {
    // A `RangeOp` whose endpoints are not both bare references names no
    // rectangle until it is evaluated, which this seam does not do.
    *out_err_code = ErrorCode::Value;
    return false;
  } else if (k == parser::NodeKind::Call) {
    std::string_view sheet;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    // Only the position is read, so the reference is no read of its cells.
    if (resolve_reference_call(arg, arena, registry, ctx.without_dynamic_read_callback(), &sheet, &top, &left, &bottom,
                               &right, &is_range, &err)) {
      resolved_rect = true;
    } else {
      // Fall through to the scalar-evaluate branch so subtree errors propagate.
    }
  }

  if (resolved_rect) {
    if (want_row) {
      const std::uint32_t height = bottom - top + 1U;
      out_cells->reserve(height);
      for (std::uint32_t r = top; r <= bottom; ++r) {
        out_cells->push_back(Value::number(static_cast<double>(r + 1U)));
      }
      if (out_rows != nullptr) {
        *out_rows = height;
      }
      if (out_cols != nullptr) {
        *out_cols = 1U;
      }
    } else {
      const std::uint32_t width = right - left + 1U;
      out_cells->reserve(width);
      for (std::uint32_t c = left; c <= right; ++c) {
        out_cells->push_back(Value::number(static_cast<double>(c + 1U)));
      }
      if (out_rows != nullptr) {
        *out_rows = 1U;
      }
      if (out_cols != nullptr) {
        *out_cols = width;
      }
    }
    return true;
  }

  // Evaluate the subtree to surface any errors verbatim (e.g. `ROW(1/0)`),
  // otherwise reject non-references with `#VALUE!`.
  const Value v = eval_node(arg, arena, registry, ctx);
  if (v.is_error()) {
    *out_err_code = v.as_error();
    return false;
  }
  *out_err_code = ErrorCode::Value;
  return false;
}

}  // namespace

bool expand_row_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                     std::uint32_t* out_rows, std::uint32_t* out_cols) {
  return expand_row_or_column_call(call, arena, registry, ctx, /*want_row=*/true, out_cells, out_err_code, out_rows,
                                   out_cols);
}

bool expand_column_call(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx, std::vector<Value>* out_cells, ErrorCode* out_err_code,
                        std::uint32_t* out_rows, std::uint32_t* out_cols) {
  return expand_row_or_column_call(call, arena, registry, ctx, /*want_row=*/false, out_cells, out_err_code, out_rows,
                                   out_cols);
}

}  // namespace eval
}  // namespace formulon
