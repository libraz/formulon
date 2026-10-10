//
// Shared bodies for the reference-family lazy impls. Hosts:
//
//   * Small helper `apply_offset` that the OFFSET and CHOOSE expanders share
//     with the intersection resolver.
//   * The rectangle-construction routines `resolve_indirect_reference`,
//     `resolve_offset_base`, and `compute_offset_rect`. These touch both
//     the INDIRECT and OFFSET pipelines (and are reached from the
//     intersection resolver too), so they live here to avoid a circular
//     dependency between the sibling TUs.

#include "eval/reference/common.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/a1_parse.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace refs_internal {

bool apply_offset(std::uint32_t base, int offset, std::uint32_t max, std::uint32_t* out) {
  const long long sum = static_cast<long long>(base) + static_cast<long long>(offset);
  if (sum < 0 || sum >= static_cast<long long>(max)) {
    return false;
  }
  *out = static_cast<std::uint32_t>(sum);
  return true;
}

bool resolve_indirect_reference(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx, IndirectReference* out, ErrorCode* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    *out_err = ErrorCode::Value;
    return false;
  }

  // Evaluate `ref_text` first so errors propagate per the dispatcher's
  // left-most-wins rule.
  const Value ref_val = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (ref_val.is_error()) {
    *out_err = ref_val.as_error();
    return false;
  }
  auto text_exp = coerce_to_text(ref_val);
  if (!text_exp) {
    *out_err = text_exp.error();
    return false;
  }

  bool a1_style = true;
  if (arity == 2U) {
    const Value a1_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
    if (a1_val.is_error()) {
      *out_err = a1_val.as_error();
      return false;
    }
    auto b = coerce_to_bool(a1_val);
    if (!b) {
      *out_err = b.error();
      return false;
    }
    a1_style = b.value();
  }
  const std::string& src = text_exp.value();
  if (src.empty()) {
    *out_err = ErrorCode::Ref;
    return false;
  }
  // The flag picks the grammar outright: A1 text under `a1 = FALSE` is as
  // invalid as R1C1 text under `a1 = TRUE`, and each parser rejects the
  // other's spelling on its own. Relative axes resolve against the cell
  // the formula sits in, which is the only reading `R[1]C[1]` has.
  R1C1Base base;
  if (ctx.has_formula_cell()) {
    base.present = true;
    base.row = ctx.formula_row();
    base.col = ctx.formula_col();
  }
  const A1Parse parsed = a1_style ? parse_a1_ref(src) : parse_r1c1_ref(src, base);
  if (!parsed.valid) {
    *out_err = ErrorCode::Ref;
    return false;
  }

  out->sheet = parsed.sheet.empty() ? std::string_view{} : arena.intern(parsed.sheet);
  out->range_syntax = parsed.is_range;
  out->is_full_col = parsed.is_full_col;
  out->is_full_row = parsed.is_full_row;
  if (parsed.is_range) {
    out->top_row = std::min(parsed.row, parsed.row2);
    out->left_col = std::min(parsed.col, parsed.col2);
    out->bottom_row = std::max(parsed.row, parsed.row2);
    out->right_col = std::max(parsed.col, parsed.col2);
    out->is_range = (out->top_row != out->bottom_row) || (out->left_col != out->right_col);
  } else {
    out->top_row = parsed.row;
    out->left_col = parsed.col;
    out->bottom_row = parsed.row;
    out->right_col = parsed.col;
    out->is_range = false;
  }
  ctx.note_dynamic_read(out->sheet, DeclaredRect{out->top_row, out->bottom_row, out->left_col, out->right_col,
                                                 /*whole_axis=*/out->is_full_col || out->is_full_row});
  return true;
}

bool resolve_offset_base(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetBase* out, ErrorCode* out_err) {
  // LET-bound NameRef lookthrough: `LET(r, A1:A3, OFFSET(r, 1, 0))` must see
  // the underlying RangeOp, not the opaque NameRef the scalar path would
  // collapse to its spill anchor. Every sibling reference/shape function
  // (ROW, COLUMN, AREAS, the conditional aggregators, PERCENTOF) already
  // applies this lookthrough; OFFSET's base resolver had not.
  const parser::AstNode& arg = resolve_name_ast(raw_arg, ctx.name_env());
  const parser::NodeKind k = arg.kind();
  if (k == parser::NodeKind::Ref) {
    const parser::Reference& r = arg.as_ref();
    out->sheet = r.sheet;
    if (r.is_full_col) {
      // `A:A`: declared shape spans the full row axis, matching Excel's
      // OFFSET semantics (no used-range clamp -- OFFSET is pure reference
      // arithmetic, not a value walk).
      out->row = 0U;
      out->col = r.col;
      out->rows = Sheet::kMaxRows;
      out->cols = 1U;
    } else if (r.is_full_row) {
      out->row = r.row;
      out->col = 0U;
      out->rows = 1U;
      out->cols = Sheet::kMaxCols;
    } else {
      out->row = r.row;
      out->col = r.col;
      out->rows = 1U;
      out->cols = 1U;
    }
    return true;
  }
  if (k == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg.as_range_lhs();
    const parser::AstNode& rhs_ast = arg.as_range_rhs();
    if (lhs_ast.kind() != parser::NodeKind::Ref || rhs_ast.kind() != parser::NodeKind::Ref) {
      *out_err = ErrorCode::Ref;
      return false;
    }
    const parser::Reference& lhs = lhs_ast.as_ref();
    const parser::Reference& rhs = rhs_ast.as_ref();
    // Multi-column (`A:C`) / multi-row (`1:3`) whole references: the
    // parser produces a RangeOp over two whole-column / whole-row Refs.
    // The bounded axis (columns for `A:C`, rows for `1:3`) is the pair's
    // min/max; the open axis spans the full grid, mirroring the
    // single-Ref branch above.
    if (lhs.is_full_col || rhs.is_full_col) {
      if (!lhs.is_full_col || !rhs.is_full_col) {
        *out_err = ErrorCode::Ref;
        return false;
      }
      out->sheet = !lhs.sheet.empty() ? lhs.sheet : rhs.sheet;
      out->row = 0U;
      out->rows = Sheet::kMaxRows;
      out->col = std::min(lhs.col, rhs.col);
      out->cols = std::max(lhs.col, rhs.col) - out->col + 1U;
      return true;
    }
    if (lhs.is_full_row || rhs.is_full_row) {
      if (!lhs.is_full_row || !rhs.is_full_row) {
        *out_err = ErrorCode::Ref;
        return false;
      }
      out->sheet = !lhs.sheet.empty() ? lhs.sheet : rhs.sheet;
      out->col = 0U;
      out->cols = Sheet::kMaxCols;
      out->row = std::min(lhs.row, rhs.row);
      out->rows = std::max(lhs.row, rhs.row) - out->row + 1U;
      return true;
    }
    // The effective sheet qualifier mirrors `expand_range`: whichever
    // endpoint carries it wins, and mismatched qualifiers are `#REF!`.
    if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
      if (!sheet_names::equal(lhs.sheet, rhs.sheet)) {
        *out_err = ErrorCode::Ref;
        return false;
      }
      out->sheet = lhs.sheet;
    } else if (!lhs.sheet.empty()) {
      out->sheet = lhs.sheet;
    } else if (!rhs.sheet.empty()) {
      out->sheet = rhs.sheet;
    }
    const std::uint32_t r_lo = std::min(lhs.row, rhs.row);
    const std::uint32_t r_hi = std::max(lhs.row, rhs.row);
    const std::uint32_t c_lo = std::min(lhs.col, rhs.col);
    const std::uint32_t c_hi = std::max(lhs.col, rhs.col);
    out->row = r_lo;
    out->col = c_lo;
    out->rows = r_hi - r_lo + 1U;
    out->cols = c_hi - c_lo + 1U;
    return true;
  }
  if (k == parser::NodeKind::Call) {
    // Nested INDIRECT / OFFSET as OFFSET's base: resolve to a rectangle
    // without dereferencing, then adopt it as the base shape. Only its
    // position is used, so it is no read of its cells.
    std::string_view sheet;
    std::uint32_t top = 0;
    std::uint32_t left = 0;
    std::uint32_t bottom = 0;
    std::uint32_t right = 0;
    bool is_range = false;
    ErrorCode err = ErrorCode::Value;
    if (!resolve_reference_call(arg, arena, registry, ctx.without_dynamic_read_callback(), &sheet, &top, &left, &bottom,
                                &right, &is_range, &err)) {
      *out_err = err;
      return false;
    }
    out->sheet = sheet;
    out->row = top;
    out->col = left;
    out->rows = bottom - top + 1U;
    out->cols = right - left + 1U;
    return true;
  }
  // Anything else (literal, scalar expr, array literal, or a NameRef
  // whose LET binding -- already looked through above -- has no AST or is
  // scalar-valued) is not a valid reference shape for OFFSET. Excel
  // returns `#VALUE!`.
  *out_err = ErrorCode::Value;
  return false;
}

bool compute_offset_rect(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, OffsetBase* out_base, std::uint32_t* out_top_row,
                         std::uint32_t* out_left_col, std::uint32_t* out_height, std::uint32_t* out_width,
                         ErrorCode* out_err) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3U || arity > 5U) {
    *out_err = ErrorCode::Value;
    return false;
  }
  if (!resolve_offset_base(call.as_call_arg(0), arena, registry, ctx, out_base, out_err)) {
    return false;
  }

  // Evaluate rows / cols and optional height / width in turn. Any error
  // propagates with its original code.
  auto eval_int = [&](std::uint32_t idx, int* out_val) -> bool {
    const Value v = eval_node(call.as_call_arg(idx), arena, registry, ctx);
    if (v.is_error()) {
      *out_err = v.as_error();
      return false;
    }
    auto parsed = coerce_to_truncated_int(v);
    if (!parsed) {
      *out_err = parsed.error();
      return false;
    }
    *out_val = parsed.value();
    return true;
  };

  // Height / width get a slightly different coercion than rows_off /
  // cols_off: Mac Excel 365 truncates the fractional part toward zero,
  // but a non-zero magnitude < 1 (e.g. `0.9`, `-0.5`) is bumped up to
  // ±1 instead of collapsing to 0 (which would otherwise misfire the
  // `height == 0 || width == 0 -> #REF!` guard below). The bump is
  // sign-preserving so that negative-fractional widths still extend in
  // the negative direction.
  auto eval_dim = [&](std::uint32_t idx, int* out_val) -> bool {
    const Value v = eval_node(call.as_call_arg(idx), arena, registry, ctx);
    if (v.is_error()) {
      *out_err = v.as_error();
      return false;
    }
    auto coerced = coerce_to_number(v);
    if (!coerced) {
      *out_err = coerced.error();
      return false;
    }
    const double d = coerced.value();
    int truncated = static_cast<int>(std::trunc(d));
    if (truncated == 0 && d != 0.0) {
      truncated = (d > 0.0) ? 1 : -1;
    }
    *out_val = truncated;
    return true;
  };

  int rows_off = 0;
  int cols_off = 0;
  if (!eval_int(1U, &rows_off) || !eval_int(2U, &cols_off)) {
    return false;
  }

  // An omitted height or width, trailing or empty (`OFFSET(A1,0,0,,2)`),
  // keeps the base reference's.
  auto given = [&](std::uint32_t idx) {
    if (idx >= arity) {
      return false;
    }
    const parser::AstNode& arg = call.as_call_arg(idx);
    return arg.kind() != parser::NodeKind::Literal || !arg.as_literal().is_blank();
  };
  int height_i = static_cast<int>(out_base->rows);
  int width_i = static_cast<int>(out_base->cols);
  if (given(3U) && !eval_dim(3U, &height_i)) {
    return false;
  }
  if (given(4U) && !eval_dim(4U, &width_i)) {
    return false;
  }
  // Zero height or width -> `#REF!`. Excel allows negative height / width
  // meaning the rectangle extends in the negative direction from the
  // anchor (anchor is the bottom-right corner of the rectangle instead
  // of the top-left). We normalise the absolute magnitude here and
  // adjust the anchor position below.
  if (height_i == 0 || width_i == 0) {
    *out_err = ErrorCode::Ref;
    return false;
  }
  const bool neg_height = height_i < 0;
  const bool neg_width = width_i < 0;
  const std::uint32_t abs_height = static_cast<std::uint32_t>(neg_height ? -height_i : height_i);
  const std::uint32_t abs_width = static_cast<std::uint32_t>(neg_width ? -width_i : width_i);

  // Apply the (rows, cols) offset to the base's top-left corner.
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  if (!apply_offset(out_base->row, rows_off, Sheet::kMaxRows, &anchor_row) ||
      !apply_offset(out_base->col, cols_off, Sheet::kMaxCols, &anchor_col)) {
    *out_err = ErrorCode::Ref;
    return false;
  }

  // For negative height / width the anchor is the rectangle's bottom
  // (or right) edge: walk `abs_dim - 1` units back to find the top-left
  // corner. For positive dimensions the anchor is already the top-left.
  long long top_row = static_cast<long long>(anchor_row);
  long long left_col = static_cast<long long>(anchor_col);
  if (neg_height) {
    top_row -= static_cast<long long>(abs_height - 1);
  }
  if (neg_width) {
    left_col -= static_cast<long long>(abs_width - 1);
  }
  const long long bottom_row = top_row + static_cast<long long>(abs_height) - 1;
  const long long right_col = left_col + static_cast<long long>(abs_width) - 1;
  if (top_row < 0 || bottom_row >= static_cast<long long>(Sheet::kMaxRows) || left_col < 0 ||
      right_col >= static_cast<long long>(Sheet::kMaxCols)) {
    *out_err = ErrorCode::Ref;
    return false;
  }

  *out_top_row = static_cast<std::uint32_t>(top_row);
  *out_left_col = static_cast<std::uint32_t>(left_col);
  *out_height = abs_height;
  *out_width = abs_width;
  ctx.note_dynamic_read(out_base->sheet, DeclaredRect{*out_top_row, *out_top_row + abs_height - 1U, *out_left_col,
                                                      *out_left_col + abs_width - 1U, /*whole_axis=*/false});
  return true;
}

}  // namespace refs_internal
}  // namespace eval
}  // namespace formulon
