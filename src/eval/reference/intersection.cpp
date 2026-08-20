//
// Implementation of the rectangle-resolution routines
// (`resolve_reference_call`, `resolve_range_endpoint`,
// `compute_intersect_rect`) that the intersect operator, ROWS/COLUMNS,
// CELL("address",...), AREAS, and spill-anchor recovery all share. These
// routines walk a reference-shaped AST (`Ref`, `RangeOp`, a
// reference-returning `Call`) into a bare rectangle WITHOUT
// dereferencing the cells; the contract lives in
// `eval/range_resolvers.h`.

#include "eval/reference/intersection.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/reference/common.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

bool resolve_reference_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            bool* out_is_range, ErrorCode* out_err) {
  // Callers are expected to handle Ref / RangeOp / ArrayLiteral before
  // falling through here; anything that isn't a call is out of scope.
  if (node.kind() != parser::NodeKind::Call) {
    *out_err = ErrorCode::Value;
    return false;
  }
  const std::string_view name = node.as_call_name();
  if (strings::case_insensitive_eq(name, "INDIRECT")) {
    refs_internal::IndirectReference indirect{};
    if (!refs_internal::resolve_indirect_reference(node, arena, registry, ctx, &indirect, out_err)) {
      return false;
    }
    // A sheet qualifier is only valid if the workbook actually holds a
    // matching sheet. Without this check a caller like `ROW(INDIRECT(
    // "NonExistent!A1"))` would happily report row 1 for a sheet that
    // doesn't exist; Excel surfaces `#REF!` in that case.
    if (!indirect.sheet.empty()) {
      const Workbook* wb = ctx.workbook();
      if (wb == nullptr || wb->sheet_by_name(indirect.sheet) == nullptr) {
        *out_err = ErrorCode::Ref;
        return false;
      }
    }
    *out_sheet = indirect.sheet;
    *out_top_row = indirect.top_row;
    *out_left_col = indirect.left_col;
    *out_bottom_row = indirect.bottom_row;
    *out_right_col = indirect.right_col;
    *out_is_range = indirect.is_range;
    return true;
  }
  if (strings::case_insensitive_eq(name, "OFFSET")) {
    refs_internal::OffsetBase base{};
    std::uint32_t top_row = 0;
    std::uint32_t left_col = 0;
    std::uint32_t height = 0;
    std::uint32_t width = 0;
    ErrorCode err = ErrorCode::Value;
    if (!refs_internal::compute_offset_rect(node, arena, registry, ctx, &base, &top_row, &left_col, &height, &width,
                                            &err)) {
      *out_err = err;
      return false;
    }
    *out_sheet = base.sheet;
    *out_top_row = top_row;
    *out_left_col = left_col;
    *out_bottom_row = top_row + height - 1U;
    *out_right_col = left_col + width - 1U;
    *out_is_range = (height > 1U) || (width > 1U);
    return true;
  }
  if (strings::case_insensitive_eq(name, "IF")) {
    // `IF(cond, then, [else])` preserves reference-shape: when both
    // branches are range references Excel routes the picked branch
    // through verbatim, so `ROWS(IF(TRUE, A1:B3, A1:B3))` reports 3
    // rather than degrading to the scalar-fallback 1x1. We short-circuit
    // on `cond` exactly like `eval_if_lazy`, then resolve the chosen
    // branch as a range endpoint (which handles Ref / RangeOp / nested
    // INDIRECT / OFFSET / CHOOSE / IF transparently).
    const std::uint32_t arity = node.as_call_arity();
    if (arity != 2U && arity != 3U) {
      *out_err = ErrorCode::Value;
      return false;
    }
    const Value cond = eval_node(node.as_call_arg(0), arena, registry, ctx);
    if (cond.is_error()) {
      *out_err = cond.as_error();
      return false;
    }
    auto coerced = coerce_to_bool(cond);
    if (!coerced) {
      *out_err = coerced.error();
      return false;
    }
    const std::uint32_t pick = coerced.value() ? 1U : (arity == 3U ? 2U : 1U);
    if (!coerced.value() && arity == 2U) {
      // `IF(FALSE, then)` returns boolean FALSE in Excel's scalar path,
      // which is not a reference. Surface `#VALUE!` so the caller falls
      // back to the scalar / non-reference branch.
      *out_err = ErrorCode::Value;
      return false;
    }
    const parser::AstNode& picked = node.as_call_arg(pick);
    if (!resolve_range_endpoint(picked, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                out_right_col, out_err)) {
      return false;
    }
    *out_is_range = (*out_top_row != *out_bottom_row) || (*out_left_col != *out_right_col);
    return true;
  }
  if (strings::case_insensitive_eq(name, "CHOOSE")) {
    const std::uint32_t arity = node.as_call_arity();
    if (arity < 2U) {
      *out_err = ErrorCode::Value;
      return false;
    }
    // Evaluate the index argument; CHOOSE expects a 1-based integer
    // selector. Anything that fails coercion (text, blank-as-strict,
    // error) propagates with its original code.
    const Value idx_val = eval_node(node.as_call_arg(0), arena, registry, ctx);
    if (idx_val.is_error()) {
      *out_err = idx_val.as_error();
      return false;
    }
    auto idx_int = refs_internal::read_int(idx_val);
    if (!idx_int) {
      *out_err = idx_int.error();
      return false;
    }
    const int idx = idx_int.value();
    const std::uint32_t n_choices = arity - 1U;
    if (idx < 1 || static_cast<std::uint32_t>(idx) > n_choices) {
      // Excel: out-of-range index -> #VALUE!
      *out_err = ErrorCode::Value;
      return false;
    }
    // The picked choice is at slot `idx` (0 = index, 1..n = choices).
    // Recurse via `resolve_range_endpoint` so plain Ref / nested
    // OFFSET-INDIRECT-CHOOSE endpoints all reduce to a rectangle.
    const parser::AstNode& picked = node.as_call_arg(static_cast<std::uint32_t>(idx));
    if (!resolve_range_endpoint(picked, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                out_right_col, out_err)) {
      return false;
    }
    *out_is_range = (*out_top_row != *out_bottom_row) || (*out_left_col != *out_right_col);
    return true;
  }
  // Any other call name is not a reference-returning builtin we know
  // how to handle here.
  *out_err = ErrorCode::Value;
  return false;
}

namespace {

// Unions the two endpoint rectangles of a `RangeOp` node, which is what
// `a:b` denotes for endpoint composition and for an intersect operand
// alike. Defined below because it recurses through
// `resolve_range_endpoint`.
bool union_range_endpoints(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                           std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                           ErrorCode* out_err);

}  // namespace

bool resolve_range_endpoint(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                            std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                            ErrorCode* out_err) {
  if (node.kind() == parser::NodeKind::Ref) {
    const parser::Reference& r = node.as_ref();
    // Whole-column / whole-row refs cannot anchor an endpoint composition
    // because the union rectangle would be unbounded; surface #VALUE! to
    // match `expand_range`'s existing degradation for these shapes.
    if (r.is_full_col || r.is_full_row) {
      *out_err = ErrorCode::Value;
      return false;
    }
    *out_sheet = r.sheet;
    *out_top_row = r.row;
    *out_left_col = r.col;
    *out_bottom_row = r.row;
    *out_right_col = r.col;
    return true;
  }
  if (node.kind() == parser::NodeKind::RangeOp) {
    // CHOOSE / IF can pick a range-shaped child (e.g. `CHOOSE(1, A1:B2,
    // A1:B3)` or `IF(TRUE, A1:B3, A1:B3)`). Resolve each side as an
    // endpoint and union the rectangles so callers see the full picked
    // range, matching Excel's `ROWS(IF(...))` / `COLUMNS(CHOOSE(...))`
    // semantics. Plain intersect-operand callers don't reach this branch
    // because they split RangeOp themselves before recursing.
    return union_range_endpoints(node, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                 out_right_col, out_err);
  }
  if (node.kind() == parser::NodeKind::Call) {
    bool is_range_unused = false;
    return resolve_reference_call(node, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                  out_right_col, &is_range_unused, out_err);
  }
  // Anything else (NameRef, BinaryOp, ArrayLiteral, etc.) is
  // not a recognized range endpoint shape.
  *out_err = ErrorCode::Ref;
  return false;
}

namespace {

bool union_range_endpoints(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                           std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                           ErrorCode* out_err) {
  std::string_view lhs_sheet;
  std::string_view rhs_sheet;
  std::uint32_t lhs_top = 0;
  std::uint32_t lhs_left = 0;
  std::uint32_t lhs_bottom = 0;
  std::uint32_t lhs_right = 0;
  std::uint32_t rhs_top = 0;
  std::uint32_t rhs_left = 0;
  std::uint32_t rhs_bottom = 0;
  std::uint32_t rhs_right = 0;
  if (!resolve_range_endpoint(node.as_range_lhs(), arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom,
                              &lhs_right, out_err) ||
      !resolve_range_endpoint(node.as_range_rhs(), arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom,
                              &rhs_right, out_err)) {
    return false;
  }
  // Two named sheets must agree; one bare endpoint inherits the other's
  // sheet, and two bare endpoints leave the caller's own sheet in force.
  if (!lhs_sheet.empty() && !rhs_sheet.empty()) {
    if (!sheet_names::equal(lhs_sheet, rhs_sheet)) {
      *out_err = ErrorCode::Ref;
      return false;
    }
    *out_sheet = lhs_sheet;
  } else if (!lhs_sheet.empty()) {
    *out_sheet = lhs_sheet;
  } else {
    *out_sheet = rhs_sheet;
  }
  *out_top_row = std::min(lhs_top, rhs_top);
  *out_left_col = std::min(lhs_left, rhs_left);
  *out_bottom_row = std::max(lhs_bottom, rhs_bottom);
  *out_right_col = std::max(lhs_right, rhs_right);
  return true;
}

// Resolves an `IntersectOp` operand AST into a rectangle. Accepts the
// same shapes the `:` operator produces: a `RangeOp` over two
// `resolve_range_endpoint`-compatible endpoints, a single `Ref`, or a
// reference-returning `Call`. Whole-column / whole-row inputs surface
// `#VALUE!`; mismatched cross-sheet endpoints surface `#REF!`. Returns
// true on success and writes the inclusive 0-based rectangle.
bool resolve_intersect_operand(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, std::string_view* out_sheet, std::uint32_t* out_top_row,
                               std::uint32_t* out_left_col, std::uint32_t* out_bottom_row, std::uint32_t* out_right_col,
                               ErrorCode* out_err) {
  if (node.kind() == parser::NodeKind::IntersectOp) {
    // Left-associative chain: `(lhs rhs)` is itself an intersection
    // rectangle. A disjoint inner intersection has no overlapping cells,
    // so an enclosing intersection is necessarily empty too -> #NULL!.
    bool inner_disjoint = false;
    if (!compute_intersect_rect(node.as_intersect_lhs(), node.as_intersect_rhs(), arena, registry, ctx, out_sheet,
                                out_top_row, out_left_col, out_bottom_row, out_right_col, &inner_disjoint, out_err)) {
      return false;
    }
    if (inner_disjoint) {
      *out_err = ErrorCode::Null;
      return false;
    }
    return true;
  }
  // A full-axis endpoint carries a meaningful coordinate only on its bounded
  // axis, so its rectangle comes from the shared derivation rather than from
  // the endpoints' raw row/col fields. Both spellings reach here: the single
  // `Ref` the parser folds `A:A` / `1:1` into, and the `RangeOp` over two
  // same-axis whole `Ref`s that `A:C` produces. The latter is a valid
  // intersect operand in Excel (`A:C B:B`) that the endpoint-union path below
  // cannot express, because a full-axis endpoint has no bounded corner to
  // union. Every other shape — a bounded pair, or one the derivation names no
  // rectangle for — falls through to its existing resolution.
  const parser::Reference* rect_lhs = nullptr;
  const parser::Reference* rect_rhs = nullptr;
  if (declared_rect_endpoints(node, &rect_lhs, &rect_rhs)) {
    const Expected<DeclaredRect, ErrorCode> rect = declared_rect(*rect_lhs, *rect_rhs);
    if (rect && rect.value().whole_axis) {
      // The parser keeps the sheet qualifier on the left endpoint.
      *out_sheet = rect_lhs->sheet;
      *out_top_row = rect.value().row_first;
      *out_left_col = rect.value().col_first;
      *out_bottom_row = rect.value().row_last;
      *out_right_col = rect.value().col_last;
      return true;
    }
  }
  if (node.kind() == parser::NodeKind::RangeOp) {
    return union_range_endpoints(node, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                 out_right_col, out_err);
  }
  // Single `Ref` or reference-returning `Call` -> 1x1 rectangle (or the
  // synthesized rectangle from OFFSET / INDIRECT).
  return resolve_range_endpoint(node, arena, registry, ctx, out_sheet, out_top_row, out_left_col, out_bottom_row,
                                out_right_col, out_err);
}

}  // namespace

bool compute_intersect_rect(const parser::AstNode& lhs, const parser::AstNode& rhs, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx, std::string_view* out_sheet,
                            std::uint32_t* out_top_row, std::uint32_t* out_left_col, std::uint32_t* out_bottom_row,
                            std::uint32_t* out_right_col, bool* out_disjoint, ErrorCode* out_err) {
  *out_disjoint = false;
  std::string_view lhs_sheet;
  std::string_view rhs_sheet;
  std::uint32_t lhs_top = 0;
  std::uint32_t lhs_left = 0;
  std::uint32_t lhs_bottom = 0;
  std::uint32_t lhs_right = 0;
  std::uint32_t rhs_top = 0;
  std::uint32_t rhs_left = 0;
  std::uint32_t rhs_bottom = 0;
  std::uint32_t rhs_right = 0;
  if (!resolve_intersect_operand(lhs, arena, registry, ctx, &lhs_sheet, &lhs_top, &lhs_left, &lhs_bottom, &lhs_right,
                                 out_err)) {
    return false;
  }
  if (!resolve_intersect_operand(rhs, arena, registry, ctx, &rhs_sheet, &rhs_top, &rhs_left, &rhs_bottom, &rhs_right,
                                 out_err)) {
    return false;
  }
  if (!lhs_sheet.empty() && !rhs_sheet.empty()) {
    if (!sheet_names::equal(lhs_sheet, rhs_sheet)) {
      *out_err = ErrorCode::Ref;
      return false;
    }
    *out_sheet = lhs_sheet;
  } else if (!lhs_sheet.empty()) {
    *out_sheet = lhs_sheet;
  } else {
    *out_sheet = rhs_sheet;
  }
  const std::uint32_t r1 = std::max(lhs_top, rhs_top);
  const std::uint32_t r2 = std::min(lhs_bottom, rhs_bottom);
  const std::uint32_t c1 = std::max(lhs_left, rhs_left);
  const std::uint32_t c2 = std::min(lhs_right, rhs_right);
  if (r1 > r2 || c1 > c2) {
    *out_disjoint = true;
    return true;
  }
  *out_top_row = r1;
  *out_left_col = c1;
  *out_bottom_row = r2;
  *out_right_col = c2;
  return true;
}

}  // namespace eval
}  // namespace formulon
