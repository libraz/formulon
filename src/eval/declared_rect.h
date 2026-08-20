//
// The rectangle a reference pair *declares*, derived in one place.
//
// Excel treats one rectangle identically however it is spelled, so every
// consumer of a static reference has to agree on what `A:A`, `A:C`, `1:3`
// and `A1:C10` name: the dependency extractor (which cells a formula
// reads), the spilling position in `evaluate()` (what a bare range
// spills), `@` / `SINGLE` (what implicit intersection projects onto), and
// `ROWS` / `COLUMNS` (the shape reported for it). A full-axis endpoint is
// the reason they can disagree — it carries a meaningful coordinate only
// on its bounded axis, and the parser leaves the other one unspecified
// (`parser/reference.h`), so a copy of this derivation that reads
// `Reference::row` for a whole-column endpoint silently answers with a
// structure default.
//
// Nothing here consults a Sheet. Narrowing a whole-axis rectangle to the
// target sheet's populated extent is an enumeration decision that belongs
// to `EvalContext::expand_range` alone and must never be reported as a
// shape; `EvalContext::declared_range_rect` is what shape reporting asks
// instead.

#ifndef FORMULON_EVAL_DECLARED_RECT_H_
#define FORMULON_EVAL_DECLARED_RECT_H_

#include <algorithm>
#include <cstdint>

#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"  // formulon::ErrorCode

namespace formulon {
namespace eval {

/// Inclusive, 0-based rectangle in sheet coordinates.
struct DeclaredRect {
  std::uint32_t row_first = 0;
  std::uint32_t row_last = 0;
  std::uint32_t col_first = 0;
  std::uint32_t col_last = 0;
  /// True when the rectangle spans a whole grid axis because an endpoint
  /// was full-axis, rather than being bounded by two named corners.
  bool whole_axis = false;

  constexpr std::uint32_t rows() const noexcept { return row_last - row_first + 1U; }
  constexpr std::uint32_t cols() const noexcept { return col_last - col_first + 1U; }
  constexpr bool single_cell() const noexcept { return row_first == row_last && col_first == col_last; }
};

/// Derives the rectangle `[lhs : rhs]` declares. A single reference is its
/// own pair (`declared_rect(ref, ref)`), which is how `A:A` and `1:1` — the
/// spellings the parser folds into one `Ref` — are handled.
///
/// Endpoint order is normalised, so `A3:A1` and `A1:A3` yield one rectangle.
/// Two failure modes are distinguished because both already have an
/// established Excel vocabulary at the call sites:
///
///   * `#VALUE!` — the pair names no rectangle. A mixed-axis pair (`A:1`),
///     a lone full-axis endpoint composed with a bounded one, and a
///     reference flagged on both axes are all this case.
///   * `#REF!` — a named coordinate lies outside the grid.
inline Expected<DeclaredRect, ErrorCode> declared_rect(const parser::Reference& lhs, const parser::Reference& rhs) {
  if ((lhs.is_full_col && lhs.is_full_row) || (rhs.is_full_col && rhs.is_full_row)) {
    return ErrorCode::Value;
  }
  const bool full_col = lhs.is_full_col || rhs.is_full_col;
  const bool full_row = lhs.is_full_row || rhs.is_full_row;
  if (full_col && full_row) {
    return ErrorCode::Value;
  }
  if (full_col) {
    // Only same-axis endpoints compose into a rectangle; `A:B3` would have
    // an unbounded corner.
    if (!lhs.is_full_col || !rhs.is_full_col) {
      return ErrorCode::Value;
    }
    if (lhs.col >= Sheet::kMaxCols || rhs.col >= Sheet::kMaxCols) {
      return ErrorCode::Ref;
    }
    // Neither endpoint's `row` is read: a whole-column reference declares
    // the full grid height whatever the parser left in that field.
    return DeclaredRect{0U, Sheet::kMaxRows - 1U, std::min(lhs.col, rhs.col), std::max(lhs.col, rhs.col),
                        /*whole_axis=*/true};
  }
  if (full_row) {
    if (!lhs.is_full_row || !rhs.is_full_row) {
      return ErrorCode::Value;
    }
    if (lhs.row >= Sheet::kMaxRows || rhs.row >= Sheet::kMaxRows) {
      return ErrorCode::Ref;
    }
    return DeclaredRect{std::min(lhs.row, rhs.row), std::max(lhs.row, rhs.row), 0U, Sheet::kMaxCols - 1U,
                        /*whole_axis=*/true};
  }
  if (lhs.row >= Sheet::kMaxRows || rhs.row >= Sheet::kMaxRows || lhs.col >= Sheet::kMaxCols ||
      rhs.col >= Sheet::kMaxCols) {
    return ErrorCode::Ref;
  }
  return DeclaredRect{std::min(lhs.row, rhs.row), std::max(lhs.row, rhs.row), std::min(lhs.col, rhs.col),
                      std::max(lhs.col, rhs.col), /*whole_axis=*/false};
}

/// Recognises the two AST spellings of a declared rectangle — a bare `Ref`
/// (`A1`, `A:A`, `1:1`) and a `RangeOp` over two `Ref`s — and writes the
/// endpoint pair for `declared_rect`. A single `Ref` yields the same
/// pointer twice.
///
/// Returns false for every other node, including a `RangeOp` whose
/// endpoints are reference-returning calls: those have a rectangle only
/// once evaluated, which is `resolve_range_endpoint`'s job, not this one's.
/// The out-params are left untouched in that case.
inline bool declared_rect_endpoints(const parser::AstNode& node, const parser::Reference** out_lhs,
                                    const parser::Reference** out_rhs) {
  if (node.kind() == parser::NodeKind::Ref) {
    *out_lhs = &node.as_ref();
    *out_rhs = *out_lhs;
    return true;
  }
  if (node.kind() != parser::NodeKind::RangeOp) {
    return false;
  }
  const parser::AstNode& lhs = node.as_range_lhs();
  const parser::AstNode& rhs = node.as_range_rhs();
  if (lhs.kind() != parser::NodeKind::Ref || rhs.kind() != parser::NodeKind::Ref) {
    return false;
  }
  *out_lhs = &lhs.as_ref();
  *out_rhs = &rhs.as_ref();
  return true;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DECLARED_RECT_H_
