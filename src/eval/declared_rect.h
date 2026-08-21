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

namespace detail {

/// Running state for `declared_rect_endpoint_pair`'s walk over a `:` chain.
/// Coordinates on an axis a leaf does not bound are folded in anyway and
/// simply never read, matching `declared_rect`'s own rule that a
/// whole-column endpoint's `row` is a structure default.
struct RangeChainFold {
  const parser::Reference* first = nullptr;
  /// The first leaf whose sheet qualifier differs from `first`'s, so the
  /// synthesised pair carries the disagreement into `effective_range_sheet`
  /// rather than restating its cross-sheet rule here.
  const parser::Reference* other_sheet = nullptr;
  /// Extents in the coordinates each leaf actually bounds: a whole-column
  /// leaf spans every row, a whole-row leaf spans every column.
  std::uint32_t row_first = 0;
  std::uint32_t row_last = 0;
  std::uint32_t col_first = 0;
  std::uint32_t col_last = 0;
  bool any_full_col = false;
  bool any_full_row = false;
  bool seen = false;
};

inline bool fold_range_chain(const parser::AstNode& node, RangeChainFold* acc) {
  if (node.kind() == parser::NodeKind::Ref) {
    const parser::Reference& ref = node.as_ref();
    // A leaf's own span, in grid coordinates. A whole-column reference
    // names no row and a whole-row reference names no column, so the
    // unnamed axis contributes the whole grid rather than the structure
    // default sitting in that field.
    const std::uint32_t leaf_row_first = ref.is_full_col ? 0U : ref.row;
    const std::uint32_t leaf_row_last = ref.is_full_col ? Sheet::kMaxRows - 1U : ref.row;
    const std::uint32_t leaf_col_first = ref.is_full_row ? 0U : ref.col;
    const std::uint32_t leaf_col_last = ref.is_full_row ? Sheet::kMaxCols - 1U : ref.col;
    acc->any_full_col = acc->any_full_col || ref.is_full_col;
    acc->any_full_row = acc->any_full_row || ref.is_full_row;
    if (!acc->seen) {
      acc->first = &ref;
      acc->row_first = leaf_row_first;
      acc->row_last = leaf_row_last;
      acc->col_first = leaf_col_first;
      acc->col_last = leaf_col_last;
      acc->seen = true;
      return true;
    }
    if (acc->other_sheet == nullptr && ref.sheet != acc->first->sheet) {
      acc->other_sheet = &ref;
    }
    acc->row_first = std::min(acc->row_first, leaf_row_first);
    acc->row_last = std::max(acc->row_last, leaf_row_last);
    acc->col_first = std::min(acc->col_first, leaf_col_first);
    acc->col_last = std::max(acc->col_last, leaf_col_last);
    return true;
  }
  if (node.kind() != parser::NodeKind::RangeOp) {
    return false;
  }
  return fold_range_chain(node.as_range_lhs(), acc) && fold_range_chain(node.as_range_rhs(), acc);
}

}  // namespace detail

/// Reduces a static reference expression to the endpoint pair that bounds
/// it, following a `:` chain to any depth.
///
/// Excel's `:` yields the bounding box of its two operands and composes, so
/// `A:C:E:G` names columns A through G and `COLUMNS` reports 7 for it. The
/// two-endpoint shapes -- a bare `Ref` and `RangeOp(Ref, Ref)` -- are
/// returned verbatim, so every spelling that resolved before keeps its
/// exact anchors and sheet qualifier and only a genuine chain synthesises
/// corners.
///
/// Returns false when `node` is not a static reference shape at all: a
/// `RangeOp` whose endpoints are reference-returning calls has a rectangle
/// only once evaluated, which is `resolve_range_endpoint`'s job.
///
/// The leaves do not have to bound the same axis. Measured on Excel 365
/// (Mac, ja-JP, 16.112.1): `=COLUMNS(A:C:1:3)` is 16384 and
/// `=ROWS(A:C:1:3)` is 1048576 -- composing a whole-column span with a
/// whole-row span bounds the entire grid, it is not an error. Likewise
/// `=COLUMNS(A:C:E1)` is 5 and `=ROWS(A:C:E1)` is 1048576. The
/// two-endpoint `A:1` earns `#VALUE!` from `declared_rect` and keeps it,
/// but that spelling is a grammar case rather than a semantic one --
/// Excel rejects it when it is typed, so no answer for it is observable.
inline bool declared_rect_endpoint_pair(const parser::AstNode& node, parser::Reference* out_lhs,
                                        parser::Reference* out_rhs) {
  const parser::Reference* pair_lhs = nullptr;
  const parser::Reference* pair_rhs = nullptr;
  if (declared_rect_endpoints(node, &pair_lhs, &pair_rhs)) {
    *out_lhs = *pair_lhs;
    *out_rhs = *pair_rhs;
    return true;
  }
  detail::RangeChainFold acc;
  if (!detail::fold_range_chain(node, &acc) || !acc.seen) {
    return false;
  }
  parser::Reference lhs{};
  parser::Reference rhs{};
  // The synthesised pair is spelled so that `declared_rect` re-derives the
  // rectangle just folded, including whether it spans a whole axis --
  // `whole_axis` is what lets range expansion clamp the unbounded side to
  // the populated extent instead of enumerating the full grid.
  if (acc.any_full_col) {
    // Some leaf spans every row, so the row axis is whole and only the
    // column extent is carried. A whole-row leaf has already widened that
    // extent to the entire grid width.
    lhs.is_full_col = true;
    rhs.is_full_col = true;
    lhs.col = acc.col_first;
    rhs.col = acc.col_last;
  } else if (acc.any_full_row) {
    lhs.is_full_row = true;
    rhs.is_full_row = true;
    lhs.row = acc.row_first;
    rhs.row = acc.row_last;
  } else {
    lhs.row = acc.row_first;
    lhs.col = acc.col_first;
    rhs.row = acc.row_last;
    rhs.col = acc.col_last;
  }
  lhs.sheet = acc.first->sheet;
  lhs.sheet_quoted = acc.first->sheet_quoted;
  rhs.sheet = acc.other_sheet != nullptr ? acc.other_sheet->sheet : acc.first->sheet;
  rhs.sheet_quoted = acc.other_sheet != nullptr ? acc.other_sheet->sheet_quoted : acc.first->sheet_quoted;
  *out_lhs = lhs;
  *out_rhs = rhs;
  return true;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DECLARED_RECT_H_
