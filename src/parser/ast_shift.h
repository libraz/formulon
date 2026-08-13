//
// Reference-rewriting transforms over an `AstNode` tree.
//
// `shift_refs` walks the AST and returns a new tree in which every
// reference (Ref, SpillRef, ExternalRef, range endpoints inside RangeOp)
// has been mapped through the supplied `RefTransform`. Non-reference
// subtrees are forwarded unchanged when no descendant rewrites occur, so
// the cost of an identity walk is bounded by the size of the input.
//
// The transform abstraction supports several distinct operations under one
// walker:
//   - **Relative shift**: applied to defined-name re-anchoring,
//     conditional-formatting per-cell evaluation, and (in a future bundle)
//     row/column insert/delete.
//   - **Sheet rename**: applied to workbook-scoped defined names whose
//     formulas pin a sheet by name (see `parser/ref_transforms.h`).
//
// Out-of-bounds rewrites collapse the affected subtree to the Excel
// `#REF!` literal — the convention every existing call site already relies
// on. Callers that need a different policy can return `std::nullopt` with
// their own override.
//
// The walker is allocation-aware but allocation-failure tolerant: it
// returns `nullptr` only when the arena cannot fulfil a request that the
// transform actually demanded. Identity preservation never allocates.

#ifndef FORMULON_PARSER_AST_SHIFT_H_
#define FORMULON_PARSER_AST_SHIFT_H_

#include <cstdint>
#include <optional>
#include <string_view>
#include <utility>

#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"

namespace formulon {
namespace parser {

/// Generic reference-rewriting policy for `shift_refs`.
///
/// Implementations override `apply` for the common case (sheet-local cell
/// references and the inner cell of an external ref). `apply_external` may
/// be overridden separately when the transform needs to react to the
/// external sheet field (which lives outside `Reference.sheet`); the
/// default implementation forwards to `apply`.
class RefTransform {
 public:
  virtual ~RefTransform() = default;

  /// Returns the rewritten reference. Returning `std::nullopt` collapses
  /// the containing subtree to the `#REF!` error literal.
  virtual std::optional<Reference> apply(const Reference& ref) const = 0;

  /// Returns the rewritten endpoints of a cell/range reference. The default
  /// implementation applies the ordinary endpoint transform independently;
  /// structural-edit transforms can override this to preserve a surviving
  /// portion of a range when an edit removes one of its endpoints.
  ///
  /// Returning `std::nullopt` collapses the whole range to `#REF!`. The
  /// endpoint references have already had any implicit sheet qualifier
  /// inherited by the walker, so an override can reason about both endpoints
  /// as one rectangle.
  virtual std::optional<std::pair<Reference, Reference>> apply_range(const Reference& lhs, const Reference& rhs) const;

  /// Whether a Ref3D node's inner cell coordinates must be carried through
  /// unchanged. Excel treats row/column structural edits this way: editing a
  /// sheet inside `First:Last!A1:B3` (or the sheet that owns the formula)
  /// does not rewrite the shared `A1:B3` tail. Relative-copy transforms still
  /// use the ordinary `apply` / `apply_range` path.
  virtual bool preserves_ref3d_coordinates() const noexcept { return false; }

  /// Rewrites the workbook-local sheet span carried by a `Ref3D` node.
  ///
  /// A 3-D reference stores its sheet endpoints outside the ordinary
  /// `Reference` payload, so a transform that edits sheet structure needs a
  /// dedicated hook rather than trying to route the endpoint names through
  /// `apply`. The default is an identity mapping. Returning `std::nullopt`
  /// collapses the complete `Ref3D` subtree to `#REF!`.
  struct Ref3DSheetSpan {
    std::string_view begin;
    std::string_view end;
  };
  virtual std::optional<Ref3DSheetSpan> apply_ref3d_span(std::string_view begin, std::string_view end) const;

  /// Hook for `ExternalRef` payloads. The default forwards to `apply`,
  /// which is sufficient for transforms that only manipulate row / column
  /// indices. Sheet-structure transforms can override this to keep external
  /// references opaque or to apply a workbook-specific policy.
  ///
  /// `book_id` and `sheet` are passed for context; transforms may inspect
  /// them but the return value is just the rewritten cell.
  virtual std::optional<Reference> apply_external(std::uint32_t book_id, std::string_view sheet,
                                                  const Reference& cell) const;

  /// Optional hook to rewrite the external-ref sheet field. Returning
  /// `std::nullopt` means "leave the sheet unchanged"; returning a string
  /// view replaces it. The walker is responsible for interning the new
  /// sheet name into the AST arena. This hook applies only to
  /// `ExternalRef`; workbook-local `Ref3D` endpoints use `apply_ref3d_span`
  /// above.
  ///
  /// The default implementation does nothing.
  virtual std::optional<std::string_view> transform_external_sheet(std::uint32_t book_id, std::string_view sheet) const;
};

/// Walks `root` and produces a new tree with every reference rewritten via
/// `transform`. When `transform` returns `std::nullopt` for any reference
/// (or when an external sheet rename collapses), the surrounding node is
/// replaced by `#REF!`.
///
/// Returns `nullptr` only on arena exhaustion. The original tree is left
/// untouched; the returned pointer aliases `&root` when no descendant
/// required rewriting (identity walks do not allocate).
const AstNode* shift_refs(const AstNode& root, Arena& arena, const RefTransform& transform);

/// Convenience wrapper for the historical relative-shift transform.
///
/// Each non-absolute coordinate of every reference inside `root` is
/// shifted by `(row_delta, col_delta)`. Absolute coordinates (`$A`, `$1`)
/// are preserved; full-column / full-row references shift only their
/// non-absolute axis. Out-of-bounds shifts collapse to `#REF!`.
///
/// The semantics here match what conditional-formatting per-cell
/// evaluation requires: rule formulas authored against the rule's anchor
/// cell are re-anchored to the candidate cell's coordinates before
/// evaluation. The same primitive will drive row/column insert/delete in a
/// follow-up bundle.
const AstNode* shift_relative_refs(const AstNode& root, Arena& arena, std::int32_t row_delta, std::int32_t col_delta);

}  // namespace parser
}  // namespace formulon

#endif  // FORMULON_PARSER_AST_SHIFT_H_
