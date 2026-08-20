//
// Shared scalar reduction for the `@` operator and `_xlfn.SINGLE`.

#ifndef FORMULON_EVAL_IMPLICIT_INTERSECTION_H_
#define FORMULON_EVAL_IMPLICIT_INTERSECTION_H_

#include <cstdint>
#include <optional>

#include "parser/ast.h"
#include "parser/reference.h"
#include "value.h"

namespace formulon::eval {

/// Applies implicit intersection to an already-evaluated value that has no
/// static range coordinates left to project. Scalars pass through unchanged;
/// dynamic arrays use their top-left element. Empty/corrupt arrays are not a
/// valid scalar and surface `#VALUE!`.
inline Value implicit_intersect_value(Value value) {
  if (!value.is_array()) {
    return value;
  }
  if (value.as_array_rows() == 0U || value.as_array_cols() == 0U || value.as_array_cells() == nullptr) {
    return Value::error(ErrorCode::Value);
  }
  return value.as_array_cells()[0];
}

/// Projects the formula cell onto the rectangle an endpoint pair declares,
/// yielding the single cell implicit intersection selects.
///
/// The rectangle comes from `declared_rect`, so a full-axis endpoint spans its
/// grid axis instead of contributing the unspecified coordinate the parser left
/// in that field. A single-column rectangle takes the formula row, a single-row
/// rectangle takes the formula column, and a 2-D rectangle intersects on both
/// axes. A formula cell outside the relevant span — and a pair that declares no
/// rectangle at all — has nothing to project onto and yields `std::nullopt`,
/// which every call site surfaces as `#VALUE!`. The result inherits
/// `lhs.sheet`, matching Excel's treatment of the left endpoint as the
/// qualifier for the whole range.
std::optional<parser::Reference> project_implicit_intersection(const parser::Reference& lhs,
                                                               const parser::Reference& rhs, std::uint32_t formula_row,
                                                               std::uint32_t formula_col);

/// Outcome of projecting an implicit-intersection operand onto the formula
/// cell.
enum class IntersectionProjection : std::uint8_t {
  /// The operand declares a rectangle and `*out_target` names the cell to
  /// read.
  kCell,
  /// The operand is reference-shaped but yields no cell here: the formula sits
  /// outside the axis it projects on, or the endpoints declare no rectangle.
  /// Both are `#VALUE!`.
  kNoCell,
  /// The operand retains no static coordinates to project. The caller
  /// evaluates it and reduces the result with `implicit_intersect_value`.
  kNotStaticReference,
};

/// Resolves the operand of `@` / `_xlfn.SINGLE` against a bound formula cell.
///
/// Both spellings route through here so they cannot drift on which operand
/// shapes project. The three reference spellings that declare a rectangle —
/// bounded `Ref:Ref`, full-axis `Ref:Ref`, and the single `Ref` the parser
/// folds `A:A` / `1:1` into — take the projection above. A bounded single
/// `Ref` is already the scalar `@` asks for and reports
/// `kNotStaticReference`, as does any dynamic-array producer.
IntersectionProjection project_implicit_intersection(const parser::AstNode& operand, std::uint32_t formula_row,
                                                     std::uint32_t formula_col, parser::Reference* out_target);

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_IMPLICIT_INTERSECTION_H_
