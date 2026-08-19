//
// Reference-level half of implicit intersection, shared by the `@` operator in
// the tree walker and `_xlfn.SINGLE`. Keeping one body means the two spellings
// cannot drift on which formula-cell positions are in range.

#include "eval/implicit_intersection.h"

#include <algorithm>

namespace formulon::eval {

std::optional<parser::Reference> project_implicit_intersection(const parser::Reference& lhs,
                                                               const parser::Reference& rhs, std::uint32_t formula_row,
                                                               std::uint32_t formula_col) {
  const std::uint32_t row_lo = std::min(lhs.row, rhs.row);
  const std::uint32_t row_hi = std::max(lhs.row, rhs.row);
  const std::uint32_t col_lo = std::min(lhs.col, rhs.col);
  const std::uint32_t col_hi = std::max(lhs.col, rhs.col);

  parser::Reference target{};
  target.sheet = lhs.sheet;
  if (col_lo == col_hi) {
    // Single-column range: project the formula row.
    if (formula_row < row_lo || formula_row > row_hi) {
      return std::nullopt;
    }
    target.row = formula_row;
    target.col = col_lo;
    return target;
  }
  if (row_lo == row_hi) {
    // Single-row range: project the formula column.
    if (formula_col < col_lo || formula_col > col_hi) {
      return std::nullopt;
    }
    target.row = row_lo;
    target.col = formula_col;
    return target;
  }
  // 2-D range: Excel intersects on the formula cell's row AND column. When the
  // cell falls inside both spans the result is the single cell at
  // (formula_row, formula_col); any other position has no intersection.
  if (formula_row < row_lo || formula_row > row_hi || formula_col < col_lo || formula_col > col_hi) {
    return std::nullopt;
  }
  target.row = formula_row;
  target.col = formula_col;
  return target;
}

}  // namespace formulon::eval
