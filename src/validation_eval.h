//
// Data-validation evaluation: whether a cell value satisfies the rule that applies to it.

#ifndef FORMULON_VALIDATION_EVAL_H_
#define FORMULON_VALIDATION_EVAL_H_

#include <cstddef>
#include <cstdint>

#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

class Sheet;
class Workbook;

namespace eval {
class FunctionRegistry;
}  // namespace eval

/// Collaborators for rule formula evaluation.
struct ValidationEvalDeps {
  /// Function dispatch table; `nullptr` selects `eval::default_registry()`.
  const eval::FunctionRegistry* registry = nullptr;
};

/// Result of checking one value against the rule covering a cell.
struct ValidationOutcome {
  /// False when no rule covers the cell; `valid` is then true.
  bool has_rule = false;
  bool valid = true;
  /// Index into `Sheet::validations()` of the applied rule (meaningful only when `has_rule`).
  std::uint32_t rule_index = 0;
  /// The applied rule's `error_style` (0 stop, 1 warning, 2 information).
  std::uint8_t error_style = 0;
};

/// Checks `proposed` against the data-validation rule covering `(row, col)` on `sheet`.
///
/// Excel applies one rule per cell: when several sqrefs overlap, the first rule in
/// `Sheet::validations()` order wins. `formula1` / `formula2` are evaluated through the
/// formula evaluator, shifted from the sqref's top-left cell to the target cell. `proposed`
/// is an already-parsed value (text such as "5" is not reinterpreted as a number); numeric
/// types reject text, and `textLength` counts UTF-16 units of the value's General text.
/// A blank value is valid under `allow_blank`, otherwise it is checked as 0 / empty text.
/// `custom` formulas read the workbook as stored, so a formula that refers to the cell itself
/// sees the stored value rather than `proposed`.
///
/// Fails with `kInvalidArgument` when `(row, col)` is outside the grid.
Expected<ValidationOutcome, Error> validate_value(const Workbook& wb, const Sheet& sheet, std::uint32_t row,
                                                  std::uint32_t col, const Value& proposed,
                                                  const ValidationEvalDeps& deps = {});

}  // namespace formulon

#endif  // FORMULON_VALIDATION_EVAL_H_
