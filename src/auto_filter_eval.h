//
// AutoFilter evaluation: which rows of a filtered range satisfy its criteria.
//
// Every body row (the range minus its header row) is matched against each
// filter column; a row is visible when every column with a criterion matches.
// `filters` lists compare the cell's displayed text case-insensitively (a cell
// whose value overflows its format matches no listed value), custom filters
// use Excel's wildcard dialect on text and numeric comparison on numbers,
// `top10` and the average filters rank only the numeric cells of the column,
// and relative-date filters read "today" from the evaluation context's wall
// clock with weeks starting on Sunday. Colour filters compare the cell's
// effective fill or font colour, with matching conditional formats applied,
// against the filter's dxf; icon filters compare the conditional-format icon.

#ifndef FORMULON_AUTO_FILTER_EVAL_H_
#define FORMULON_AUTO_FILTER_EVAL_H_

#include <cstddef>
#include <vector>

#include "auto_filter.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Arena;
class Sheet;
class Workbook;

namespace eval {
class EvalContext;
class FunctionRegistry;
}  // namespace eval

/// Collaborators for criteria that need the formula evaluator. Every member
/// may be null: a null `registry` selects `eval::default_registry()`, and a
/// null `arena` or `eval_ctx` makes the call build its own, with the clock
/// pinned to `Workbook::pinned_now()` (the host clock when unpinned).
struct AutoFilterEvalDeps {
  Arena* arena = nullptr;
  const eval::FunctionRegistry* registry = nullptr;
  /// Supplies "today" for relative-date filters through `wall_clock()`.
  const eval::EvalContext* eval_ctx = nullptr;
};

/// Returns one flag per body row of `filter.range` (row `first_row + 1 + i`
/// at index `i`), true when the row satisfies every column criterion. The
/// sheet is not modified. Fails with `kAutoFilterInvalid` when `filter` is
/// opaque or fails `validate_auto_filter`.
Expected<std::vector<bool>, Error> evaluate_auto_filter(const Workbook& wb, const Sheet& sheet,
                                                        const AutoFilter& filter, const AutoFilterEvalDeps& deps = {});

/// Evaluates `filter` on sheet `sheet_index` and sets each body row's hidden
/// flag to the negation of its match, so a matching row hidden by hand is
/// shown again; rows outside the body are untouched. Row-visibility
/// dependents (SUBTOTAL / AGGREGATE) are marked dirty when any flag changes.
/// Fails with `kInvalidArgument` for a bad sheet index, otherwise as
/// `evaluate_auto_filter`.
Expected<void, Error> apply_auto_filter(Workbook& wb, std::size_t sheet_index, const AutoFilter& filter,
                                        const AutoFilterEvalDeps& deps = {});

/// True when the sheet AutoFilter, or the AutoFilter of a table on the
/// sheet, carries at least one criterion. Excel then treats every hidden row
/// of the sheet as filtered, inside the filter range or not.
bool sheet_has_filter_criteria(const Workbook* wb, const Sheet& sheet) noexcept;

}  // namespace formulon

#endif  // FORMULON_AUTO_FILTER_EVAL_H_
