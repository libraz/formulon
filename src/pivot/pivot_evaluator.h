//
// Pivot-table evaluator: given a populated `PivotTable` definition +
// the cache it references, produce a `PivotResult` containing the
// row/col hierarchies, the aggregated values matrix, subtotals, and
// the grand total.
//
// MVP scope (for GETPIVOTDATA): SUM / COUNT / AVERAGE / MAX / MIN /
// PRODUCT / CountNumbers aggregations; manual-filter visibility
// (PivotItem::visible); row/col hierarchies built from
// `row_field_order` / `col_field_order`; subtotals when the field
// declares `subtotal_top` or `subtotal_fns`; grand totals when the
// table flags request them. STDEV / VAR / "% of row" / running total /
// date grouping / slicer filters are deferred.
//
// Design references:
//   * src/pivot/pivot_result.h (output structure)

#ifndef FORMULON_PIVOT_PIVOT_EVALUATOR_H_
#define FORMULON_PIVOT_PIVOT_EVALUATOR_H_

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::pivot {

/// Evaluates `table` against `cache` and returns the result.
///
/// Returns `kEvalPivotMissing` when `table.pivot_cache_id()` does not
/// match `cache.cache_id()`. Returns `kEvalPivotInvalid` when a
/// `PivotDataField::field_index` is out of bounds against
/// `cache.fields()`.
Expected<PivotResult, Error> evaluate(const PivotTable& table, const PivotCache& cache);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_EVALUATOR_H_
