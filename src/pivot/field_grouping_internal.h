#ifndef FORMULON_PIVOT_FIELD_GROUPING_INTERNAL_H_
#define FORMULON_PIVOT_FIELD_GROUPING_INTERNAL_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::pivot::detail {

/// An `Error` carrying `code`, `message` and the optional `context`.
Error group_error(FormulonErrorCode code, std::string message, std::string context = {});

/// Group items over `base` in index order, and whether they are the trailing
/// fields of the cache.
struct GroupRun {
  std::size_t first = 0;
  std::size_t count = 0;
};

Expected<GroupRun, Error> group_run(const PivotCache& cache, std::size_t base);

/// Checks `spec` against the base data and resolves the bounds the group
/// fields take.
Expected<GroupBounds, Error> resolve_date_group_bounds(const PivotCache& cache, std::size_t base,
                                                       const DateGroupSpec& spec, bool date1904);

/// `code` when `state`'s items would outnumber what a pivot field holds
/// (Excel's unique-item limit), with context
/// `part=pivot_field_items requested=<n> limit=<n>`.
Expected<void, Error> check_numeric_item_count(const NumericGroupState& state, FormulonErrorCode code);

/// Checks a numeric grouping of `base` against its present values and
/// resolves the bounds an empty `start` / `end` takes from them. Errors as
/// `set_numeric_group`, plus `kInvalidArgument` for more group items than a
/// pivot field holds.
Expected<NumericGroupState, Error> resolve_numeric_group(const PivotCache& cache, std::size_t base,
                                                         std::optional<double> start, std::optional<double> end,
                                                         double interval);

/// Makes `group` the numeric range of `state` and regenerates its group items
/// as text in `cache`, as `set_numeric_group` describes them, with
/// `profile`'s decimal separator. Fails with `kPreconditionFailed` for more
/// group items than a pivot field holds, leaving `group` as it was.
Expected<void, Error> set_numeric_range(PivotCache& cache, FieldGroup& group, const NumericGroupState& state,
                                        ExcelProfile profile);

/// The group items of a calendar range group by `period` over the days
/// `start` to `end` (`interval` days a bucket for `GroupBy::Days`), labelled
/// in `profile`'s locale: `<start`, the periods, `>end`.
std::vector<std::string> date_range_group_labels(GroupBy period, std::uint32_t interval, double start, double end,
                                                 bool date1904, ExcelProfile profile);

/// Appends to `cache` the numeric group field of `state` over database field
/// `base`, named as the base plus the smallest free suffix from 2, points the
/// base's `par` at it and materialises the cache's group fields. Its items
/// are labelled in `profile`'s locale. Returns its index.
Expected<std::size_t, Error> group_cache_by_number(PivotCache& cache, std::size_t base, const NumericGroupState& state,
                                                   bool date1904, ExcelProfile profile);

}  // namespace formulon::pivot::detail

#endif  // FORMULON_PIVOT_FIELD_GROUPING_INTERNAL_H_
