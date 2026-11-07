#ifndef FORMULON_PIVOT_FIELD_GROUPING_INTERNAL_H_
#define FORMULON_PIVOT_FIELD_GROUPING_INTERNAL_H_

#include <cstddef>
#include <string>

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

}  // namespace formulon::pivot::detail

#endif  // FORMULON_PIVOT_FIELD_GROUPING_INTERNAL_H_
