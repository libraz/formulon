//
// Group-item numbering for pivot group fields.
//
// A group field (`databaseField="0"` with `<fieldGroup base>`) has no source
// column; each record's value is the index into its `<groupItems>` that the
// base field's value falls in. A range group reserves item 0 for values below
// its start and the last item for values above its end. The numbers are
// materialised into `PivotCacheRecord::cells` once, so every consumer reads a
// group field through `cell_value` like any database field.

#ifndef FORMULON_PIVOT_FIELD_GROUPING_H_
#define FORMULON_PIVOT_FIELD_GROUPING_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>

#include "pivot/pivot_cache.h"
#include "value.h"

namespace formulon::pivot {

/// Inclusive bounds of a range group: date serials for a date group, plain
/// numbers otherwise.
struct GroupBounds {
  double start = 0.0;
  double end = 0.0;
};

/// True when `field` is a group field whose cells are derived from a base.
bool is_group_field(const PivotCacheField& field) noexcept;

/// Serial of an ISO 8601 group bound (`startDate` / `endDate`) under the given
/// epoch; empty when `iso` does not parse.
std::optional<double> group_date_serial(std::string_view iso, bool date1904) noexcept;

/// Bounds `range` states: its parsed `startDate` / `endDate`, or
/// `start_num` / `end_num` for a numeric group. Empty when a date bound is
/// absent or does not parse.
std::optional<GroupBounds> stated_group_bounds(const RangeGroup& range, bool date1904) noexcept;

/// Group-item number of base value `value` under `range` with `bounds` and
/// `item_count` group items: 0 below the start, the last item above the end,
/// and 0 for a blank or a value outside the date domain.
std::uint32_t range_group_item_number(const RangeGroup& range, const GroupBounds& bounds, std::size_t item_count,
                                      const Value& value, bool date1904) noexcept;

/// Group-item number of the base field's shared item `base_item` under a
/// discrete group; empty when the map does not cover it.
std::optional<std::uint32_t> discrete_group_item_number(const FieldGroup& group, std::size_t base_item) noexcept;

/// Writes every group field's group-item number into each record's cell for
/// that field, as a non-index, non-date number. A date range group without
/// stated bounds takes the base data's first day and the day after its last.
void materialize_group_fields(PivotCache& cache, bool date1904);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FIELD_GROUPING_H_
