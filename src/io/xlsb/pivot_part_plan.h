//
// The pivot parts an XLSB save writes, planned one cache group at a time: a
// cache with every table drawing from it. A group read from an `.xlsb` and
// unchanged since (every retained part still fingerprints as loaded, no table
// edited) is written from its retained bytes. Any other group is encoded from
// the model through the bridges with the workbook profile's host, and a group
// the bridges refuse is left out whole and counted by the caller as deferred.
// Internal to `src/io/xlsb/`.

#ifndef FORMULON_IO_XLSB_PIVOT_PART_PLAN_H_
#define FORMULON_IO_XLSB_PIVOT_PART_PLAN_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "io/xlsb/emission_plan.h"
#include "passthrough_part.h"

namespace formulon {

class Workbook;

namespace io {
namespace xlsb {

/// What `plan_pivot_parts` decided for each table.
struct PivotPartOutcome {
  /// Per sheet, per pivot: its table part is written, planned or kept passthrough with its group.
  std::vector<std::vector<bool>> written;
  /// Tables left out, per sheet.
  std::vector<std::size_t> lost_tables;
  /// Caches left out that no table draws from.
  std::size_t lost_caches = 0;
};

/// Fills `plan.pivot_caches` and `plan.sheet_pivots`, numbering generated
/// parts clear of `taken`. A cache no table draws from is planned like any
/// other group: kept while unchanged, else encoded, else left out.
PivotPartOutcome plan_pivot_parts(const Workbook& wb, const std::unordered_set<std::string>& taken, EmissionPlan& plan);

/// True for a retained pivot part whose group the pivot plan writes, or not;
/// such a part never goes through the passthrough list. A group whose table
/// parts are retained without its cache parts stays passthrough while fresh.
bool is_retained_pivot_part(const Workbook& wb, const PassthroughPart& part);

/// The PIVOTCACHEIDS block to splice: `retained`'s entries, then one per
/// planned cache under the `pivot_cache_rel_key` the workbook rels map.
std::vector<std::uint8_t> book_pivot_cache_ids(const std::vector<std::uint8_t>& retained, const EmissionPlan& plan);

/// The relationship target naming package part `to` from the rels of part `from`.
std::string relative_part_target(std::string_view from, std::string_view to);

/// The id a written cache definition names its records relationship by.
inline constexpr std::string_view kPivotCacheRecordsRelId = "rId1";

/// A `.rels` part relating its source to `target` by `rId1`.
std::string single_relationship_part(std::string_view type, std::string_view target);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_PART_PLAN_H_
