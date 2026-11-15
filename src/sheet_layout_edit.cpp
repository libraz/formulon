#include "sheet_layout_edit.h"

#include <algorithm>
#include <cstdint>
#include <limits>
#include <utility>
#include <vector>

#include "utils/index_sort.h"

namespace formulon {

void overlay_column_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, ColumnSpanApply apply,
                         const void* ctx) {
  std::vector<ColumnLayout> next;
  next.reserve(layout.columns.size() + 1U);
  std::vector<std::pair<std::uint32_t, std::uint32_t>> covered;
  covered.reserve(layout.columns.size());

  for (const ColumnLayout& entry : layout.columns) {
    if (entry.last < first || entry.first > last) {
      next.push_back(entry);
      continue;
    }

    if (entry.first < first) {
      ColumnLayout left = entry;
      left.last = first - 1U;
      next.push_back(left);
    }

    const std::uint32_t intersection_first = std::max(entry.first, first);
    const std::uint32_t intersection_last = std::min(entry.last, last);
    ColumnLayout intersection = entry;
    intersection.first = intersection_first;
    intersection.last = intersection_last;
    apply(intersection, ctx);
    next.push_back(intersection);
    covered.emplace_back(intersection_first, intersection_last);

    if (entry.last > last) {
      ColumnLayout right = entry;
      right.first = last + 1U;
      next.push_back(right);
    }
  }

  // Merge covered intervals before synthesising gaps. Normal SheetLayout
  // entries are disjoint, but this also avoids duplicating a default segment
  // when a caller supplied overlapping aggregate entries directly.
  sort_by_index(covered, [](const std::pair<std::uint32_t, std::uint32_t>& lhs,
                            const std::pair<std::uint32_t, std::uint32_t>& rhs) { return lhs < rhs; });
  std::vector<std::pair<std::uint32_t, std::uint32_t>> covered_union;
  for (const auto& interval : covered) {
    if (covered_union.empty() ||
        static_cast<std::uint64_t>(interval.first) > static_cast<std::uint64_t>(covered_union.back().second) + 1U) {
      covered_union.push_back(interval);
    } else if (interval.second > covered_union.back().second) {
      covered_union.back().second = interval.second;
    }
  }

  std::uint64_t cursor = first;
  for (const auto& interval : covered_union) {
    if (cursor < interval.first) {
      ColumnLayout gap;
      gap.first = static_cast<std::uint32_t>(cursor);
      gap.last = interval.first - 1U;
      apply(gap, ctx);
      next.push_back(gap);
    }
    cursor = std::max(cursor, static_cast<std::uint64_t>(interval.second) + 1U);
  }
  if (cursor <= last) {
    ColumnLayout gap;
    gap.first = static_cast<std::uint32_t>(cursor);
    gap.last = last;
    apply(gap, ctx);
    next.push_back(gap);
  }

  sort_by_index(next, ColumnSpanOrder{});
  std::vector<ColumnLayout> merged;
  merged.reserve(next.size());
  for (const ColumnLayout& entry : next) {
    if (!merged.empty() && merged.back().last != std::numeric_limits<std::uint32_t>::max() &&
        merged.back().last + 1U == entry.first && SameColumnLayoutState(merged.back(), entry)) {
      merged.back().last = entry.last;
    } else {
      merged.push_back(entry);
    }
  }
  layout.columns = std::move(merged);
}

}  // namespace formulon
