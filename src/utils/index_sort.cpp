//
// The single shared sort body. See `utils/index_sort.h` for why it exists.

#include "utils/index_sort.h"

#include <algorithm>

namespace formulon {

void sorted_index_order(std::vector<std::uint32_t>& order, std::uint32_t count, IndexLess less) {
  order.resize(count);
  for (std::uint32_t index = 0; index < count; ++index) {
    order[index] = index;
  }
  sort_index_order(order, less);
}

void sort_index_order(std::vector<std::uint32_t>& order, IndexLess less) {
  // The index fallback is what makes the order stable, so callers never have
  // to reach for `std::stable_sort` -- which would pull in the merge path and
  // its temporary buffer on top of this body.
  std::sort(order.begin(), order.end(), [&less](std::uint32_t lhs, std::uint32_t rhs) {
    if (less(lhs, rhs)) {
      return true;
    }
    if (less(rhs, lhs)) {
      return false;
    }
    return lhs < rhs;
  });
}

}  // namespace formulon
