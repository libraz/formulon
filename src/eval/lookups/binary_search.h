//
// Search kernels shared by the lookup family. Each takes `cmp(i)`, the sign of
// (cell i - lookup value); the classic approximate paths additionally return
// `kForeignType` when cell i is not of the lookup value's type.

#ifndef FORMULON_EVAL_LOOKUPS_BINARY_SEARCH_H_
#define FORMULON_EVAL_LOOKUPS_BINARY_SEARCH_H_

#include <cstddef>
#include <cstdint>

namespace formulon {
namespace eval {

// Returned by a classic comparer for a cell of another type than the lookup value.
constexpr int kForeignType = 2;

// Outcome of `sorted_search`: the remembered equal probe (SIZE_MAX if none) and
// the closed-interval bounds left when the search ended.
struct SortedSearchResult {
  std::size_t found;
  std::ptrdiff_t lo;
  std::ptrdiff_t hi;
};

// Closed-interval binary search with mid = (lo + hi) >> 1 and one comparison
// per probe. An equal probe is remembered and the search moves on toward the
// start of the sorted order (`prefer_last` false) or the end (true).
template <typename Cmp>
SortedSearchResult sorted_search(std::size_t n, bool descending, bool prefer_last, const Cmp& cmp) {
  std::ptrdiff_t lo = 0;
  std::ptrdiff_t hi = static_cast<std::ptrdiff_t>(n) - 1;
  std::size_t found = SIZE_MAX;
  while (lo <= hi) {
    const std::ptrdiff_t mid = (lo + hi) >> 1;
    const int c = cmp(static_cast<std::size_t>(mid));
    if (c == 0) {
      found = static_cast<std::size_t>(mid);
      if (prefer_last) {
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    } else if ((c < 0) != descending) {
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return {found, lo, hi};
}

// MATCH(..., 1), LOOKUP and VLOOKUP / HLOOKUP with range_lookup TRUE. Binary
// search where an equal probe continues right and a smaller one becomes the
// candidate. A foreign-type probe (another type, or an error) is looked past:
// the first same-type cell after it decides, a greater one sends the search
// left, a smaller one becomes the candidate, and an equal one ends the search
// at the last cell of its run. Returns SIZE_MAX when nothing matches.
template <typename Cmp>
std::size_t approximate_search(std::size_t n, const Cmp& cmp) {
  std::ptrdiff_t lo = 0;
  std::ptrdiff_t hi = static_cast<std::ptrdiff_t>(n) - 1;
  std::size_t found = SIZE_MAX;
  std::size_t candidate = SIZE_MAX;
  while (lo <= hi) {
    const std::ptrdiff_t mid = (lo + hi) >> 1;
    int c = cmp(static_cast<std::size_t>(mid));
    if (c == kForeignType) {
      std::size_t j = static_cast<std::size_t>(mid) + 1U;
      while (j < n && (c = cmp(j)) == kForeignType) {
        ++j;
      }
      if (j >= n || c > 0) {
        hi = mid - 1;
      } else if (c < 0) {
        candidate = j;
        lo = static_cast<std::ptrdiff_t>(j) + 1;
      } else {
        while (j + 1U < n && cmp(j + 1U) == 0) {
          ++j;
        }
        return j;
      }
      continue;
    }
    if (c == 0) {
      found = static_cast<std::size_t>(mid);
      lo = mid + 1;
    } else if (c < 0) {
      candidate = static_cast<std::size_t>(mid);
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return found != SIZE_MAX ? found : candidate;
}

// MATCH(..., -1): linear scan over same-type cells; an equal cell returns at
// once, a greater one becomes the candidate, the first smaller one stops.
template <typename Cmp>
std::size_t descending_scan(std::size_t n, const Cmp& cmp) {
  std::size_t candidate = SIZE_MAX;
  for (std::size_t i = 0; i < n; ++i) {
    const int c = cmp(i);
    if (c == kForeignType) {
      continue;
    }
    if (c == 0) {
      return i;
    }
    if (c < 0) {
      break;
    }
    candidate = i;
  }
  return candidate;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_LOOKUPS_BINARY_SEARCH_H_
