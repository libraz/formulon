//
// One shared sort body for the whole engine.
//
// `std::sort` is instantiated per (comparator type, iterator type) pair, and
// libc++ materialises the full introsort -- quicksort partition, the
// sort3/sort4/sort5 networks and the insertion-sort tail -- into every one of
// them. A single instantiation over a heap-owning element type costs several
// kilobytes of code, so a handful of unrelated call sites turn into a
// measurable share of the WASM binary.
//
// The engine already applied the fix locally in `GroupIndexOrder`
// (`eval/groupby_pivotby/common.h`): sort indices through one type-erased
// comparator so a single sort body serves every axis order. This header
// generalises that owner. Call sites keep their own element comparator; what
// they no longer keep is a private copy of the sort.
//
// The order is stable: equal elements retain their input order, because the
// shared comparator falls back to the index. That is a stronger guarantee
// than `std::sort` gives, and it costs one extra comparator call per
// comparison plus the index vector -- the trade this file exists to make.

#ifndef FORMULON_UTILS_INDEX_SORT_H_
#define FORMULON_UTILS_INDEX_SORT_H_

#include <cstdint>
#include <utility>
#include <vector>

namespace formulon {

/// Type-erased "does index `lhs` order before index `rhs`" predicate.
///
/// `context` is borrowed, not owned: it has to outlive the sort call. The
/// factories below are the only intended way to build one.
struct IndexLess {
  const void* context;
  bool (*invoke)(const void*, std::uint32_t, std::uint32_t);

  bool operator()(std::uint32_t first_index, std::uint32_t second_index) const {
    return invoke(context, first_index, second_index);
  }
};

/// Wraps any callable taking two element indices. The callable is borrowed;
/// bind it to a named local so it outlives the returned predicate.
template <typename Less>
IndexLess make_index_less(const Less& less) {
  return IndexLess{&less, [](const void* context, std::uint32_t lhs, std::uint32_t rhs) {
                     return (*static_cast<const Less*>(context))(lhs, rhs);
                   }};
}

/// Arranges the indices already in `order` by `less`, ties keeping their input
/// order. This is the only sort body in the binary for every caller routed
/// through this header.
void sort_index_order(std::vector<std::uint32_t>& order, IndexLess less);

/// Fills `order` with `0 .. count - 1` and arranges it by `less`. Same
/// guarantees as `sort_index_order`; the convenience is the fill.
void sorted_index_order(std::vector<std::uint32_t>& order, std::uint32_t count, IndexLess less);

/// Rearranges `values` into `order`, moving each element exactly once.
template <typename T>
void apply_index_order(std::vector<T>& values, const std::vector<std::uint32_t>& order) {
  std::vector<T> reordered;
  reordered.reserve(order.size());
  for (const std::uint32_t index : order) {
    reordered.push_back(std::move(values[index]));
  }
  values.swap(reordered);
}

/// Stable-sorts `values` with an element-level comparator, without giving the
/// call site its own copy of the sort body.
///
/// `element_less` is a strict weak ordering over two elements, exactly as it
/// would be written for `std::sort`. Compared with `std::sort` this adds an
/// index vector and one indirect call per comparison; it is meant for the
/// cold and mid-frequency paths where that is the cheaper side of the trade.
template <typename T, typename Less>
void sort_by_index(std::vector<T>& values, const Less& element_less) {
  if (values.size() < 2U) {
    return;
  }
  const auto index_less = [&values, &element_less](std::uint32_t lhs, std::uint32_t rhs) {
    return element_less(values[lhs], values[rhs]);
  };
  std::vector<std::uint32_t> order;
  sorted_index_order(order, static_cast<std::uint32_t>(values.size()), make_index_less(index_less));
  apply_index_order(values, order);
}

}  // namespace formulon

#endif  // FORMULON_UTILS_INDEX_SORT_H_
