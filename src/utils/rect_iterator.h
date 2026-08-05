//
// `RectRange` is a header-only, Value-independent range adapter for the
// inclusive `[r0, r1] x [c0, c1]` row-major walks that every sheet-touching
// subsystem otherwise hand-rolls as a nested `for (row = r0; row <= r1; ...)`
// / `for (col = c0; col <= c1; ...)` pair.
//
// Rationale:
//   - The bare double-for is correct, but it is also the canonical site for
//     off-by-one mistakes and accidental transpositions. Centralising the walk
//     puts the row-major contract in exactly one place and lets the compiler
//     verify it via the iterator's `operator++`.
//   - The adapter is `constexpr` and inline; after LTO the generated code is
//     identical to the hand-rolled loop, so no WASM size regression is
//     expected.
//   - The struct intentionally depends on nothing but `<cstdint>` and
//     `<iterator>`. It must remain Value-independent so cf/, print/, eval/,
//     and io/ can all consume it without dragging in evaluator headers.
//
// Empty-range contract: when `r0 > r1` or `c0 > c1`, `begin() == end()` and
// `size() == 0`. This mirrors the natural fall-through of the nested
// `<=`-bounded loops the iterator replaces (the body simply never executes),
// so migrated call sites need no extra empty checks.

#ifndef FORMULON_UTILS_RECT_ITERATOR_H_
#define FORMULON_UTILS_RECT_ITERATOR_H_

#include <cstdint>
#include <iterator>

namespace formulon::utils {

/// Row-major `(row, col)` pair yielded by `RectRange::iterator`.
struct RectCoord {
  std::uint32_t row;
  std::uint32_t col;
};

/// Adapter that walks the inclusive rectangle `[r0, r1] x [c0, c1]` in
/// row-major order (outer = row, inner = col), exposing a forward iterator
/// usable in range-`for`.
///
/// `r0 > r1` or `c0 > c1` is treated as the empty range. The iterator never
/// throws and performs no allocation.
class RectRange {
 public:
  constexpr RectRange(std::uint32_t r0, std::uint32_t c0, std::uint32_t r1, std::uint32_t c1) noexcept
      : r0_(r0), c0_(c0), r1_(r1), c1_(c1) {}

  /// Forward iterator over `RectCoord`. The end sentinel is a distinct state
  /// (`row_ == r1_ + 1`, `col_ == c0_`) reached after the last valid cell.
  class iterator {
   public:
    using value_type = RectCoord;
    using reference = RectCoord;
    using pointer = void;
    using difference_type = std::ptrdiff_t;
    using iterator_category = std::forward_iterator_tag;

    constexpr iterator() noexcept = default;

    constexpr RectCoord operator*() const noexcept { return RectCoord{row_, col_}; }

    constexpr iterator& operator++() noexcept {
      if (col_ < c1_) {
        ++col_;
      } else {
        col_ = c0_;
        ++row_;
      }
      return *this;
    }

    constexpr iterator operator++(int) noexcept {
      iterator copy = *this;
      ++(*this);
      return copy;
    }

    constexpr bool operator==(const iterator& other) const noexcept { return row_ == other.row_ && col_ == other.col_; }
    constexpr bool operator!=(const iterator& other) const noexcept { return !(*this == other); }

   private:
    friend class RectRange;
    constexpr iterator(std::uint32_t row, std::uint32_t col, std::uint32_t c0, std::uint32_t c1) noexcept
        : row_(row), col_(col), c0_(c0), c1_(c1) {}

    std::uint32_t row_{0};
    std::uint32_t col_{0};
    std::uint32_t c0_{0};
    std::uint32_t c1_{0};
  };

  constexpr iterator begin() const noexcept {
    if (empty()) {
      return end();
    }
    return iterator{r0_, c0_, c0_, c1_};
  }

  constexpr iterator end() const noexcept {
    // Sentinel: one row past the last valid row, anchored at the starting
    // column so post-increment from the last `(r1_, c1_)` lands here exactly.
    // For an empty range this still compares equal to `begin()` because
    // `begin()` short-circuits to `end()`.
    return iterator{r1_ + 1U, c0_, c0_, c1_};
  }

  constexpr bool empty() const noexcept { return r0_ > r1_ || c0_ > c1_; }

  /// Total cell count as `uint64_t` so a full-sheet rectangle
  /// (`Sheet::kMaxRows * Sheet::kMaxCols`) cannot overflow.
  constexpr std::uint64_t size() const noexcept {
    if (empty()) {
      return 0U;
    }
    const std::uint64_t rows = static_cast<std::uint64_t>(r1_ - r0_) + 1U;
    const std::uint64_t cols = static_cast<std::uint64_t>(c1_ - c0_) + 1U;
    return rows * cols;
  }

  constexpr std::uint32_t r0() const noexcept { return r0_; }
  constexpr std::uint32_t c0() const noexcept { return c0_; }
  constexpr std::uint32_t r1() const noexcept { return r1_; }
  constexpr std::uint32_t c1() const noexcept { return c1_; }

 private:
  std::uint32_t r0_;
  std::uint32_t c0_;
  std::uint32_t r1_;
  std::uint32_t c1_;
};

}  // namespace formulon::utils

#endif  // FORMULON_UTILS_RECT_ITERATOR_H_
