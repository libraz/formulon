//
// `index_from_double(v, limit)` narrows an externally supplied `double` to a
// container index, and only when the result is in range.
//
// A bare `static_cast<std::size_t>(v)` guarded by nothing more than `v < 0`
// is undefined for NaN (which fails that test) and for any magnitude the
// destination type cannot represent. On wasm32 the compiler lowers such a
// narrowing to a trapping `i32.trunc_f64_u`, so one hostile value reaching
// the cast aborts the whole instance rather than producing a wrong answer.
// Range-checking against the container bound before the cast makes the
// narrowing defined by construction.

#ifndef FORMULON_UTILS_CHECKED_INDEX_H_
#define FORMULON_UTILS_CHECKED_INDEX_H_

#include <cstddef>
#include <optional>

namespace formulon {

/// Converts an externally supplied `double` to an index into a container of
/// `limit` elements.
///
/// Returns `std::nullopt` unless `v` is non-negative and strictly below
/// `limit`; NaN and both infinities fail those comparisons and are rejected
/// with them, as is every input when `limit` is 0. `-0.0` converts to 0.
inline std::optional<std::size_t> index_from_double(double v, std::size_t limit) noexcept {
  if (!(v >= 0.0) || !(v < static_cast<double>(limit))) {
    return std::nullopt;
  }
  const auto index = static_cast<std::size_t>(v);
  // `static_cast<double>(limit)` rounds to nearest, so for a `limit` above
  // 2^53 the bound above can admit a value that truncates to `limit` itself.
  if (index >= limit) {
    return std::nullopt;
  }
  return index;
}

}  // namespace formulon

#endif  // FORMULON_UTILS_CHECKED_INDEX_H_
