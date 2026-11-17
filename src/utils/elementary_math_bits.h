//
// Bit-level views of a double shared by the utils/elementary_math sources.

#ifndef FORMULON_UTILS_ELEMENTARY_MATH_BITS_H_
#define FORMULON_UTILS_ELEMENTARY_MATH_BITS_H_

#include <cstdint>
#include <cstring>

namespace formulon {
namespace elementary {
namespace detail {

inline double from_bits(std::uint64_t u) noexcept {
  double d;
  std::memcpy(&d, &u, sizeof(d));
  return d;
}

inline std::uint64_t to_bits(double d) noexcept {
  std::uint64_t u;
  std::memcpy(&u, &d, sizeof(u));
  return u;
}

}  // namespace detail
}  // namespace elementary
}  // namespace formulon

#endif  // FORMULON_UTILS_ELEMENTARY_MATH_BITS_H_
