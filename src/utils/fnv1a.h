//
// `fnv1a64(data)` computes the 64-bit FNV-1a hash of a byte sequence: a
// small, dependency-free, deterministic fingerprint (not a cryptographic
// hash) for content-identity checks such as detecting whether a retained
// package part still matches the model state it was derived from.

#ifndef FORMULON_UTILS_FNV1A_H_
#define FORMULON_UTILS_FNV1A_H_

#include <cstdint>
#include <string_view>

namespace formulon {

/// 64-bit FNV-1a offset basis.
inline constexpr std::uint64_t kFnv1a64OffsetBasis = 0xcbf29ce484222325ULL;
/// 64-bit FNV-1a prime.
inline constexpr std::uint64_t kFnv1a64Prime = 0x100000001b3ULL;

/// 64-bit FNV-1a over `data`. Deterministic across processes and
/// platforms: no seed, and bytes are folded in one at a time regardless of
/// endianness.
inline std::uint64_t fnv1a64(std::string_view data) {
  std::uint64_t hash = kFnv1a64OffsetBasis;
  for (char c : data) {
    hash ^= static_cast<std::uint64_t>(static_cast<unsigned char>(c));
    hash *= kFnv1a64Prime;
  }
  return hash;
}

}  // namespace formulon

#endif  // FORMULON_UTILS_FNV1A_H_
