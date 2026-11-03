#include "utils/derived_guid.h"

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "utils/strings.h"

namespace formulon {

namespace {

// FNV-1a 128-bit state as two 64-bit halves. The prime is 2^88 + 0x13B, so a
// multiply is `h * 0x13B + (h << 88)` modulo 2^128.
struct Fnv128 {
  std::uint64_t hi = 0x6C62272E07BB0142ULL;
  std::uint64_t lo = 0x62B821756295C58DULL;

  void add(unsigned char byte) noexcept {
    lo ^= byte;
    constexpr std::uint64_t kPrimeLow = 0x13BULL;
    const std::uint64_t p0 = (lo & 0xFFFFFFFFULL) * kPrimeLow;
    const std::uint64_t p1 = (lo >> 32) * kPrimeLow;
    const std::uint64_t mid = (p0 >> 32) + (p1 & 0xFFFFFFFFULL);
    const std::uint64_t carry = (p1 >> 32) + (mid >> 32);
    const std::uint64_t new_lo = (p0 & 0xFFFFFFFFULL) | (mid << 32);
    hi = hi * kPrimeLow + carry + (lo << 24);
    lo = new_lo;
  }

  void add(std::string_view bytes) noexcept {
    for (const char c : bytes) {
      add(static_cast<unsigned char>(c));
    }
  }
};

}  // namespace

std::string derived_guid(std::string_view source_guid, std::string_view salt) {
  Fnv128 hash;
  hash.add(source_guid);
  hash.add(static_cast<unsigned char>(0U));
  hash.add(salt);

  std::array<unsigned char, 16> bytes{};
  for (std::size_t i = 0; i < 8U; ++i) {
    bytes[i] = static_cast<unsigned char>(hash.hi >> (56U - 8U * i));
    bytes[8U + i] = static_cast<unsigned char>(hash.lo >> (56U - 8U * i));
  }
  bytes[6] = static_cast<unsigned char>((bytes[6] & 0x0FU) | 0x40U);
  bytes[8] = static_cast<unsigned char>((bytes[8] & 0x3FU) | 0x80U);

  const bool braced = source_guid.size() >= 2U && source_guid.front() == '{' && source_guid.back() == '}';
  std::string out;
  out.reserve(38U);
  if (braced) {
    out.push_back('{');
  }
  for (std::size_t i = 0; i < bytes.size(); ++i) {
    if (i == 4U || i == 6U || i == 8U || i == 10U) {
      out.push_back('-');
    }
    out.push_back(strings::hex_digit_upper(bytes[i] >> 4U));
    out.push_back(strings::hex_digit_upper(bytes[i] & 0x0FU));
  }
  if (braced) {
    out.push_back('}');
  }
  return out;
}

void rederive_attr_guid(std::string& attrs, std::string_view key, std::string_view salt) {
  const std::string needle = " " + std::string(key) + "=\"";
  const std::size_t at = attrs.find(needle);
  if (at == std::string::npos) {
    return;
  }
  const std::size_t begin = at + needle.size();
  const std::size_t end = attrs.find('"', begin);
  if (end == std::string::npos) {
    return;
  }
  attrs.replace(begin, end - begin, derived_guid(std::string_view(attrs).substr(begin, end - begin), salt));
}

}  // namespace formulon
