//
// Name tables mapping an enum's ordinal to its OOXML attribute spelling, shared by the
// reader (name -> enum) and the writer (enum -> name) so the two directions cannot drift.

#ifndef FORMULON_IO_ENUM_NAME_TABLE_H_
#define FORMULON_IO_ENUM_NAME_TABLE_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <string_view>

namespace formulon::io {

/// One `(enum value, attribute spelling)` pair. Several entries may share a
/// value or a name; lookups return the first match.
struct EnumName {
  template <typename E>
  constexpr EnumName(E v, std::string_view n) : value(static_cast<std::uint8_t>(v)), name(n) {}

  std::uint8_t value;
  std::string_view name;
};

/// Returns the value of the first entry named `text`, or `fallback`.
inline std::uint8_t find_enum_value(const EnumName* entries, std::size_t count, std::string_view text,
                                    std::uint8_t fallback) {
  for (std::size_t i = 0; i < count; ++i) {
    if (entries[i].name == text) {
      return entries[i].value;
    }
  }
  return fallback;
}

/// Returns the name of the first entry holding `value`, or `fallback`.
inline std::string_view find_enum_name(const EnumName* entries, std::size_t count, std::uint8_t value,
                                       std::string_view fallback) {
  for (std::size_t i = 0; i < count; ++i) {
    if (entries[i].value == value) {
      return entries[i].name;
    }
  }
  return fallback;
}

/// Typed front end of `find_enum_value`.
template <std::size_t N, typename E>
E enum_from_name(const std::array<EnumName, N>& table, std::string_view text, E fallback) {
  return static_cast<E>(find_enum_value(table.data(), N, text, static_cast<std::uint8_t>(fallback)));
}

/// Typed front end of `find_enum_name`.
template <std::size_t N, typename E>
std::string_view enum_name(const std::array<EnumName, N>& table, E value, std::string_view fallback = {}) {
  return find_enum_name(table.data(), N, static_cast<std::uint8_t>(value), fallback);
}

}  // namespace formulon::io

#endif  // FORMULON_IO_ENUM_NAME_TABLE_H_
