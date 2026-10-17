//
// Simple (1:1) Unicode case mapping as UPPER, LOWER and PROPER apply it.
// Characters outside the BMP map to themselves, and no mapping changes a
// character's length. Invalid UTF-8 bytes pass through unchanged.

#ifndef FORMULON_UTILS_CASE_MAP_H_
#define FORMULON_UTILS_CASE_MAP_H_

#include <cstdint>
#include <string>
#include <string_view>

namespace formulon {

/// Upper-case form of `cp`.
std::uint32_t case_map_upper(std::uint32_t cp) noexcept;

/// Lower-case form of `cp` taken on its own: U+03A3 gives the final form U+03C2.
std::uint32_t case_map_lower_isolated(std::uint32_t cp) noexcept;

/// UPPER(text).
std::string case_map_to_upper(std::string_view text);

/// LOWER(text). U+03A3 becomes U+03C2 when it ends the text or precedes a
/// space or ASCII punctuation mark other than `( ) * ^ ` ~`, else U+03C3.
std::string case_map_to_lower(std::string_view text);

/// PROPER(text): a character after a non-letter takes its word-start form, a
/// letter after a letter its lower-case form.
std::string case_map_to_proper(std::string_view text);

}  // namespace formulon

#endif  // FORMULON_UTILS_CASE_MAP_H_
