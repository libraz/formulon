//
// Internal contract between the collation key builder (`collation.cpp`) and
// its kana classification (`collation_kana.cpp`). Not part of the public API.

#ifndef FORMULON_EVAL_COLLATION_INTERNAL_H_
#define FORMULON_EVAL_COLLATION_INTERNAL_H_

#include <cstdint>

namespace formulon {
namespace eval {
namespace detail {

bool is_hiragana_letter(std::uint32_t cp) noexcept;
bool is_katakana_letter(std::uint32_t cp) noexcept;

/// The vowel kana a following prolonged-sound mark stands for, or 0.
std::uint32_t kana_vowel(std::uint32_t cp) noexcept;

/// The character composed from the first character of `cp`'s compatibility
/// form and a voiced sound mark, or 0 when that does not compose.
std::uint32_t voiced_form(std::uint32_t cp) noexcept;

/// Half-width kana followed by U+FF9E or U+FF9F composed into one full-width kana, or 0.
std::uint32_t halfwidth_voiced(std::uint32_t half, std::uint32_t mark) noexcept;

/// Kana-type weight at the dynamic tertiary level: katakana below hiragana, small below normal.
std::int32_t win_kana_d(std::uint32_t cp) noexcept;

}  // namespace detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COLLATION_INTERNAL_H_
