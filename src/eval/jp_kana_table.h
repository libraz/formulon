//
// Shared half-width-katakana to full-width-katakana base mapping. Both
// `jp_fold` (used by lookup-key folding) and `text_width` (used by JIS /
// DBCS / ASC) consume this table; previously each TU carried its own
// copy of the same 56-entry block.
//
// Only the half-to-full *base* table and composition lookup tables are
// shared. The full-to-half side is width-conversion-specific (entries
// decompose voiced forms into a base codepoint plus U+FF9E / U+FF9F so the
// half-width output is two codepoints) and stays inside `text_width.cpp`.

#ifndef FORMULON_EVAL_JP_KANA_TABLE_H_
#define FORMULON_EVAL_JP_KANA_TABLE_H_

#include <cstdint>

namespace formulon {
namespace eval {

/// Returns the full-width katakana base for half-width codepoint `cp`,
/// or 0 if `cp` does not fall in the half-width katakana base block
/// (U+FF66..U+FF9D). Use `half_to_full_punctuation` for U+FF61..U+FF65
/// and the spacing dakuten / handakuten (U+FF9E / U+FF9F) handling
/// remains caller-specific.
std::uint32_t half_to_full_katakana_base(std::uint32_t cp) noexcept;

/// Returns the full-width punctuation mapping for U+FF61..U+FF65, or 0
/// otherwise. Provided as a separate accessor because `text_width` and
/// `jp_fold` mix the two ranges differently in their walkers.
std::uint32_t half_to_full_punctuation(std::uint32_t cp) noexcept;

/// Returns the dakuten form of a full-width katakana base, or 0 if the base
/// cannot take dakuten. This explicit lookup avoids assuming Unicode code
/// point adjacency: small tsu between the sa and ta rows breaks that pattern.
std::uint32_t voiced_full_from_base(std::uint32_t full_base) noexcept;

/// Returns the handakuten form of a full-width katakana base, or 0 if the
/// base cannot take handakuten.
std::uint32_t semi_voiced_full_from_base(std::uint32_t full_base) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JP_KANA_TABLE_H_
