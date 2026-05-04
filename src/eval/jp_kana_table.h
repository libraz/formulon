// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Shared half-width-katakana to full-width-katakana base mapping. Both
// `jp_fold` (used by lookup-key folding) and `text_width` (used by JIS /
// DBCS / ASC) consume this table; previously each TU carried its own
// copy of the same 56-entry block.
//
// Only the half-to-full *base* table is shared. The full-to-half side is
// width-conversion-specific (entries decompose voiced forms into a base
// codepoint plus U+FF9E / U+FF9F so the half-width output is two bytes)
// and stays inside `text_width.cpp`. Voiced / semi-voiced composition
// lives in each consumer because the input shape differs:
//   * `jp_fold` peeks the next codepoint of a folding stream;
//   * `text_width` operates on a pre-walked codepoint sequence.

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

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JP_KANA_TABLE_H_
