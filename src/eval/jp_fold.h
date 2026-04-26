// Copyright 2026 libraz. Licensed under the MIT License.
//
// ja-JP canonical text folding for Mac Excel COUNTIF / SUMIF / AVERAGEIF
// equality and the lookup family (VLOOKUP / HLOOKUP / MATCH / XLOOKUP /
// XMATCH). Mac Excel collapses several Japanese character variants to a
// single canonical form before comparing criterion / lookup text against
// cell text:
//
//   * Hiragana (U+3041..U+3096) folds to Katakana (+0x60 offset).
//   * Half-width katakana (U+FF61..U+FF9D) folds to full-width katakana,
//     composing the trailing voicing (U+FF9E ﾞ) and semi-voicing
//     (U+FF9F ﾟ) marks where applicable (e.g. ｶﾞ -> ガ).
//   * Full-width ASCII (U+FF01..U+FF5E) folds to half-width ASCII (-0xFEE0).
//   * Ideographic space U+3000 folds to ASCII space U+0020.
//
// Other code points pass through unchanged. The fold is purely a textual
// normalisation; it does not interpret or strip any wildcard semantics.
// The criterion-matching code in `criteria.cpp` calls this BEFORE the
// case-insensitive ASCII compare, so e.g. `＊` -> `*` is then matched as
// the literal byte `*` rather than as a wildcard.
//
// Caveat for the lookup family: empirically Mac Excel does NOT fold
// full-width DIGITS U+FF10..U+FF19 in MATCH / VLOOKUP / HLOOKUP / XLOOKUP /
// XMATCH text-equality, even though it does fold them in COUNTIF / SUMIF
// (see `tests/oracle/cases/lookup_kana_folding_probes.yaml` vs
// `tests/oracle/cases/countif_kana_folding_probes.yaml`). Lookup callers
// pass `fold_fullwidth_digits = false` to suppress that single sub-range;
// criteria callers (the default) keep folding digits.

#ifndef FORMULON_EVAL_JP_FOLD_H_
#define FORMULON_EVAL_JP_FOLD_H_

#include <string>
#include <string_view>

namespace formulon {
namespace eval {

/// Returns `input` rewritten to its Mac Excel ja-JP canonical form for
/// criterion-equality comparison. See header comment for the full mapping.
/// The result is always valid UTF-8 and never contains the half-width
/// voicing marks U+FF9E / U+FF9F when they could be composed with a
/// preceding consonant.
///
/// `fold_fullwidth_digits` controls whether U+FF10..U+FF19 are folded to
/// half-width digits. Pass `true` (the default) for COUNTIF / SUMIF /
/// AVERAGEIF parity; pass `false` for the lookup family (MATCH / VLOOKUP /
/// HLOOKUP / XLOOKUP / XMATCH) where Mac Excel deliberately leaves
/// full-width digits unfolded.
std::string fold_jp_text(std::string_view input, bool fold_fullwidth_digits = true);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JP_FOLD_H_
