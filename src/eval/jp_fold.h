// Copyright 2026 libraz. Licensed under the MIT License.
//
// ja-JP canonical text folding for Mac Excel COUNTIF / SUMIF / AVERAGEIF
// equality. Mac Excel collapses several Japanese character variants to a
// single canonical form before comparing criterion text against cell text:
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
std::string fold_jp_text(std::string_view input);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JP_FOLD_H_
