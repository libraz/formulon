// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
//
// Caveat for the D-function header path: empirically Mac Excel does NOT
// fold half-width katakana (U+FF61..U+FF9D, plus the standalone voicing
// marks U+FF9E / U+FF9F) when resolving the `field` argument or a
// criteria-block header against the database header row, even though it
// does fold them in COUNTIF / SUMIF cell-vs-criterion comparisons (see
// `tests/oracle/cases/dfunc_kana_folding_probes.yaml`, in particular
// `dsum_criteria_header_halfwidth_vs_fullwidth_db_header` and
// `dsum_field_arg_halfwidth_vs_fullwidth_header`). D-function callers
// pass `fold_halfwidth_kana = false` to suppress the entire half-width
// katakana branch (including the voicing-mark composition path, so e.g.
// `ｶﾞ` stays as the two-codepoint sequence FF76 FF9E and does not
// compose to ガ). Criteria and lookup callers keep half-width folding on.
// Do not "fix" this asymmetry without re-running the oracle against
// Mac Excel 365 — it is a real product behaviour, not a bug.

#ifndef FORMULON_EVAL_JP_FOLD_H_
#define FORMULON_EVAL_JP_FOLD_H_

#include <string>
#include <string_view>

#include "value.h"

namespace formulon {
namespace eval {

/// Returns `input` rewritten to its Mac Excel ja-JP canonical form for
/// criterion-equality comparison. See header comment for the full mapping.
/// The result is always valid UTF-8 and never contains the half-width
/// voicing marks U+FF9E / U+FF9F when they could be composed with a
/// preceding consonant (and `fold_halfwidth_kana == true`).
///
/// `fold_fullwidth_digits` controls whether U+FF10..U+FF19 are folded to
/// half-width digits. Pass `true` (the default) for COUNTIF / SUMIF /
/// AVERAGEIF parity; pass `false` for the lookup family (MATCH / VLOOKUP /
/// HLOOKUP / XLOOKUP / XMATCH) where Mac Excel deliberately leaves
/// full-width digits unfolded.
///
/// `fold_halfwidth_kana` controls whether half-width katakana
/// U+FF61..U+FF9D and the related standalone voicing marks U+FF9E /
/// U+FF9F are folded. Pass `true` (the default) for COUNTIF / SUMIF /
/// AVERAGEIF and the lookup family; pass `false` for D-function
/// header resolution (DSUM / DCOUNT / DGET / etc.) where Mac Excel
/// deliberately leaves half-width katakana — and the half-width
/// voicing-mark composition — unfolded. When `false`, no codepoint in
/// U+FF61..U+FF9F is rewritten and `ｶﾞ` (FF76 FF9E) remains as the
/// two-codepoint sequence rather than composing to ガ.
std::string fold_jp_text(std::string_view input, bool fold_fullwidth_digits = true, bool fold_halfwidth_kana = true);

/// Convenience wrapper that composes `fold_jp_text` with
/// `strings::to_ascii_lower` in a single pass-allocation. Used by every
/// case-insensitive lookup-family text comparison (VLOOKUP / HLOOKUP /
/// MATCH / XLOOKUP / XMATCH) so the kana-fold + ASCII-lowercase pair
/// stays in lockstep across `lookups/classic.cpp` and `lookups/xlookup.cpp`.
///
/// `fold_fullwidth_digits` matches the underlying `fold_jp_text` flag —
/// pass `false` for lookup-family parity with Mac Excel 365 (full-width
/// digits stay unfolded), pass `true` for COUNTIF / SUMIF parity.
std::string fold_and_lower(std::string_view input, bool fold_fullwidth_digits = false);

/// Composes a half-width voicing mark (U+FF9E ﾞ / U+FF9F ﾟ) onto the
/// preceding half-width katakana base, emitting the full-width voiced /
/// semi-voiced katakana (`ｶﾞ` -> `ガ`, `ﾊﾟ` -> `パ`). Every other
/// codepoint — including plain half-width katakana that is NOT followed
/// by a voicing mark, hiragana, full-width katakana, and ASCII — passes
/// through byte-for-byte unchanged.
///
/// This is the minimal normalisation Windows Excel 365 applies to the
/// lookup family: a half-width base + standalone voicing mark is a
/// malformed encoding that Excel composes before comparing, so
/// `XLOOKUP("ｶﾞ", ...)` matches a cell holding `ガ`. Unlike
/// `fold_jp_text`, it does NOT fold plain half/full width or
/// hiragana<->katakana, matching the Windows-Excel asymmetry pinned by
/// `tests/oracle/variants/win-365-ja_JP/golden/lookup_kana_folding_probes`.
std::string compose_jp_halfwidth_voicing(std::string_view input);

// ---------------------------------------------------------------------------
// High-level fold-aware comparison helpers
// ---------------------------------------------------------------------------
//
// Consolidates the ad-hoc `fold_jp_text(a) op fold_jp_text(b)` patterns
// scattered across `eval/lookups/*.cpp`, `eval/groupby_pivotby_lazy.cpp`,
// and `eval/database_lazy.cpp`. Each call site previously inlined the
// double-fold + compare, picking one of:
//   * groupby      : `fold_jp_text(a) == fold_jp_text(b)` (default flags,
//                    case-sensitive byte compare)
//   * lookups      : `case_insensitive_compare(fold_jp_text(a, false),
//                                              fold_jp_text(b, false))`
//                    (lookups disable full-width digit folding; D-functions
//                    additionally disable half-width katakana folding)
//   * database     : `case_insensitive_eq(fold_jp_text(a, true, false),
//                                          fold_jp_text(b, true, false))`
//
// These helpers cover all three shapes; call sites pick the flag triple
// (case_insensitive / fold_fullwidth_digits / fold_halfwidth_kana) that
// matches their Mac Excel oracle row.

/// Options controlling the high-level fold-aware comparisons below.
///
/// Default values match the COUNTIF / SUMIF criterion-equality contract:
/// case-insensitive ASCII compare with full-width digits and half-width
/// katakana both folded. Lookup callers (MATCH / VLOOKUP / HLOOKUP /
/// XLOOKUP / XMATCH) pass `fold_fullwidth_digits = false`; D-function
/// header callers (DSUM / DCOUNT / DGET / etc.) pass both
/// `fold_fullwidth_digits = true, fold_halfwidth_kana = false`. The
/// GROUPBY / PIVOTBY key-equality path passes `case_insensitive = false`.
struct FoldCompareOptions {
  bool case_insensitive = true;
  bool fold_fullwidth_digits = true;
  bool fold_halfwidth_kana = true;
};

/// Returns true iff `a` and `b` are byte-equal after each is rewritten via
/// `fold_jp_text` under `opts`. When `opts.case_insensitive` is true the
/// post-fold compare uses `strings::case_insensitive_eq` (ASCII letters
/// only); otherwise the compare is exact-byte.
///
/// Both inputs are assumed to be valid UTF-8 — the underlying decoder
/// pass-through clamps malformed bytes but the equality verdict for such
/// inputs is not specified.
bool equal_folded(std::string_view a, std::string_view b, FoldCompareOptions opts = {});

/// Lexicographic compare of `a` and `b` after each is rewritten via
/// `fold_jp_text` under `opts`. Returns a negative value when `a < b`,
/// zero when `a == b`, and a positive value when `a > b`. When
/// `opts.case_insensitive` is true the compare uses
/// `strings::case_insensitive_compare`; otherwise it is a byte-wise
/// `std::string_view::compare`.
int compare_folded(std::string_view a, std::string_view b, FoldCompareOptions opts = {});

/// Value-level Text equality with JP folding. Returns true iff `a` and
/// `b` are Text values that fold-compare equal under `opts`. For any
/// other kind pair (Number/Number, Bool/Bool, Error/Error, Blank/Blank)
/// falls back to direct payload equality. Cross-kind pairs are never
/// equal. Array/Ref/Lambda payloads (which are not produced by cell
/// reads) are conservatively reported as not equal.
///
/// This matches the GROUPBY / UNIQUE-with-fold semantics in
/// `groupby_pivotby_lazy.cpp::group_cell_equal`. Number compare is
/// IEEE-754 `==` (so `+0.0 == -0.0` here, mirroring `group_cell_equal`).
bool value_equal_folded_text(const Value& a, const Value& b, FoldCompareOptions opts = {});

/// Value-level Text-aware ordered compare with JP folding. Returns the
/// sign of (a - b) when both are the same kind:
///   * Number: numeric compare.
///   * Text:   `compare_folded` under `opts`.
///   * Bool:   FALSE < TRUE.
///   * Error:  compare by `static_cast<int>` of the error code.
///   * Blank:  always 0.
/// For cross-kind pairs the caller is expected to apply its own bucket
/// ordering first; this helper returns 0 to signal "kinds are equivalent
/// for ordering purposes — supply your own tiebreaker".
int value_compare_folded_text(const Value& a, const Value& b, FoldCompareOptions opts = {});

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_JP_FOLD_H_
