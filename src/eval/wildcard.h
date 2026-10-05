//
// Excel's wildcard dialect: `*` matches any run of codepoints, `?` matches
// exactly one codepoint, and `~` escapes the next metacharacter. Matching is
// ASCII case-insensitive and walks UTF-8 codepoint-by-codepoint. Shared by
// the criteria matcher, SEARCH/SEARCHB and the exact-match lookups.

#ifndef FORMULON_EVAL_WILDCARD_H_
#define FORMULON_EVAL_WILDCARD_H_

#include <cstddef>
#include <string_view>

namespace formulon {
namespace eval {

/// Byte-level wildcard matcher. `pattern` may contain unescaped `*` / `?`
/// and `~`-escaped literals; `text` is matched verbatim. Shared between
/// `matches_criterion` (for `COUNTIF`/`SUMIF`/`AVERAGEIF` equality paths)
/// and scalar lookups such as `MATCH(..., 0)` that honour the same
/// DOS-style wildcard dialect.
///
/// Iterative backtrack (no recursion) keeps stack depth constant. `?`
/// matches exactly one code point (1-4 UTF-8 bytes), not one byte.
bool wildcard_match(std::string_view pattern, std::string_view text);

/// Scans `rhs` for an unescaped `*` or `?`. A leading `~` escapes the next
/// byte, so `"~*"` contains no wildcard but `"*~*"` does. Returns `true` on
/// the first unescaped metacharacter encountered.
bool scan_has_wildcard(std::string_view rhs);

/// Returns the byte offset where `pattern` first matches a prefix of
/// `text` starting at that offset (wildcards: `*`, `?`, `~` for escape).
/// Returns `std::string_view::npos` when no position in `text` begins a
/// match. The match is anchored at its start but not at its end — it
/// succeeds as soon as `pattern` is consumed, regardless of any unmatched
/// suffix in `text`. This mirrors the SEARCH contract where the pattern
/// need only match somewhere inside the haystack.
///
/// Callers that need case-insensitive matching must lower-case both inputs
/// first (the matcher is byte-exact). UTF-16 offset mapping is likewise the
/// caller's responsibility; this helper only reports a byte offset into
/// `text`.
std::size_t wildcard_find(std::string_view pattern, std::string_view text);

/// Same contract as `wildcard_find`, but `?` matches only a single
/// SBCS-byte codepoint under the ja-JP DBCS rule (`byte_count_jajp(cp)
/// == 1`): ASCII and half-width katakana match, every other codepoint
/// (kanji, hiragana, full-width punctuation, emoji) refuses to match.
/// `*`, `~?`, `~*`, and literal characters retain their normal
/// semantics. Used exclusively by SEARCHB to mirror Mac Excel 365's
/// byte-oriented wildcard dialect.
///
/// The returned offset is always aligned to a UTF-8 codepoint boundary in
/// `text`, so callers translating the offset to a DBCS position (via
/// `build_dbcs_char_map`) can use direct equality on `byte_offset`.
std::size_t wildcard_find_dbcs(std::string_view pattern, std::string_view text);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_WILDCARD_H_
