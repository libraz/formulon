//
// Excel's wildcard dialect: `*` matches any run of codepoints, `?` matches
// exactly one codepoint, and `~` escapes the next metacharacter. Matching is
// ASCII case-insensitive and walks UTF-8 codepoint-by-codepoint. Shared by
// the criteria matcher and the exact-match lookups.

#ifndef FORMULON_EVAL_WILDCARD_H_
#define FORMULON_EVAL_WILDCARD_H_

#include <cstddef>
#include <string>
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

/// Removes the escapes of a pattern without wildcards: `~X` becomes `X` for
/// any `X`, and a trailing lone `~` stays. Used for the literal comparison of
/// a criterion or lookup value whose only metacharacters are escaped.
std::string unescape_wildcard_literal(std::string_view pattern);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_WILDCARD_H_
