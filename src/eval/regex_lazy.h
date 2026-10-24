//
// Lazy impls for the Excel REGEX function family backed by PCRE2:
//   * REGEXTEST(text, pattern, [case_sensitivity]) -> bool
//   * REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity]) ->
//     scalar text or array of matches / capture groups
//   * REGEXREPLACE(text, pattern, replacement, [occurrence],
//     [case_sensitivity]) -> text
//
// They ride the lazy seam because their `text` argument must reach the
// impl with its (rows, cols) shape preserved (matches Mac Excel: an
// array-valued `text` hands back an array-shaped result, broadcasting
// cellwise). The result shape follows the argument's *value*, so a range,
// an array literal, an arithmetic broadcast and a call returning an array
// (UNIQUE / SORT / FILTER ...) all broadcast alike.
//
// All three evaluate their arguments in signature order, so when more
// than one argument is in error the leftmost one surfaces.
//
// Compile + match runs through a single shared kernel in
// `regex_lazy.cpp`; see that file for the option set, error mapping,
// and resource limits applied to every regex evaluation.

#ifndef FORMULON_EVAL_REGEX_LAZY_H_
#define FORMULON_EVAL_REGEX_LAZY_H_

#include <cstdint>
#include <memory>
#include <string_view>

#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {

namespace parser {
class AstNode;
}  // namespace parser

namespace eval {

class EvalContext;
class FunctionRegistry;

/// `REGEXTEST(text, pattern, [case_sensitivity=0])` -> Bool. Returns TRUE
/// when `pattern` matches anywhere in `text`. `case_sensitivity` must
/// coerce to {0, 1}; any other value yields `#VALUE!`. The convention
/// matches Mac Excel 365 / MS docs: `0`/FALSE (default) is
/// case-sensitive, `1`/TRUE is case-insensitive. Exhausting the match
/// limit is `#VALUE!`; any other PCRE2 resource error (depth / heap limit)
/// returns FALSE.
Value eval_regextest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx);

/// `REGEXEXTRACT(text, pattern, [return_mode=0], [case_sensitivity=0])`.
/// `return_mode` selects the result shape:
///   * 0 -> scalar text of the first whole match.
///   * 1 -> row array (1 x N) of all whole matches.
///   * 2 -> row array (1 x G) of capture groups from the FIRST match.
/// Any other mode (3 included) is `#VALUE!`.
/// Mode 2 yields `#VALUE!` when the pattern has no capture groups
/// AND a match occurred (no match still yields `#N/A`). Exhausting the
/// match limit is `#VALUE!`; any other resource exhaustion (depth / heap
/// limit, retained-match budget) is `#CALC!`.
Value eval_regexextract_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

/// `REGEXREPLACE(text, pattern, replacement, [occurrence=0],
/// [case_sensitivity=0])`. `occurrence == 0` replaces every match;
/// `occurrence == N > 0` replaces only the N-th match from the start and
/// `N < 0` only the |N|-th from the end. When |N| exceeds the number of
/// matches the original `text` is returned unchanged. Exhausting the match
/// limit, or a result past the 32,767-character text cap, is `#VALUE!`;
/// any other resource exhaustion is `#CALC!`.
///
/// `replacement` has a `$` layer (`$n`, `${n}`, `${name}`, `$name`,
/// `$<name>`, `$&`, `` $` ``, `$'`, `$_`, `$$`) and a backslash layer
/// (group references, character escapes, `\u` / `\l`, `\U`..`\E` /
/// `\L`..`\E` case runs, `\Q`..`\E`); the full grammar is listed in
/// `regex_lazy.cpp`. A bare `$`, an unknown escape or a reference to a
/// missing group is `#VALUE!` once a match is expanded.
Value eval_regexreplace_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx);

/// A case-sensitive PCRE2 pattern compiled once and searched against many
/// subjects; XMATCH / XLOOKUP match_mode 3 uses it so the REGEX* family stays
/// the only place patterns are compiled.
class RegexSearch {
 public:
  /// Compiles `pattern`. An oversized or invalid pattern leaves `ok()` false;
  /// an empty pattern compiles and matches every subject.
  explicit RegexSearch(std::string_view pattern);
  ~RegexSearch();
  RegexSearch(const RegexSearch&) = delete;
  RegexSearch& operator=(const RegexSearch&) = delete;

  bool ok() const noexcept;

  /// Whether the pattern matches anywhere in `subject`, with REGEXTEST's
  /// outcome: exhausting the match limit is `#VALUE!`, any other resource
  /// exhaustion counts as no match.
  Expected<bool, ErrorCode> matches(std::string_view subject) const;

 private:
  struct Impl;
  std::unique_ptr<Impl> impl_;
};

/// Running count of `pcre2_compile` calls issued by the REGEX* family in
/// this process.
///
/// One REGEX* call compiles its (pattern, case-sensitivity) pair exactly
/// once and reuses the compiled program for every subject cell, so a
/// broadcast over an N-cell range must not move this counter by N. The
/// counter exists so that contract is observable — pattern compilation costs
/// an order of magnitude more than a match against a short cell string, and
/// a regression there is invisible in the results.
std::uint64_t regex_compile_count() noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REGEX_LAZY_H_
