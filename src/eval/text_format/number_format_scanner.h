//
// Internal header -- do not include outside `src/eval/text_format/`.
//
// Scanning primitives shared between the format-string tokenizer and the
// section splitter / classifier. All helpers are pure / `noexcept` and
// operate on raw `std::string_view` slices of the user-supplied format.
//
// The token / section types these helpers feed are declared in
// `number_format_types.h`; the consumer translation units are
// `number_format_tokenizer.cpp` and `number_format_section.cpp`.

#ifndef FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SCANNER_H_
#define FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SCANNER_H_

#include <cstddef>
#include <string_view>

#include "eval/text_format/number_format_types.h"

namespace formulon {
namespace text_format {
namespace number_format_detail {

// Returns true iff `c` is one of `yYmMdDhHsS`. Used to detect date-family
// tokens; AM/PM handled separately because they span multiple characters.
bool is_date_letter(char c) noexcept;

// Parses a run of `letter` characters (case-insensitive for the given
// target). Advances `*i` past the run. Returns the run length (>= 1 given
// the caller already matched at least one character).
std::size_t scan_run(std::string_view fmt, std::size_t& i, char letter) noexcept;

// Returns true if `body` is a color qualifier: either a named color or the
// indexed `色N` form with N in 1..56.
//
// A format string is interpreted in the UI locale, so the accepted spellings
// are the localized ones. Under the ja-JP profile that means `黒`, `青`, `水`,
// `緑`, `紫`, `赤`, `白`, `黄` and `色N` — the English `Red` / `ColorN`
// spellings are *not* accepted and fall through to the invalid-bracket path,
// exactly as Excel rejects `=TEXT(5,"[Red]0.00")` with #VALUE!. The English
// names OOXML stores in `numFmtCode` are a separate, locale-independent
// surface matched where stored formats are read.
//
// The name is a prefix: anything after it in the same bracket is ignored, so
// `[赤色]` and `[色1x]` are both accepted. Leading whitespace is not.
// Excel discards the color itself inside TEXT, so a match is treated the same
// as `[$...]` — inert.
bool is_color_specifier(std::string_view body) noexcept;

// Detects an Excel conditional-section directive of the form `[op N]`,
// where `op` is one of `>`, `>=`, `<`, `<=`, `=`, `<>` and `N` is a literal
// double. Leading whitespace is tolerated (Excel accepts `[> 100]`).
//
// Returns:
//   * 1 on a successful parse: writes the operator and value to `*out_op` /
//     `*out_value`.
//   * -1 if the body looks like a predicate (starts with one of the six
//     operator forms) but the numeric tail fails to parse fully. The caller
//     should treat this as an invalid bracket so `apply_format` surfaces
//     `#VALUE!` exactly as the legacy fallback path did.
//   * 0 if the body does not look like a predicate at all (caller should
//     continue trying other bracket interpretations).
int parse_cond_directive(std::string_view body, CondOp* out_op, double* out_value) noexcept;

// Detects `[DBNumN]` (body length exactly 6 bytes after stripping brackets).
// Matching is case-insensitive: `[DBNum1]`, `[dbnum2]`, `[DbNum3]` all parse.
// Returns the directive index (1, 2, or 3) on a hit, otherwise 0.
int parse_dbnum_directive(std::string_view body) noexcept;

// Returns true if `tok` is a date-family token (including elapsed brackets).
bool is_date_tok(Tok t) noexcept;

}  // namespace number_format_detail
}  // namespace text_format
}  // namespace formulon

#endif  // FORMULON_EVAL_TEXT_FORMAT_NUMBER_FORMAT_SCANNER_H_
